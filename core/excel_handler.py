"""Excel file handling for VINAFCO Manifest Cleaner.

This module contains functions for reading Excel manifest files
and writing formatted output files.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

try:
    from xlsxwriter.utility import xl_col_to_name
    XLSXWRITER_AVAILABLE = True
except ImportError:
    xl_col_to_name = None
    XLSXWRITER_AVAILABLE = False

from core.parsers import (
    CONT_NO_PATTERN,
    parse_shipper_line_info,
    parse_weight_line,
    is_container_line,
    is_party_marker_line,
)
from core.processors import (
    determine_fe,
    determine_cargo_type,
    clean_party_name,
    get_iso_size,
    refine_description,
)
from utils.helpers import get_cell_value_as_str
from utils.normalization import normalize_text


logger = logging.getLogger(__name__)


class VinafcoExcelHandler:
    """Handles Excel file reading and writing for VINAFCO manifests."""
    
    def __init__(self, config: Dict[str, Any]):
        """Initialize handler with configuration.
        
        Args:
            config: Dictionary containing Excel handling configuration
        """
        self.config = config
        
        # Excel settings
        self.peek_rows = config.get('peek_rows', 30)
        self.engine_xlsx = config.get('excel_engine_xlsx', 'openpyxl')
        self.engine_xls = config.get('excel_engine_xls', 'xlrd')
        self.output_sheet = config.get('output_sheet_name', 'Processed_Manifest')
        self.max_col_width = config.get('max_column_width', 50)
        
        # Formatting colors
        self.seal_invalid_color = config.get('seal_invalid_bg_color', '#FFFF00')
        self.empty_row_color = config.get('empty_container_bg_color', '#DAEEF3')
        
        # Column configuration
        self.target_sheet = config.get('target_sheet_name', 'MANIFEST')
        self.party_marker_len = config.get('party_marker_len', 3)
        self.operator_soc = config.get('operator_soc', 'SVF')
        self.operator_coc = config.get('operator_coc', 'VFC')
        
        # Input column names
        self.col_party_raw = config.get('col_party_raw', '(S) SHIPPER,(C) CONSIGNEE,(N)NOTIFY PARTY')
        self.col_bill_no = config.get('col_bill_no', 'Bill No')
        self.col_container_no = config.get('col_container_no', 'CONTAINER No.')
        self.col_seal = config.get('col_seal', 'SEAL')
        self.col_type = config.get('col_type', 'TYPE')
        self.col_comm = config.get('col_comm', 'COM.M')
        self.col_gw = config.get('col_gw', 'GW')
        self.col_remarks = config.get('col_remarks', 'REMARKS')
        
        # Expected column identifiers (lowercase for matching)
        self.expected_party_lower = config.get('expected_party_lower', '(s) shipper,(c) consignee,(n)notify party')
        self.expected_bill_lower = config.get('expected_bill_lower', 'bill no')
        
        # Output column mapping
        self.output_mapping = config.get('output_column_mapping', {})
        self.desired_columns = config.get('desired_output_columns', [])
        
        # Output column names
        self.seal_col_output = config.get('seal_column_output_name', 'Số niêm chì')
        self.consignee_col_output = config.get('consignee_column_output_name', 'Chủ hàng')
        self.fe_col_output = config.get('fe_column_output_name', 'F/E')
        self.goods_col_output = config.get('goods_column_output_name', 'Hàng hóa')
        self.temp_count_col = config.get('temp_count_column_name', 'Số lượng_temp')
    
    def read_excel(self, file_path: Path, progress_callback=None) -> Tuple[pd.DataFrame, int]:
        """Read and parse VINAFCO manifest Excel file.
        
        Args:
            file_path: Path to Excel file
            progress_callback: Optional callback(value, message) for progress updates
            
        Returns:
            Tuple of (DataFrame, header_row_index)
            
        Raises:
            FileNotFoundError: If file doesn't exist
            ValueError: If header not found or columns missing
        """
        if progress_callback:
            progress_callback(0, "Starting to read Excel file...")
        
        # Convert to Path if string
        if isinstance(file_path, str):
            file_path = Path(file_path)
        
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        # Determine engine based on actual file content (magic bytes),
        # not just extension - some .xls files are actually XLSX format
        try:
            with open(file_path, 'rb') as f:
                magic = f.read(4)
            is_xlsx_format = (magic == b'PK\x03\x04')  # ZIP/XLSX signature
        except Exception:
            is_xlsx_format = file_path.suffix.lower() in ('.xlsx', '.xlsm')
        
        if is_xlsx_format:
            engine = self.engine_xlsx  # openpyxl
            logger.info(f"Detected XLSX format (ZIP signature) for: {file_path.name}")
        else:
            engine = self.engine_xls   # xlrd
            logger.info(f"Detected XLS format (OLE2) for: {file_path.name}")
        
        is_xlsx = is_xlsx_format  # used by fallback logic below
        
        # Peek at first rows to find header
        try:
            df_peek = pd.read_excel(
                file_path,
                sheet_name=self.target_sheet,
                header=None,
                nrows=self.peek_rows,
                keep_default_na=False,
                engine=engine
            )
        except Exception as e:
            logger.warning(f"Failed with engine {engine}, sheet '{self.target_sheet}': {e}")
            # Fallback: try with correct engine but first sheet (sheet_name=0)
            fallback_engine = 'openpyxl' if is_xlsx else None
            try:
                df_peek = pd.read_excel(
                    file_path,
                    sheet_name=0,
                    header=None,
                    nrows=self.peek_rows,
                    keep_default_na=False,
                    engine=fallback_engine
                )
                logger.info("Fallback: reading first sheet succeeded")
            except Exception as e2:
                logger.error(f"Fallback also failed: {e2}")
                raise ValueError(
                    f"Cannot read Excel file. Original error: {e}"
                ) from e2
        
        if progress_callback:
            progress_callback(5, "Finding header row...")
        
        # Find header row
        header_row_idx = -1
        for idx, row in df_peek.iterrows():
            row_values = [str(val).strip().lower() for val in row.tolist()]
            if len(row_values) > 1:
                if row_values[0] == self.expected_party_lower and row_values[1] == self.expected_bill_lower:
                    header_row_idx = idx
                    logger.info(f"Found header at row index {idx}")
                    break
        
        if header_row_idx == -1:
            raise ValueError(
                f"Could not find header row in first {self.peek_rows} rows. "
                f"Expected: '{self.expected_party_lower}' and '{self.expected_bill_lower}'"
            )
        
        if progress_callback:
            progress_callback(10, "Reading full data...")
        
        # Read full data with header - try named sheet first, then first sheet
        try:
            df = pd.read_excel(
                file_path,
                sheet_name=self.target_sheet,
                header=header_row_idx,
                keep_default_na=False,
                engine=engine
            )
        except Exception as e:
            logger.warning(f"Full read with sheet '{self.target_sheet}' failed: {e}, trying first sheet")
            fallback_engine = 'openpyxl' if is_xlsx else None
            df = pd.read_excel(
                file_path,
                sheet_name=0,
                header=header_row_idx,
                keep_default_na=False,
                engine=fallback_engine
            )
        
        # Normalize column names
        df.columns = [' '.join(re.sub(r'[\n\r\t]+', ' ', str(col)).split()).strip() 
                      for col in df.columns]
        
        logger.info(f"Columns found: {df.columns.tolist()}")
        
        # Verify required columns (with fuzzy matching for variations)
        required_cols = [
            self.col_party_raw, self.col_bill_no, self.col_container_no,
            self.col_seal, self.col_type, self.col_comm, self.col_gw, self.col_remarks
        ]
        missing = [col for col in required_cols if col not in df.columns]
        
        # Try fuzzy matching for missing columns
        # e.g., 'GW' matches 'GW (MT)', 'CONTAINER No.' matches 'CONTAINER No. 1'
        if missing:
            rename_map = {}
            still_missing = []
            for expected_col in missing:
                matched = False
                for actual_col in df.columns:
                    # Check if actual column starts with expected name
                    if actual_col.upper().startswith(expected_col.upper()):
                        rename_map[actual_col] = expected_col
                        logger.info(f"Fuzzy column match: '{actual_col}' -> '{expected_col}'")
                        matched = True
                        break
                if not matched:
                    still_missing.append(expected_col)
            
            if rename_map:
                df.rename(columns=rename_map, inplace=True)
            
            if still_missing:
                raise ValueError(f"Missing required columns: {still_missing}")
        
        return df, header_row_idx
    
    def parse_manifest_data(
        self,
        df: pd.DataFrame,
        header_row_idx: int,
        progress_callback=None
    ) -> List[Dict[str, Any]]:
        """Parse manifest DataFrame into structured container records.
        
        Args:
            df: Raw DataFrame from Excel
            header_row_idx: Index of header row (for Excel row reference)
            progress_callback: Optional callback for progress updates
            
        Returns:
            List of dictionaries containing container data
        """
        if progress_callback:
            progress_callback(15, "Parsing data...")
        
        intermediate_data = []
        
        # Tracking variables
        current_bl_group_id = None
        current_is_soc = False
        current_bill_no = None
        active_shipper = []
        active_consignee = []
        active_notify = []
        active_containers = []
        collecting_for_party = None
        bl_group_counter = 0
        
        total_rows = len(df)
        header_offset = header_row_idx + 1
        
        def finalize_group():
            """Process and store the current B/L group."""
            nonlocal current_bl_group_id
            
            if active_containers and current_bl_group_id:
                shipper_text = "\n".join(active_shipper).strip()
                consignee_text = "\n".join(active_consignee).strip()
                notify_text = "\n".join(active_notify).strip()
                
                cleaned_consignee = clean_party_name(consignee_text, self.party_marker_len)
                
                # Handle "SAME AS CONSIGNEE" for notify
                if "SAME AS CONSIGNEE" in notify_text.upper() or "NHU NGUOI NHAN HANG" in notify_text.upper():
                    notify_text = consignee_text
                
                for cont in active_containers:
                    record = {
                        "Master_BL_Group_ID": current_bl_group_id,
                        "Shipper": shipper_text,
                        "Consignee": consignee_text,
                        "Notify": notify_text,
                        "Cleaned_Consignee_Name_For_Output": cleaned_consignee,
                        **cont
                    }
                    intermediate_data.append(record)
            
            # Reset for next group
            active_shipper.clear()
            active_consignee.clear()
            active_notify.clear()
            active_containers.clear()
        
        # Process each row
        for idx, row in df.iterrows():
            excel_row = idx + header_offset + 1
            
            party_raw = get_cell_value_as_str(row.get(self.col_party_raw, ''))
            bill_no = get_cell_value_as_str(row.get(self.col_bill_no, ''))
            container_no = get_cell_value_as_str(row.get(self.col_container_no, ''))
            
            is_party_line = is_party_marker_line(party_raw)
            is_cont_line = is_container_line(container_no)
            is_detail_line = bool(party_raw and not is_party_line and not is_cont_line)
            
            # Handle party marker lines
            if is_party_line:
                party_info = parse_shipper_line_info(party_raw, self.party_marker_len)
                
                if party_raw.startswith("(S)"):
                    # New B/L group starts
                    if current_bl_group_id:
                        finalize_group()
                    
                    bl_group_counter += 1
                    current_bl_group_id = f"MASTER_BL_{bl_group_counter}"
                    current_is_soc = party_info["is_soc_from_shipper_line"]
                    current_bill_no = None
                    
                    active_shipper = [party_info["name"]]
                    active_consignee = []
                    active_notify = []
                    active_containers = []
                    collecting_for_party = active_shipper
                    
                elif party_raw.startswith("(C)"):
                    active_consignee = [party_info["name"]]
                    collecting_for_party = active_consignee
                    
                elif party_raw.startswith("(N)"):
                    active_notify = [party_info["name"]]
                    collecting_for_party = active_notify
            
            # Collect party detail lines
            elif collecting_for_party is not None and is_detail_line:
                collecting_for_party.append(party_raw)
            
            # Handle container lines
            if is_cont_line:
                collecting_for_party = None
                
                if bill_no:
                    current_bill_no = bill_no
                
                # Handle orphan containers
                if not current_bl_group_id:
                    bl_group_counter += 1
                    current_bl_group_id = f"ORPHAN_BL_{bl_group_counter}"
                    active_shipper = ["N/A"]
                    active_consignee = ["N/A"]
                    active_notify = ["N/A"]
                    active_containers = []
                    current_is_soc = False
                
                bill_no_to_use = current_bill_no or f"({current_bl_group_id}_no_bill)"
                
                # Extract container data
                seal = get_cell_value_as_str(row.get(self.col_seal, ''))
                size_type = get_cell_value_as_str(row.get(self.col_type, ''))
                commodity = get_cell_value_as_str(row.get(self.col_comm, ''))
                gw_raw = get_cell_value_as_str(row.get(self.col_gw, ''))
                remarks = get_cell_value_as_str(row.get(self.col_remarks, ''))
                
                # Parse weight
                weight_info = parse_weight_line(gw_raw)
                
                # Determine SOC status (Check Remarks and Description)
                remarks_upper = remarks.upper()
                desc_upper = commodity.upper()
                
                # Check for SOC keywords
                is_soc_remark = "S.O.C" in remarks_upper or "SOC" in remarks_upper
                is_soc_desc = bool(re.search(r'\bSOC\b', desc_upper)) or "SOC LADEN" in desc_upper or "SOC EMPTY" in desc_upper
                
                is_soc = current_is_soc or is_soc_remark or is_soc_desc
                operator = self.operator_soc if is_soc else self.operator_coc
                
                # Determine F/E and cargo type
                fe_status = determine_fe(commodity, weight_info, size_type)
                iso_size = get_iso_size(size_type)
                cargo_type = determine_cargo_type(size_type, fe_status)
                
                # Check seal validity and build warnings
                warnings = []
                if seal and not (seal.isdigit() and len(seal) == 6):
                    warnings.append("⚠️ Seal không hợp lệ")
                
                container_data = {
                    "BL_No": bill_no_to_use,
                    "Container_No": container_no,
                    "Seal_No": seal if seal else None,
                    "Size_Type": size_type if size_type else None,
                    "Operator": operator,
                    "ISO_Size": iso_size,
                    "FE_Status": fe_status,
                    "Weight_Value": weight_info["value"] if weight_info else None,
                    "Weight_Unit": weight_info["unit"] if weight_info else None,
                    "Description": normalize_text(commodity),
                    "Cargo_Type": cargo_type,
                    "Input_Remarks": remarks,
                    "Confidence": 100,  # Vinafco has direct column mapping
                    "Warning": " | ".join(warnings) if warnings else "",
                }
                active_containers.append(container_data)
            
            # Update progress
            if progress_callback and total_rows > 0 and idx % max(1, total_rows // 20) == 0:
                progress = 15 + int((idx / total_rows) * 70)
                progress_callback(min(progress, 85), "Processing containers...")
        
        # Finalize last group
        finalize_group()
        
        logger.info(f"Parsed {len(intermediate_data)} container records")
        return intermediate_data
    
    def build_output_dataframe(
        self,
        intermediate_data: List[Dict[str, Any]],
        progress_callback=None
    ) -> pd.DataFrame:
        """Build final output DataFrame from intermediate data.
        
        Args:
            intermediate_data: List of container record dictionaries
            progress_callback: Optional callback for progress updates
            
        Returns:
            Formatted output DataFrame
        """
        if progress_callback:
            progress_callback(85, "Building output table...")
        
        if not intermediate_data:
            logger.warning("No intermediate data to build DataFrame")
            return pd.DataFrame()
        
        df = pd.DataFrame(intermediate_data)
        
        # Filter valid records
        df = df[
            df['Cleaned_Consignee_Name_For_Output'].notna() & 
            (df['Cleaned_Consignee_Name_For_Output'].str.strip() != '')
        ].copy()
        
        if df.empty:
            return pd.DataFrame()
        
        # Calculate container count per B/L group
        df[self.temp_count_col] = df.groupby('Master_BL_Group_ID')['Container_No'].transform('count')
        
        # Select and rename columns
        cols_to_select = [col for col in self.output_mapping.keys() if col in df.columns]
        if self.temp_count_col in df.columns and self.temp_count_col not in cols_to_select:
            cols_to_select.append(self.temp_count_col)
        
        result_df = df[cols_to_select].copy()
        result_df.rename(columns=self.output_mapping, inplace=True)
        
        # Round weight values
        weight_col = self.output_mapping.get('Weight_Value')
        if weight_col and weight_col in result_df.columns:
            result_df[weight_col] = pd.to_numeric(result_df[weight_col], errors='coerce').round(2)
        
        # Add STT column
        result_df.insert(0, 'STT', range(1, len(result_df) + 1))
        
        # Add Seal No 1 column
        seal_1_col = 'Số niêm chì 1'
        if self.seal_col_output in result_df.columns:
            insert_loc = result_df.columns.get_loc(self.consignee_col_output) if self.consignee_col_output in result_df.columns else result_df.columns.get_loc(self.seal_col_output) + 1
            result_df.insert(insert_loc, seal_1_col, result_df[self.seal_col_output])
        
        # Refine descriptions
        fe_col = self.output_mapping.get('FE_Status', self.fe_col_output)
        goods_col = self.output_mapping.get('Description', self.goods_col_output)
        cargo_col = self.output_mapping.get('Cargo_Type', 'Loại hàng')
        
        if all(col in result_df.columns for col in [goods_col, fe_col, cargo_col]):
            result_df = refine_description(result_df, goods_col, fe_col, cargo_col)
        
        # Reorder columns to match desired output
        final_cols = []
        for col in self.desired_columns:
            if col in result_df.columns:
                final_cols.append(col)
            else:
                result_df[col] = pd.NA
                final_cols.append(col)
        
        # Add any remaining columns
        for col in result_df.columns:
            if col not in final_cols:
                final_cols.append(col)
        
        result_df = result_df[final_cols]
        
        if progress_callback:
            progress_callback(95, "Output table ready")
        
        return result_df
    
    def save_excel(
        self,
        df: pd.DataFrame,
        output_path: Path,
        progress_callback=None
    ) -> bool:
        """Save DataFrame to formatted Excel file.
        
        Args:
            df: DataFrame to save
            output_path: Output file path
            progress_callback: Optional callback for progress updates
            
        Returns:
            True if successful
        """
        if progress_callback:
            progress_callback(95, "Saving file...")
        
        if not XLSXWRITER_AVAILABLE:
            logger.warning("xlsxwriter not available, saving without formatting")
            df.to_excel(output_path, index=False, sheet_name=self.output_sheet)
            return True
        
        try:
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name=self.output_sheet)
                
                workbook = writer.book
                worksheet = writer.sheets[self.output_sheet]
                
                # Define formats
                yellow_format = workbook.add_format({'bg_color': self.seal_invalid_color})
                blue_format = workbook.add_format({'bg_color': self.empty_row_color})
                
                # Seal validation formatting
                if self.seal_col_output in df.columns and xl_col_to_name:
                    seal_idx = df.columns.get_loc(self.seal_col_output)
                    seal_letter = xl_col_to_name(seal_idx)
                    last_row = len(df) + 1
                    
                    if last_row > 1:
                        formula = f'=AND(NOT(ISBLANK({seal_letter}2)), NOT(AND(ISNUMBER(--TRIM({seal_letter}2)), OR(LEN(TRIM({seal_letter}2))=6, LEN(TRIM({seal_letter}2))=7))))'
                        worksheet.conditional_format(
                            f'{seal_letter}2:{seal_letter}{last_row}',
                            {'type': 'formula', 'criteria': formula, 'format': yellow_format}
                        )
                
                # Empty row highlighting
                fe_col = self.output_mapping.get('FE_Status', self.fe_col_output)
                if fe_col in df.columns and xl_col_to_name:
                    fe_idx = df.columns.get_loc(fe_col)
                    fe_letter = xl_col_to_name(fe_idx)
                    last_row = len(df) + 1
                    
                    if last_row > 1:
                        num_cols = len(df.columns)
                        first_letter = xl_col_to_name(0)
                        last_letter = xl_col_to_name(num_cols - 1)
                        
                        formula = f'TRIM(${fe_letter}2)="E"'
                        worksheet.conditional_format(
                            f'{first_letter}2:{last_letter}{last_row}',
                            {'type': 'formula', 'criteria': formula, 'format': blue_format}
                        )
                
                # Auto-fit columns
                for i, col in enumerate(df.columns):
                    col_data = df[col].astype(str)
                    max_len = max(col_data.map(len).max() or 0, len(str(col)))
                    width = min(max_len + 2, self.max_col_width)
                    worksheet.set_column(i, i, width)
            
            if progress_callback:
                progress_callback(100, "File saved successfully")
            
            logger.info(f"Saved output to: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error saving Excel: {e}", exc_info=True)
            raise


def read_excel_data(file_path: Path, config: Dict[str, Any], progress_callback=None) -> Tuple[pd.DataFrame, int]:
    """Convenience function to read Excel data.
    
    Args:
        file_path: Path to Excel file
        config: Configuration dictionary
        progress_callback: Optional progress callback
        
    Returns:
        Tuple of (DataFrame, header_row_index)
    """
    handler = VinafcoExcelHandler(config)
    return handler.read_excel(file_path, progress_callback)


def save_and_format_excel(df: pd.DataFrame, output_path: Path, config: Dict[str, Any], progress_callback=None) -> bool:
    """Convenience function to save formatted Excel.
    
    Args:
        df: DataFrame to save
        output_path: Output file path
        config: Configuration dictionary
        progress_callback: Optional progress callback
        
    Returns:
        True if successful
    """
    handler = VinafcoExcelHandler(config)
    return handler.save_excel(df, output_path, progress_callback)
