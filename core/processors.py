"""Data processing functions for VINAFCO manifest data.

This module contains functions for processing and transforming
manifest data, including F/E determination, name cleaning, and
cargo type classification.
"""

from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional

import pandas as pd

from config.mappings import (
    SIZE_TO_ISO,
    SIZE_TYPE_TO_CARGO_TYPE_BASE,
    TARE_WEIGHTS_TONS,
    REEFER_TARE_WEIGHTS_TONS,
    SAFETY_MARGIN_TONS,
    MIN_PAYLOAD_TONS,
)
from core.parsers import (
    EMPTY_DESC_PATTERN,
    STOP_KEYWORDS_PATTERN,
    parse_shipper_line_info,
)

import re

logger = logging.getLogger(__name__)


# Pattern to match cargo item with quantity and container type
# Example: "GACH MEN (01) x20'DC COC" -> ("GACH MEN", 1, "20DC")
CARGO_QTY_PATTERN = re.compile(
    r"([A-Za-zÀ-ỹ\s\-/\.()]+?)"  # Cargo name (including Vietnamese chars)
    r"\s*\((\d+)\)\s*"           # Quantity in parentheses like (01)
    r"x\s*(\d{2})'?([A-Z0-9]{1,4})"  # Container type: xNN'XX or xNNXX
    r"\s*(?:COC|SOC)?",           # Optional COC/SOC marker
    re.IGNORECASE | re.UNICODE
)


def get_iso_size(size_type: Optional[str]) -> Optional[str]:
    """Convert size/type code to ISO standard format.
    
    Args:
        size_type: Raw size/type code (e.g., "20'DC", "40HC")
        
    Returns:
        ISO standard code or None if not found
        
    Example:
        >>> get_iso_size("20'DC")
        '20DC'
        >>> get_iso_size("40GP")
        '40HC'
    """
    if not size_type:
        return None
    
    normalized = size_type.upper().replace("'", "").strip()
    return SIZE_TO_ISO.get(normalized) or SIZE_TO_ISO.get(size_type.upper())


def determine_fe(
    description: Optional[str],
    weight_info: Optional[Dict[str, Any]],
    size_type: Optional[str]
) -> str:
    """Determine Full/Empty status of a container.
    
    Uses a multi-rule system prioritizing weight over description.
    
    Priority:
        1. Weight definitively indicates FULL (gross > tare + margin)
        2. Description says EMPTY → E
        3. Weight indicates EMPTY (gross ≤ tare)
        4. Very low weight → E
        5. Non-empty description → F
        6. Default → F
    
    Args:
        description: Cargo description text
        weight_info: Dictionary with 'value' (float in TONS) and 'unit'
        size_type: Container size/type code (e.g., "40HC")
        
    Returns:
        "F" for Full, "E" for Empty
        
    Example:
        >>> determine_fe("EMPTY CONTAINER", {"value": 4.0, "unit": "TONS"}, "40HC")
        'E'
        >>> determine_fe("FROZEN FISH", {"value": 25.0, "unit": "TONS"}, "40HC")
        'F'
    """
    # Get tare weight based on size
    tare_weight = 0.0
    if size_type:
        size_prefix = size_type[:2].replace("'", "")
        is_reefer = "RF" in size_type.upper() or "RH" in size_type.upper()
        
        if is_reefer:
            tare_weight = REEFER_TARE_WEIGHTS_TONS.get(size_prefix, 4.0)
        else:
            tare_weight = TARE_WEIGHTS_TONS.get(size_prefix, 3.0)
    
    # Rule 1: Weight definitively indicates FULL
    if weight_info and weight_info.get("unit") == "TONS":
        gross_weight = float(weight_info.get("value") or 0)
        if gross_weight > tare_weight + SAFETY_MARGIN_TONS + MIN_PAYLOAD_TONS:
            return "F"
    
    # Rule 2: Description says EMPTY
    if description and EMPTY_DESC_PATTERN.search(str(description)):
        return "E"
    
    # Rule 3 & 4: Weight indicates EMPTY
    if weight_info and weight_info.get("unit") == "TONS":
        gross_weight = float(weight_info.get("value") or 0)
        
        # Very low weight
        if gross_weight < MIN_PAYLOAD_TONS:
            return "E"
        
        # Weight at or below tare
        if tare_weight > 0 and gross_weight <= tare_weight:
            return "E"
    
    # Rule 5: Any non-empty description defaults to FULL
    if description and str(description).strip():
        return "F"
    
    # Rule 6: Default
    return "F"


def determine_cargo_type(size_type: Optional[str], fe_status: str) -> str:
    """Determine cargo type based on container type and F/E status.
    
    Args:
        size_type: Container size/type code (e.g., "40HC", "20RF")
        fe_status: Full/Empty status ("F" or "E")
        
    Returns:
        Cargo type string (e.g., "GENERAL", "EMPTY", "REEFER", "EMPTY REEFER")
        
    Example:
        >>> determine_cargo_type("40HC", "F")
        'GENERAL'
        >>> determine_cargo_type("40HC", "E")
        'EMPTY'
        >>> determine_cargo_type("20RF", "F")
        'REEFER'
        >>> determine_cargo_type("20RF", "E")
        'EMPTY REEFER'
    """
    if not size_type:
        return "UNKNOWN"
    
    # Normalize size type
    normalized = size_type.upper().replace("'", "").strip()
    
    # Get base cargo type
    base_type = SIZE_TYPE_TO_CARGO_TYPE_BASE.get(normalized)
    
    if not base_type:
        # Try with ISO code
        iso_code = get_iso_size(size_type)
        if iso_code:
            base_type = SIZE_TYPE_TO_CARGO_TYPE_BASE.get(iso_code, "GENERAL")
        else:
            base_type = "GENERAL"
    
    # Apply F/E status
    if fe_status == "E":
        if base_type == "GENERAL":
            return "EMPTY"
        else:
            return f"EMPTY {base_type}"
    
    return base_type


def clean_party_name(full_party_text: str, party_marker_len: int = 3) -> str:
    """Extract clean party name from full text.
    
    Removes contact information (phone, email, address, etc.)
    and keeps only the company/person name.
    
    Args:
        full_party_text: Full party text with potential contact info
        party_marker_len: Length of party marker
        
    Returns:
        Cleaned party name
        
    Example:
        >>> clean_party_name("ABC COMPANY Tel: 123456 Email: a@b.com")
        'ABC COMPANY'
    """
    if not full_party_text or pd.isna(full_party_text):
        return ""
    
    # Get first line only
    first_line = str(full_party_text).split('\n')[0].strip()
    
    # Use parser to extract name if it's a party marker line
    if first_line.startswith(("(S)", "(C)", "(N)")):
        parsed = parse_shipper_line_info(f"(X){first_line[3:]}", party_marker_len)
        if parsed["name"]:
            name = parsed["name"]
        else:
            name = first_line[party_marker_len:].strip()
    else:
        name = first_line
    
    # Find where contact info starts
    stop_match = STOP_KEYWORDS_PATTERN.search(name)
    if stop_match:
        name = name[:stop_match.start()].strip()
    
    # Remove trailing punctuation
    return name.rstrip(';,./\\- ')


def refine_description(df: pd.DataFrame, goods_col: str, fe_col: str, cargo_type_col: str) -> pd.DataFrame:
    """Refine goods description based on F/E status.
    
    For empty containers, sets description to "EMPTY".
    
    Args:
        df: DataFrame to modify
        goods_col: Name of goods description column
        fe_col: Name of F/E status column
        cargo_type_col: Name of cargo type column
        
    Returns:
        Modified DataFrame
    """
    if goods_col not in df.columns:
        logger.warning(f"Column '{goods_col}' not found, skipping description refinement")
        return df
    
    if fe_col not in df.columns:
        logger.warning(f"Column '{fe_col}' not found, skipping description refinement")
        return df
    
    df = df.copy()
    
    # Clean up description
    df[goods_col] = df[goods_col].astype(str).str.strip()
    df[goods_col] = df[goods_col].replace('nan', '')
    
    # Set empty container description
    if cargo_type_col in df.columns:
        is_empty_fe = df[fe_col] == 'E'
        is_empty_cargo = df[cargo_type_col].str.startswith("EMPTY", na=False) | (df[cargo_type_col] == "EMPTY")
        df.loc[is_empty_fe & is_empty_cargo, goods_col] = "EMPTY"
    else:
        df.loc[df[fe_col] == 'E', goods_col] = "EMPTY"
    
    return df


class VinafcoManifestProcessor:
    """Main processor class for VINAFCO manifest files.
    
    Handles the complete processing pipeline from reading Excel
    to generating formatted output.
    """
    
    def __init__(self, config: Dict[str, Any]):
        """Initialize processor with configuration.
        
        Args:
            config: Dictionary containing processing configuration
        """
        self.config = config
        self.df_raw: Optional[pd.DataFrame] = None
        self.intermediate_data: List[Dict[str, Any]] = []
        self.final_df: Optional[pd.DataFrame] = None
        
        # Load config values
        self.target_sheet = config.get('target_sheet_name', 'MANIFEST')
        self.party_marker_len = config.get('party_marker_len', 3)
        self.operator_soc = config.get('operator_soc', 'SVF')
        self.operator_coc = config.get('operator_coc', 'VFC')
        
        # Column names
        self.col_party_raw = config.get('col_party_raw', '(S) SHIPPER,(C) CONSIGNEE,(N)NOTIFY PARTY')
        self.col_bill_no = config.get('col_bill_no', 'Bill No')
        self.col_container_no = config.get('col_container_no', 'CONTAINER No.')
        self.col_seal = config.get('col_seal', 'SEAL')
        self.col_type = config.get('col_type', 'TYPE')
        self.col_comm = config.get('col_comm', 'COM.M')
        self.col_gw = config.get('col_gw', 'GW')
        self.col_remarks = config.get('col_remarks', 'REMARKS')
        
        logger.info(f"VinafcoManifestProcessor initialized with config: {config}")
    
    def get_container_count(self) -> int:
        """Get number of containers in final output."""
        if self.final_df is not None:
            return len(self.final_df)
        return len(self.intermediate_data)
    
    def get_full_empty_counts(self) -> tuple:
        """Get counts of full and empty containers.
        
        Returns:
            Tuple of (full_count, empty_count)
        """
        if self.final_df is None:
            return (0, 0)
        
        fe_col = None
        for col in self.final_df.columns:
            if 'F/E' in col or col.upper() == 'FE':
                fe_col = col
                break
        
        if fe_col is None:
            return (0, 0)
        
        full_count = len(self.final_df[self.final_df[fe_col] == 'F'])
        empty_count = len(self.final_df[self.final_df[fe_col] == 'E'])
        
        return (full_count, empty_count)


def parse_cargo_items_with_qty(description: str) -> List[tuple]:
    """Parse cargo description to extract items with quantities and container types.
    
    Args:
        description: Raw description like "GACH MEN (01) x20'DC COC, PHAN BON (01) x20'DC COC"
        
    Returns:
        List of tuples: (cargo_name, quantity, container_type)
    """
    if not description or pd.isna(description):
        return []
    
    desc_str = str(description).strip()
    items = []
    
    for match in CARGO_QTY_PATTERN.finditer(desc_str):
        cargo_name = match.group(1).strip()
        qty = int(match.group(2))
        size_prefix = match.group(3)
        size_suffix = match.group(4).upper()
        container_type = f"{size_prefix}{size_suffix}"
        
        if cargo_name:
            items.append((cargo_name, qty, container_type))
    
    return items


def match_description_by_quantity(
    containers: List[Dict[str, Any]],
    description: str
) -> List[str]:
    """Match cargo descriptions to containers based on quantity and order.
    
    When a manifest line has multiple cargo items with quantities like:
    "GACH MEN (01) x20'DC COC, PHAN BON (01) x20'DC COC"
    
    This function assigns each cargo to containers in order.
    
    Args:
        containers: List of container dicts with 'Size_Type' or 'TYPE' key
        description: Raw description text
        
    Returns:
        List of descriptions matched to each container in order
    """
    if not containers:
        return []
    
    n = len(containers)
    result = [""] * n
    
    cargo_items = parse_cargo_items_with_qty(description)
    
    if not cargo_items:
        # No quantity info - return original for all
        return [description.strip() if description else ""] * n
    
    # Group containers by type with their indices
    type_to_indices: Dict[str, List[int]] = {}
    for i, cont in enumerate(containers):
        cont_type = str(cont.get("Size_Type") or cont.get("TYPE") or cont.get("type") or "").upper().replace("'", "")
        if cont_type not in type_to_indices:
            type_to_indices[cont_type] = []
        type_to_indices[cont_type].append(i)
    
    # Track assignment position for each container type
    type_assign_pos: Dict[str, int] = {t: 0 for t in type_to_indices}
    
    # Assign cargo to containers in order
    for cargo_name, qty, cont_type in cargo_items:
        cont_type_upper = cont_type.upper().replace("'", "")
        if cont_type_upper not in type_to_indices:
            continue
        
        indices = type_to_indices[cont_type_upper]
        start_pos = type_assign_pos[cont_type_upper]
        
        for _ in range(qty):
            if start_pos >= len(indices):
                break
            idx = indices[start_pos]
            if not result[idx]:
                result[idx] = cargo_name
            start_pos += 1
        
        type_assign_pos[cont_type_upper] = start_pos
    
    # Fill remaining with original description
    for i, desc in enumerate(result):
        if not desc:
            result[i] = description.strip() if description else ""
    
    return result
