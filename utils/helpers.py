"""Utility functions for VINAFCO Manifest Cleaner.

This module contains helper functions used throughout the application
including path resolution, logging setup, and data conversion utilities.
"""

from __future__ import annotations

import configparser
import logging
import sys
from pathlib import Path
from typing import Any, Dict, Optional

import pandas as pd


logger = logging.getLogger(__name__)


def get_application_path() -> Path:
    """Get the application's base path.
    
    Handles both normal Python execution and PyInstaller frozen executables.
    
    Returns:
        Path to the application directory
    """
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # Running as PyInstaller bundle
        return Path(sys._MEIPASS)
    else:
        # Running as normal Python script
        return Path(__file__).parent.parent


def get_resource_path(relative_path: str) -> Path:
    """Get absolute path to a resource file.
    
    Args:
        relative_path: Path relative to application directory
        
    Returns:
        Absolute path to the resource
    """
    base_path = get_application_path()
    return base_path / relative_path


def setup_logging(log_file_name: str = "vinafco_manifest.log", log_level: int = logging.INFO) -> None:
    """Configure application logging.
    
    Args:
        log_file_name: Name of log file
        log_level: Logging level (default INFO)
    """
    # Create logs directory
    log_dir = get_application_path() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    
    log_path = log_dir / log_file_name
    
    # Clear existing handlers
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    
    # Configure logging
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(levelname)s - %(module)s - %(funcName)s - %(lineno)d - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        handlers=[
            logging.FileHandler(log_path, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    logger.info(f"Logging initialized. Log file: {log_path}")


def get_cell_value_as_str(cell_value: Any) -> str:
    """Convert cell value to clean string.
    
    Handles None, NaN, and various types, returning empty string
    for null values.
    
    Args:
        cell_value: Cell value from DataFrame
        
    Returns:
        Cleaned string value
    """
    if pd.isna(cell_value) or cell_value is None or cell_value == '':
        return ""
    return str(cell_value).strip()


def load_ini_config(config_file: str = "config.ini") -> Dict[str, Any]:
    """Load configuration from INI file.
    
    Args:
        config_file: Name of config file
        
    Returns:
        Dictionary containing configuration values
        
    Raises:
        FileNotFoundError: If config file doesn't exist
        ValueError: If required sections are missing
    """
    config_path = get_resource_path(config_file)
    
    if not config_path.exists():
        raise FileNotFoundError(f"Configuration file not found: {config_path}")
    
    parser = configparser.ConfigParser(interpolation=None)
    parser.read(config_path, encoding='utf-8')
    
    # Check required sections
    required_sections = ['General', 'VinafcoSpecific', 'RegexPatterns', 'Formatting', 'LogicThresholds']
    missing = [s for s in required_sections if not parser.has_section(s)]
    
    if missing:
        raise ValueError(f"Missing required config sections: {missing}")
    
    # Build config dictionary
    config = {}
    
    # General settings
    gen = parser['General']
    config['log_file_name'] = gen.get('log_file_name', 'vinafco_manifest.log')
    config['excel_engine_xlsx'] = gen.get('excel_engine_xlsx', 'openpyxl')
    config['excel_engine_xls'] = gen.get('excel_engine_xls', 'xlrd')
    config['peek_rows'] = gen.getint('peek_rows', 30)
    config['output_sheet_name'] = gen.get('output_sheet_name', 'Processed_Manifest')
    config['max_column_width'] = gen.getint('max_column_width', 50)
    config['language'] = gen.get('language', 'vi')
    
    # VINAFCO specific settings
    vs = parser['VinafcoSpecific']
    config['target_sheet_name'] = vs.get('target_sheet_name', 'MANIFEST')
    config['party_marker_len'] = vs.getint('party_marker_len', 3)
    config['operator_soc'] = vs.get('operator_soc', 'SVF')
    config['operator_coc'] = vs.get('operator_coc', 'VFC')
    config['default_weight_unit'] = vs.get('default_weight_unit_if_missing', 'TONS')
    
    # Parse expected columns (lowercase)
    expected_cols_str = vs.get('expected_input_columns_lower', '')
    for line in expected_cols_str.splitlines():
        if ':' in line:
            key, value = line.split(':', 1)
            key = key.strip()
            value = value.strip()
            if key == 'party_raw':
                config['expected_party_lower'] = value
            elif key == 'bill_no':
                config['expected_bill_lower'] = value
    
    # Parse input column names
    input_cols_str = vs.get('input_column_names', '')
    col_mapping = {
        'party_raw': 'col_party_raw',
        'bill_no': 'col_bill_no',
        'container_no': 'col_container_no',
        'seal': 'col_seal',
        'type': 'col_type',
        'comm': 'col_comm',
        'gw': 'col_gw',
        'remarks': 'col_remarks',
    }
    for line in input_cols_str.splitlines():
        if ':' in line:
            key, value = line.split(':', 1)
            key = key.strip()
            value = value.strip()
            if key in col_mapping:
                config[col_mapping[key]] = value
    
    # Parse output column mapping
    output_mapping_str = vs.get('output_column_mapping', '')
    config['output_column_mapping'] = {}
    for line in output_mapping_str.splitlines():
        if ':' in line:
            key, value = line.split(':', 1)
            config['output_column_mapping'][key.strip()] = value.strip()
    
    # Parse desired output columns
    desired_cols_str = vs.get('desired_output_columns', '')
    config['desired_output_columns'] = [col.strip() for col in desired_cols_str.split(',') if col.strip()]
    
    # Output column names
    config['seal_column_output_name'] = vs.get('seal_column_output_name', 'Số niêm chì')
    config['consignee_column_output_name'] = vs.get('consignee_column_output_name', 'Chủ hàng')
    config['fe_column_output_name'] = vs.get('fe_column_output_name', 'F/E')
    config['temp_count_column_name'] = vs.get('temp_count_column_name', 'Số lượng_temp')
    config['goods_column_output_name'] = config['output_column_mapping'].get('Description', 'Hàng hóa')
    
    # Formatting settings
    fmt = parser['Formatting']
    config['seal_invalid_bg_color'] = fmt.get('seal_invalid_bg_color', '#FFFF00')
    config['empty_container_bg_color'] = fmt.get('empty_container_bg_color', '#DAEEF3')
    
    # Logic thresholds
    lt = parser['LogicThresholds']
    config['tare_40ft'] = lt.getfloat('tare_40ft_empty_vgm', 3.8)
    config['tare_20ft'] = lt.getfloat('tare_20ft_empty_vgm', 2.3)
    config['tare_40rf'] = lt.getfloat('tare_40rf_empty_vgm', 4.8)
    config['tare_20rf'] = lt.getfloat('tare_20rf_empty_vgm', 3.0)
    config['min_payload'] = lt.getfloat('min_payload_for_f_vgm', 0.1)
    
    logger.info(f"Loaded configuration from: {config_path}")
    return config


def format_file_size(size_bytes: int) -> str:
    """Format file size in human-readable format.
    
    Args:
        size_bytes: File size in bytes
        
    Returns:
        Formatted string (e.g., "1.5 MB")
    """
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


def sanitize_filename(filename: str) -> str:
    """Remove invalid characters from filename.
    
    Args:
        filename: Original filename
        
    Returns:
        Sanitized filename safe for filesystem
    """
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename.strip()


def is_valid_excel_file(file_path: Path) -> bool:
    """Check if file is a valid Excel file.
    
    Args:
        file_path: Path to check
        
    Returns:
        True if valid Excel file
    """
    if not file_path.exists():
        return False
    
    valid_extensions = {'.xls', '.xlsx', '.xlsm'}
    return file_path.suffix.lower() in valid_extensions
