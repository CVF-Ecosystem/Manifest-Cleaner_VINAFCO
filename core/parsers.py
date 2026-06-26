"""Regex patterns and parsing functions for VINAFCO Manifest.

This module contains all regex patterns used for parsing
manifest data and functions to extract specific information.
"""

from __future__ import annotations

import re
import logging
from typing import Any, Dict, Optional

import pandas as pd


logger = logging.getLogger(__name__)

# --- Compiled Regex Patterns ---

# Container number pattern: 4 letters + 7 digits (e.g., ABCD1234567)
CONT_NO_PATTERN = re.compile(r'^([A-Z]{4}\d{7})')

# Keywords that indicate end of party name (contact info starts)
STOP_KEYWORDS_PATTERN = re.compile(
    r'(?:Mail\s*:|Tel\s*:|Fax\s*:|Địa chỉ\s*:|SĐT\s*:|Email\s*:|'
    r'Số CCCD\s*:|Tại\s*:|Chức vụ\s*:|Tài khoản\s*:|Đ/C\s*:|MST\s*:|'
    r'Phone\s*:|Điện thoại\s*:|ADD\s*:|ADDRESS\s*:)',
    re.IGNORECASE | re.UNICODE
)

# Size/Type pattern in shipper line (e.g., 20DC, 40'HC)
SIZE_TYPE_PATTERN = re.compile(r"(\d{2}'?[A-Z]{2,3})")

# SOC marker in shipper line
SOC_MARKER_PATTERN = re.compile(r'\bSOC\b', re.IGNORECASE)

# Empty container description pattern
EMPTY_DESC_PATTERN = re.compile(
    r'\b(EMPTY|RONG|CHUYEN RONG|CONT RONG|CONTAINER RONG)\b',
    re.IGNORECASE | re.UNICODE
)

# Weight pattern for VINAFCO format
# Matches: "25,000.50 KGS", "25.000,50 TONS", "25000", etc.
WEIGHT_PATTERN = re.compile(
    r'^([\d.,]+)\s*(KGS?|TONS?|T)?$',
    re.IGNORECASE
)


def parse_shipper_line_info(line_text: str, party_marker_len: int = 3) -> Dict[str, Any]:
    """Parse shipper/consignee/notify line to extract name and SOC status.
    
    Args:
        line_text: Raw line text starting with (S), (C), or (N)
        party_marker_len: Length of party marker (default 3 for "(S)")
        
    Returns:
        Dictionary with 'name' and 'is_soc_from_shipper_line' keys
        
    Example:
        >>> parse_shipper_line_info("(S) ABC COMPANY SOC, 20DC")
        {'name': 'ABC COMPANY SOC', 'is_soc_from_shipper_line': True}
    """
    party_name = ""
    is_soc = False
    
    if line_text.startswith(("(S)", "(C)", "(N)", "(X)")):
        # Get content after marker
        actual_offset = party_marker_len if not line_text.startswith("(X)") else len("(X)")
        content_after_marker = line_text[actual_offset:].lstrip()
        
        # Find where size/type info starts (if any)
        size_match = re.search(r"\s*,\s*(" + SIZE_TYPE_PATTERN.pattern + r")$", content_after_marker)
        if size_match:
            party_name = content_after_marker[:size_match.start()].strip()
        else:
            party_name = content_after_marker.strip()
        
        # Check for SOC marker in shipper line
        if line_text.startswith("(S)") and SOC_MARKER_PATTERN.search(line_text):
            is_soc = True
    
    return {
        "name": party_name,
        "is_soc_from_shipper_line": is_soc
    }


def parse_weight_line(line: str, default_unit: str = "TONS") -> Optional[Dict[str, Any]]:
    """Parse weight string to extract value in TONS.
    
    Handles various formats:
    - "25,000.50 KGS" -> 25.00050 TONS
    - "25.000,50 TONS" -> 25.0005 TONS  
    - "25000" -> depends on default_unit
    
    Args:
        line: Raw weight string
        default_unit: Default unit if not specified (KGS or TONS)
        
    Returns:
        Dictionary with 'value' (in TONS) and 'unit' (always "TONS"),
        or None if parsing fails
        
    Example:
        >>> parse_weight_line("25,000 KGS")
        {'value': 25.0, 'unit': 'TONS'}
    """
    if not line or pd.isna(line):
        return None
    
    raw_value = str(line).strip()
    if not raw_value:
        return None
    
    match = WEIGHT_PATTERN.match(raw_value)
    if not match:
        logger.debug(f"Weight string '{raw_value}' does not match pattern")
        return None
    
    numeric_part = match.group(1)
    unit_part = match.group(2)
    
    # Determine unit
    if unit_part:
        unit = unit_part.upper()
    else:
        unit = default_unit.upper()
    
    # Standardize numeric format
    # Handle different decimal/thousands separators
    num_dots = numeric_part.count('.')
    num_commas = numeric_part.count(',')
    
    standardized = numeric_part
    
    if num_dots == 1 and num_commas >= 1:
        # Format: 25,000.50 (comma as thousands, dot as decimal)
        standardized = numeric_part.replace(',', '')
    elif num_commas == 1 and num_dots >= 1:
        # Format: 25.000,50 (dot as thousands, comma as decimal)
        standardized = numeric_part.replace('.', '').replace(',', '.')
    elif num_commas == 1 and num_dots == 0:
        # Format: 25,50 (comma as decimal)
        standardized = numeric_part.replace(',', '.')
    elif num_dots > 1:
        # Format: 25.000.000 (dots as thousands only)
        standardized = numeric_part.replace('.', '', num_dots - 1)
    elif num_commas > 1:
        # Format: 25,000,000 (commas as thousands only)
        standardized = numeric_part.replace(',', '')
    
    try:
        weight_val = float(standardized)
        
        # Convert to TONS if in KGS
        if unit in ["KG", "KGS"]:
            weight_val /= 1000.0
        elif unit not in ["TONS", "TON", "T"]:
            logger.warning(f"Unknown weight unit '{unit}', assuming TONS")
        
        return {
            "value": round(weight_val, 4),
            "unit": "TONS"
        }
        
    except ValueError:
        logger.warning(f"Cannot convert weight '{numeric_part}' to float")
        return None


def is_container_line(container_no: str) -> bool:
    """Check if a string is a valid container number.
    
    Args:
        container_no: String to check
        
    Returns:
        True if valid container number format
    """
    if not container_no:
        return False
    return bool(CONT_NO_PATTERN.match(container_no.strip()))


def is_party_marker_line(text: str) -> bool:
    """Check if a line starts with party marker.
    
    Args:
        text: Line text to check
        
    Returns:
        True if line starts with (S), (C), or (N)
    """
    return text.startswith(("(S)", "(C)", "(N)"))
