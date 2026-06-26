
import unicodedata
import re
import pandas as pd

def _remove_vietnamese_diacritics(text: str) -> str:
    """Remove Vietnamese diacritics from text."""
    if not isinstance(text, str):
        return str(text) if text is not None else ""
    
    text = unicodedata.normalize('NFD', text)
    text = re.sub(r'[\u0300-\u036f]', '', text)
    # Handle D/d
    text = text.replace('Đ', 'D').replace('đ', 'd')
    return text

def normalize_text(val) -> str:
    """
    Normalize text: remove diacritics, uppercase, clean whitespace.
    Safely handles None/NaN values.
    """
    if pd.isna(val) or val == "":
        return ""
    
    text = str(val)
    text_no_diacritics = _remove_vietnamese_diacritics(text)
    normalized = text_no_diacritics.strip().upper()
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized
