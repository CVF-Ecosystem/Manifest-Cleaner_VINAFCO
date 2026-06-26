"""Configuration module for VINAFCO Manifest Cleaner.

This module contains application constants, UI text translations,
and data mappings used throughout the application.
"""

from config.constants import (
    APP_TITLE,
    APP_VERSION,
    APP_AUTHOR_VI,
    APP_AUTHOR_EN,
    UI_TEXT,
    DEFAULT_LANGUAGE,
    COLORS,
)

from config.mappings import (
    SIZE_TO_ISO,
    SIZE_TYPE_TO_CARGO_TYPE_BASE,
    TARE_WEIGHTS_TONS,
    SAFETY_MARGIN_TONS,
)

__all__ = [
    "APP_TITLE",
    "APP_VERSION",
    "APP_AUTHOR_VI",
    "APP_AUTHOR_EN",
    "UI_TEXT",
    "DEFAULT_LANGUAGE",
    "COLORS",
    "SIZE_TO_ISO",
    "SIZE_TYPE_TO_CARGO_TYPE_BASE",
    "TARE_WEIGHTS_TONS",
    "SAFETY_MARGIN_TONS",
]
