#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""VINAFCO Manifest Cleaner Application Entry Point.

This is the main entry point for the VINAFCO Manifest Cleaner application.
Run this script to start the application.

Version: 2.0.0
Author: Tien-Tan Thuan Port
"""

from __future__ import annotations

import sys
import logging
from pathlib import Path

# Add application directory to path for imports
app_dir = Path(__file__).parent
if str(app_dir) not in sys.path:
    sys.path.insert(0, str(app_dir))


def main():
    """Application entry point."""
    try:
        # Import and initialize logging
        from utils.helpers import setup_logging
        setup_logging("vinafco_manifest.log")
        
        logger = logging.getLogger(__name__)
        logger.info("=" * 60)
        logger.info("VINAFCO Manifest Cleaner v2.0.0 Starting...")
        logger.info("=" * 60)
        
        # Import and create application
        from gui.main_window import create_app
        
        app = create_app()
        app.run()
        
    except ImportError as e:
        # Handle missing dependencies
        error_msg = f"Missing required library: {e}"
        print(f"ERROR: {error_msg}")
        print("\nPlease install required packages:")
        print("  pip install pandas openpyxl xlrd xlsxwriter")
        
        try:
            import tkinter.messagebox as mb
            mb.showerror("Missing Dependencies", error_msg)
        except:
            pass
        
        sys.exit(1)
        
    except Exception as e:
        # Handle unexpected errors
        error_msg = f"Unexpected error: {e}"
        print(f"CRITICAL ERROR: {error_msg}")
        
        try:
            logging.exception("Critical application error")
        except:
            pass
        
        try:
            import tkinter.messagebox as mb
            mb.showerror("Critical Error", error_msg)
        except:
            pass
        
        sys.exit(1)


if __name__ == "__main__":
    main()
