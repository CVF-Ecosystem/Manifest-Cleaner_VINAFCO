"""Main Window GUI for VINAFCO Manifest Cleaner.

This module contains the main application window with all UI components
including Settings menu, Recent files panel, Preview window, and keyboard shortcuts.
"""

from __future__ import annotations

import configparser
import json
import logging
import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd

from config.constants import APP_TITLE, APP_VERSION, UI_TEXT, COLORS
from core.excel_handler import VinafcoExcelHandler
from utils.helpers import (
    get_application_path, 
    get_resource_path, 
    setup_logging,
    load_ini_config,
    is_valid_excel_file,
    format_file_size
)


logger = logging.getLogger(__name__)


class StringManager:
    """Manages bilingual UI strings."""
    
    def __init__(self, language: str = 'vi'):
        self.language = language
        self.strings: Dict[str, Dict[str, str]] = {}
        self._load_strings()
    
    def _load_strings(self) -> None:
        """Load strings from INI files."""
        lang_file = f"strings_{self.language}.ini"
        lang_path = get_resource_path(lang_file)
        
        if not lang_path.exists():
            logger.warning(f"Language file not found: {lang_path}, falling back to vi")
            lang_path = get_resource_path("strings_vi.ini")
        
        if lang_path.exists():
            parser = configparser.ConfigParser(interpolation=None)
            parser.read(lang_path, encoding='utf-8')
            
            for section in parser.sections():
                self.strings[section] = dict(parser[section])
            
            logger.info(f"Loaded language: {self.language}")
    
    def get(self, section: str, key: str, default: str = "") -> str:
        """Get a string value."""
        return self.strings.get(section, {}).get(key, default or key)
    
    def set_language(self, language: str) -> None:
        """Change language and reload strings."""
        self.language = language
        self._load_strings()


class SettingsManager:
    """Manages application settings."""
    
    DEFAULT_SETTINGS = {
        'language': 'vi',
        'theme': 'light',
        'auto_open_output': True,
        'show_preview': True,
        'confirm_exit': True,
        'remember_path': True,
        'last_directory': '',
        'recent_files': [],
        'max_recent_files': 10,
        'window_geometry': '900x700',
        'preview_collapsed': False
    }
    
    def __init__(self, config_file: str = "app_config.json"):
        self.config_path = get_resource_path(config_file)
        self.settings: Dict[str, Any] = self.DEFAULT_SETTINGS.copy()
        self._load()
    
    def _load(self) -> None:
        """Load settings from JSON file."""
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    saved = json.load(f)
                    self.settings.update(saved)
                logger.info(f"Settings loaded from: {self.config_path}")
            except Exception as e:
                logger.error(f"Error loading settings: {e}")
    
    def save(self) -> None:
        """Save settings to JSON file."""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
            logger.info("Settings saved")
        except Exception as e:
            logger.error(f"Error saving settings: {e}")
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get a setting value."""
        return self.settings.get(key, default)
    
    def set(self, key: str, value: Any) -> None:
        """Set a setting value."""
        self.settings[key] = value
        self.save()
    
    def add_recent_file(self, file_path: str) -> None:
        """Add a file to recent files list."""
        recent = self.settings['recent_files']
        
        # Remove if already exists
        if file_path in recent:
            recent.remove(file_path)
        
        # Add to front
        recent.insert(0, file_path)
        
        # Trim to max
        max_files = self.settings['max_recent_files']
        self.settings['recent_files'] = recent[:max_files]
        self.save()
    
    def clear_recent_files(self) -> None:
        """Clear recent files list."""
        self.settings['recent_files'] = []
        self.save()


class PreviewWindow(tk.Toplevel):
    """Preview window for processed data."""
    
    def __init__(self, parent: tk.Tk, df: pd.DataFrame, strings: StringManager):
        super().__init__(parent)
        self.df = df
        self.strings = strings
        
        self.title(self.strings.get('Labels', 'preview_title', 'Preview'))
        self.geometry("1000x600")
        self.minsize(800, 400)
        
        self._setup_ui()
        self._populate_data()
        
        # Center on parent
        self.transient(parent)
        self.grab_set()
    
    def _setup_ui(self) -> None:
        """Setup preview window UI."""
        # Header frame
        header_frame = ttk.Frame(self)
        header_frame.pack(fill=tk.X, padx=10, pady=5)
        
        row_count = len(self.df)
        ttk.Label(
            header_frame, 
            text=f"{self.strings.get('Preview', 'row_count', 'Total rows:')} {row_count}",
            font=('Segoe UI', 10, 'bold')
        ).pack(side=tk.LEFT)
        
        # Close button
        ttk.Button(
            header_frame,
            text=self.strings.get('Buttons', 'close', 'Close'),
            command=self.destroy
        ).pack(side=tk.RIGHT)
        
        # Treeview with scrollbars
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Scrollbars
        y_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        x_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Treeview
        columns = list(self.df.columns)
        self.tree = ttk.Treeview(
            tree_frame, 
            columns=columns, 
            show='headings',
            yscrollcommand=y_scroll.set,
            xscrollcommand=x_scroll.set
        )
        
        y_scroll.config(command=self.tree.yview)
        x_scroll.config(command=self.tree.xview)
        
        # Configure columns
        for col in columns:
            self.tree.heading(col, text=col, anchor=tk.W)
            self.tree.column(col, width=100, minwidth=50, anchor=tk.W)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
    
    def _populate_data(self) -> None:
        """Populate treeview with DataFrame data."""
        # Limit to first 500 rows for performance
        display_df = self.df.head(500)
        
        for idx, row in display_df.iterrows():
            values = [str(v) if pd.notna(v) else '' for v in row.values]
            self.tree.insert('', tk.END, values=values)


class SettingsDialog(tk.Toplevel):
    """Settings dialog window."""
    
    def __init__(self, parent: tk.Tk, settings: SettingsManager, 
                 strings: StringManager, on_save: Callable):
        super().__init__(parent)
        self.settings = settings
        self.strings = strings
        self.on_save = on_save
        
        self.title(self.strings.get('Labels', 'settings_title', 'Settings'))
        self.geometry("400x350")
        self.resizable(False, False)
        
        self._setup_ui()
        
        # Center on parent
        self.transient(parent)
        self.grab_set()
    
    def _setup_ui(self) -> None:
        """Setup settings dialog UI."""
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Language setting
        lang_frame = ttk.Frame(main_frame)
        lang_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(
            lang_frame, 
            text=self.strings.get('Labels', 'settings_language', 'Language:')
        ).pack(side=tk.LEFT)
        
        self.lang_var = tk.StringVar(value=self.settings.get('language', 'vi'))
        lang_combo = ttk.Combobox(
            lang_frame, 
            textvariable=self.lang_var,
            values=['vi', 'en'],
            state='readonly',
            width=15
        )
        lang_combo.pack(side=tk.RIGHT)
        
        # Auto-open output checkbox
        self.auto_open_var = tk.BooleanVar(
            value=self.settings.get('auto_open_output', True)
        )
        ttk.Checkbutton(
            main_frame,
            text=self.strings.get('Labels', 'settings_auto_open', 'Auto-open file after processing'),
            variable=self.auto_open_var
        ).pack(fill=tk.X, pady=5, anchor=tk.W)
        
        # Show preview checkbox
        self.show_preview_var = tk.BooleanVar(
            value=self.settings.get('show_preview', True)
        )
        ttk.Checkbutton(
            main_frame,
            text=self.strings.get('Labels', 'settings_show_preview', 'Show preview'),
            variable=self.show_preview_var
        ).pack(fill=tk.X, pady=5, anchor=tk.W)
        
        # Confirm exit checkbox
        self.confirm_exit_var = tk.BooleanVar(
            value=self.settings.get('confirm_exit', True)
        )
        ttk.Checkbutton(
            main_frame,
            text=self.strings.get('Labels', 'settings_confirm_exit', 'Confirm on exit'),
            variable=self.confirm_exit_var
        ).pack(fill=tk.X, pady=5, anchor=tk.W)
        
        # Remember path checkbox
        self.remember_path_var = tk.BooleanVar(
            value=self.settings.get('remember_path', True)
        )
        ttk.Checkbutton(
            main_frame,
            text=self.strings.get('Labels', 'settings_remember_path', 'Remember last folder'),
            variable=self.remember_path_var
        ).pack(fill=tk.X, pady=5, anchor=tk.W)
        
        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=20)
        
        ttk.Button(
            btn_frame,
            text=self.strings.get('Buttons', 'ok', 'OK'),
            command=self._save_settings
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            btn_frame,
            text=self.strings.get('Buttons', 'cancel', 'Cancel'),
            command=self.destroy
        ).pack(side=tk.RIGHT)
    
    def _save_settings(self) -> None:
        """Save settings and close dialog."""
        self.settings.set('language', self.lang_var.get())
        self.settings.set('auto_open_output', self.auto_open_var.get())
        self.settings.set('show_preview', self.show_preview_var.get())
        self.settings.set('confirm_exit', self.confirm_exit_var.get())
        self.settings.set('remember_path', self.remember_path_var.get())
        
        self.on_save()
        self.destroy()


class ManifestCleanerApp:
    """Main application class for VINAFCO Manifest Cleaner."""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_TITLE} {APP_VERSION}")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # Initialize managers
        self.settings = SettingsManager()
        self.strings = StringManager(self.settings.get('language', 'vi'))
        
        # Load config
        try:
            self.config = load_ini_config()
        except Exception as e:
            logger.error(f"Error loading config: {e}")
            messagebox.showerror("Error", f"Cannot load config.ini: {e}")
            self.config = {}
        
        # State variables
        self.current_file: Optional[Path] = None
        self.processed_df: Optional[pd.DataFrame] = None
        self.is_processing = False
        self.recent_expanded = True
        
        # Setup UI
        self._setup_styles()
        self._create_menu()
        self._create_main_ui()
        self._bind_shortcuts()
        
        # Setup logging
        setup_logging(self.config.get('log_file_name', 'vinafco_manifest.log'))
        
        logger.info(f"Application started: {APP_TITLE} {APP_VERSION}")
    
    def _setup_styles(self) -> None:
        """Configure ttk styles."""
        style = ttk.Style()
        
        # Configure button styles
        style.configure(
            'Primary.TButton',
            font=('Segoe UI', 10, 'bold'),
            padding=(15, 8)
        )
        
        style.configure(
            'Secondary.TButton',
            font=('Segoe UI', 9),
            padding=(10, 5)
        )
        
        # Configure label styles
        style.configure(
            'Header.TLabel',
            font=('Segoe UI', 12, 'bold')
        )
        
        style.configure(
            'Status.TLabel',
            font=('Segoe UI', 10)
        )
    
    def _create_menu(self) -> None:
        """Create application menu bar."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(
            label=self.strings.get('Menu', 'file', 'File'), 
            menu=file_menu
        )
        
        file_menu.add_command(
            label=self.strings.get('Menu', 'open_file', 'Open File...'),
            command=self._select_file,
            accelerator='Ctrl+O'
        )
        
        # Recent files submenu
        self.recent_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(
            label=self.strings.get('Menu', 'open_recent', 'Open Recent'),
            menu=self.recent_menu
        )
        self._update_recent_menu()
        
        file_menu.add_separator()
        file_menu.add_command(
            label=self.strings.get('Menu', 'exit', 'Exit'),
            command=self._on_closing,
            accelerator='Ctrl+Q'
        )
        
        # Settings menu
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(
            label=self.strings.get('Menu', 'settings', 'Settings'),
            menu=settings_menu
        )
        
        # Language submenu
        lang_menu = tk.Menu(settings_menu, tearoff=0)
        settings_menu.add_cascade(
            label=self.strings.get('Menu', 'language', 'Language'),
            menu=lang_menu
        )
        
        self.lang_var = tk.StringVar(value=self.settings.get('language', 'vi'))
        lang_menu.add_radiobutton(
            label='Tiếng Việt',
            variable=self.lang_var,
            value='vi',
            command=lambda: self._change_language('vi')
        )
        lang_menu.add_radiobutton(
            label='English',
            variable=self.lang_var,
            value='en',
            command=lambda: self._change_language('en')
        )
        
        settings_menu.add_separator()
        settings_menu.add_command(
            label=self.strings.get('Labels', 'settings_title', 'Settings...'),
            command=self._show_settings
        )
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(
            label=self.strings.get('Menu', 'help', 'Help'),
            menu=help_menu
        )
        
        help_menu.add_command(
            label=self.strings.get('Menu', 'keyboard_shortcuts', 'Keyboard Shortcuts'),
            command=self._show_shortcuts,
            accelerator='F1'
        )
        help_menu.add_separator()
        help_menu.add_command(
            label=self.strings.get('Menu', 'about', 'About'),
            command=self._show_about
        )
    
    def _create_main_ui(self) -> None:
        """Create main application UI."""
        # Main container
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Top section - File selection
        self._create_file_section(main_frame)
        
        # Middle section - Recent files (collapsible)
        self._create_recent_section(main_frame)
        
        # Action buttons
        self._create_action_buttons(main_frame)
        
        # Status section
        self._create_status_section(main_frame)
        
        # Statistics section
        self._create_stats_section(main_frame)
        
        # Footer
        self._create_footer(main_frame)
    
    def _create_file_section(self, parent: ttk.Frame) -> None:
        """Create file selection section."""
        file_frame = ttk.LabelFrame(
            parent, 
            text=self.strings.get('Labels', 'input_file', 'Input File'),
            padding=10
        )
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        # File path entry
        self.file_path_var = tk.StringVar(
            value=self.strings.get('Labels', 'no_file_selected', 'No file selected')
        )
        
        file_entry = ttk.Entry(
            file_frame, 
            textvariable=self.file_path_var,
            state='readonly',
            font=('Segoe UI', 10)
        )
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Browse button
        ttk.Button(
            file_frame,
            text=self.strings.get('Buttons', 'select_file', '📁 Select File'),
            command=self._select_file,
            style='Primary.TButton'
        ).pack(side=tk.RIGHT)
    
    def _create_recent_section(self, parent: ttk.Frame) -> None:
        """Create collapsible recent files section."""
        self.recent_frame = ttk.LabelFrame(
            parent,
            text=self.strings.get('Labels', 'recent_files', '📋 Recent Files'),
            padding=5
        )
        self.recent_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Toggle button
        self.toggle_btn = ttk.Button(
            self.recent_frame,
            text=self.strings.get('Labels', 'collapse_recent', '▲ Collapse'),
            command=self._toggle_recent,
            style='Secondary.TButton'
        )
        self.toggle_btn.pack(anchor=tk.E)
        
        # Recent files list
        self.recent_listbox_frame = ttk.Frame(self.recent_frame)
        self.recent_listbox_frame.pack(fill=tk.X)
        
        self.recent_listbox = tk.Listbox(
            self.recent_listbox_frame,
            height=4,
            font=('Segoe UI', 9),
            selectmode=tk.SINGLE
        )
        self.recent_listbox.pack(fill=tk.X, pady=5)
        self.recent_listbox.bind('<Double-Button-1>', self._open_recent_file)
        
        self._populate_recent_files()
    
    def _create_action_buttons(self, parent: ttk.Frame) -> None:
        """Create action buttons section."""
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=10)
        
        # Process button - simple without style
        self.process_btn = ttk.Button(
            btn_frame,
            text="🔄 Xử lý",
            command=self._start_processing
        )
        self.process_btn.pack(side=tk.LEFT, padx=(0, 10))
        logger.info(f"Process button created")
        
        # Preview button
        self.preview_btn = ttk.Button(
            btn_frame,
            text=self.strings.get('Buttons', 'preview', '👁️ Preview'),
            command=self._show_preview,
            style='Secondary.TButton',
            state=tk.DISABLED
        )
        self.preview_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Save button
        self.save_btn = ttk.Button(
            btn_frame,
            text=self.strings.get('Buttons', 'save', '💾 Save File'),
            command=self._save_file,
            style='Primary.TButton',
            state=tk.DISABLED
        )
        self.save_btn.pack(side=tk.LEFT)
    
    def _create_status_section(self, parent: ttk.Frame) -> None:
        """Create status display section."""
        status_frame = ttk.LabelFrame(
            parent,
            text=self.strings.get('Labels', 'status', 'Status'),
            padding=10
        )
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Status label
        self.status_var = tk.StringVar(
            value=self.strings.get('Labels', 'ready', 'Ready')
        )
        ttk.Label(
            status_frame,
            textvariable=self.status_var,
            style='Status.TLabel'
        ).pack(anchor=tk.W)
        
        # Progress bar
        self.progress = ttk.Progressbar(
            status_frame,
            mode='indeterminate',
            length=300
        )
        self.progress.pack(fill=tk.X, pady=(10, 0))
    
    def _create_stats_section(self, parent: ttk.Frame) -> None:
        """Create statistics display section."""
        stats_frame = ttk.LabelFrame(
            parent,
            text=self.strings.get('Labels', 'statistics', '📊 Statistics'),
            padding=10
        )
        stats_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Stats grid
        stats_grid = ttk.Frame(stats_frame)
        stats_grid.pack(fill=tk.X)
        
        # Container count
        ttk.Label(stats_grid, text=self.strings.get('Labels', 'container_count', 'Containers:')).grid(row=0, column=0, sticky=tk.W, padx=5)
        self.container_count_var = tk.StringVar(value="0")
        ttk.Label(stats_grid, textvariable=self.container_count_var, font=('Segoe UI', 10, 'bold')).grid(row=0, column=1, sticky=tk.W)
        
        # Full count
        ttk.Label(stats_grid, text=self.strings.get('Labels', 'full_count', 'Full (F):')).grid(row=0, column=2, sticky=tk.W, padx=(20, 5))
        self.full_count_var = tk.StringVar(value="0")
        ttk.Label(stats_grid, textvariable=self.full_count_var, font=('Segoe UI', 10, 'bold')).grid(row=0, column=3, sticky=tk.W)
        
        # Empty count
        ttk.Label(stats_grid, text=self.strings.get('Labels', 'empty_count', 'Empty (E):')).grid(row=0, column=4, sticky=tk.W, padx=(20, 5))
        self.empty_count_var = tk.StringVar(value="0")
        ttk.Label(stats_grid, textvariable=self.empty_count_var, font=('Segoe UI', 10, 'bold')).grid(row=0, column=5, sticky=tk.W)
        
        # SOC count
        ttk.Label(stats_grid, text=self.strings.get('Labels', 'soc_count', 'SOC:')).grid(row=1, column=0, sticky=tk.W, padx=5, pady=(5, 0))
        self.soc_count_var = tk.StringVar(value="0")
        ttk.Label(stats_grid, textvariable=self.soc_count_var, font=('Segoe UI', 10, 'bold')).grid(row=1, column=1, sticky=tk.W, pady=(5, 0))
        
        # COC count
        ttk.Label(stats_grid, text=self.strings.get('Labels', 'coc_count', 'COC:')).grid(row=1, column=2, sticky=tk.W, padx=(20, 5), pady=(5, 0))
        self.coc_count_var = tk.StringVar(value="0")
        ttk.Label(stats_grid, textvariable=self.coc_count_var, font=('Segoe UI', 10, 'bold')).grid(row=1, column=3, sticky=tk.W, pady=(5, 0))
    
    def _create_footer(self, parent: ttk.Frame) -> None:
        """Create footer section."""
        footer_frame = ttk.Frame(parent)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=10)
        
        ttk.Label(
            footer_frame,
            text=self.strings.get('UI', 'footer_author', 'Developed by: Tien-Tan Thuan Port'),
            font=('Segoe UI', 9, 'italic'),
            foreground='gray'
        ).pack()
    
    def _bind_shortcuts(self) -> None:
        """Bind keyboard shortcuts."""
        self.root.bind('<Control-o>', lambda e: self._select_file())
        self.root.bind('<Control-O>', lambda e: self._select_file())
        self.root.bind('<Control-s>', lambda e: self._save_file())
        self.root.bind('<Control-S>', lambda e: self._save_file())
        self.root.bind('<Control-p>', lambda e: self._show_preview())
        self.root.bind('<Control-P>', lambda e: self._show_preview())
        self.root.bind('<Control-r>', lambda e: self._start_processing())
        self.root.bind('<Control-R>', lambda e: self._start_processing())
        self.root.bind('<Control-q>', lambda e: self._on_closing())
        self.root.bind('<Control-Q>', lambda e: self._on_closing())
        self.root.bind('<F1>', lambda e: self._show_shortcuts())
        self.root.bind('<F5>', lambda e: self._refresh())
        
        # Window close handler
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _update_recent_menu(self) -> None:
        """Update recent files menu."""
        self.recent_menu.delete(0, tk.END)
        
        recent_files = self.settings.get('recent_files', [])
        
        if recent_files:
            for file_path in recent_files[:10]:
                self.recent_menu.add_command(
                    label=Path(file_path).name,
                    command=lambda p=file_path: self._open_file(p)
                )
            
            self.recent_menu.add_separator()
            self.recent_menu.add_command(
                label=self.strings.get('Menu', 'clear_recent', 'Clear List'),
                command=self._clear_recent_files
            )
        else:
            self.recent_menu.add_command(
                label=self.strings.get('Labels', 'no_file_selected', 'No recent files'),
                state=tk.DISABLED
            )
    
    def _populate_recent_files(self) -> None:
        """Populate recent files listbox."""
        self.recent_listbox.delete(0, tk.END)
        
        recent_files = self.settings.get('recent_files', [])
        for file_path in recent_files[:10]:
            display = f"📄 {Path(file_path).name}"
            self.recent_listbox.insert(tk.END, display)
    
    def _toggle_recent(self) -> None:
        """Toggle recent files section visibility."""
        if self.recent_expanded:
            self.recent_listbox_frame.pack_forget()
            self.toggle_btn.config(
                text=self.strings.get('Labels', 'expand_recent', '▼ Expand')
            )
            self.recent_expanded = False
        else:
            self.recent_listbox_frame.pack(fill=tk.X)
            self.toggle_btn.config(
                text=self.strings.get('Labels', 'collapse_recent', '▲ Collapse')
            )
            self.recent_expanded = True
    
    def _select_file(self) -> None:
        """Open file selection dialog."""
        initial_dir = self.settings.get('last_directory', '')
        
        if not initial_dir or not Path(initial_dir).exists():
            initial_dir = str(Path.home())
        
        file_types = [
            (self.strings.get('Dialogs', 'filter_excel', 'Excel Files'), "*.xls;*.xlsx;*.xlsm"),
            (self.strings.get('Dialogs', 'filter_all', 'All Files'), "*.*")
        ]
        
        file_path = filedialog.askopenfilename(
            title=self.strings.get('Dialogs', 'title_open', 'Select manifest file'),
            initialdir=initial_dir,
            filetypes=file_types
        )
        
        if file_path:
            self._open_file(file_path)
    
    def _open_file(self, file_path: str) -> None:
        """Open and prepare a file for processing."""
        path = Path(file_path)
        
        if not path.exists():
            messagebox.showerror(
                self.strings.get('Dialogs', 'error_title', 'Error'),
                self.strings.get('Messages', 'file_not_found', 'File not found!')
            )
            return
        
        if not is_valid_excel_file(path):
            messagebox.showerror(
                self.strings.get('Dialogs', 'error_title', 'Error'),
                self.strings.get('Messages', 'invalid_file', 'Invalid file!')
            )
            return
        
        self.current_file = path
        self.file_path_var.set(str(path))
        self.processed_df = None
        
        # Reset UI
        self.preview_btn.config(state=tk.DISABLED)
        self.save_btn.config(state=tk.DISABLED)
        self._reset_stats()
        self.status_var.set(self.strings.get('Labels', 'ready', 'Ready'))
        
        # Update settings
        self.settings.add_recent_file(str(path))
        if self.settings.get('remember_path', True):
            self.settings.set('last_directory', str(path.parent))
        
        # Update UI
        self._update_recent_menu()
        self._populate_recent_files()
        
        logger.info(f"File selected: {path}")
    
    def _open_recent_file(self, event) -> None:
        """Handle double-click on recent file."""
        selection = self.recent_listbox.curselection()
        if selection:
            idx = selection[0]
            recent_files = self.settings.get('recent_files', [])
            if idx < len(recent_files):
                self._open_file(recent_files[idx])
    
    def _clear_recent_files(self) -> None:
        """Clear recent files list."""
        self.settings.clear_recent_files()
        self._update_recent_menu()
        self._populate_recent_files()
        self.status_var.set(self.strings.get('Messages', 'recent_cleared', 'Recent files list cleared'))
    
    def _start_processing(self) -> None:
        """Start file processing in background thread."""
        logger.info("=== _start_processing CALLED ===")
        
        if not self.current_file:
            logger.warning("No file selected")
            messagebox.showwarning(
                self.strings.get('Dialogs', 'warning_title', 'Warning'),
                self.strings.get('Messages', 'select_file_first', 'Please select a file first!')
            )
            return
        
        if self.is_processing:
            logger.warning("Already processing")
            return
        
        logger.info(f"Starting processing for: {self.current_file}")
        
        self.is_processing = True
        self._set_ui_state(processing=True)
        self.status_var.set("Processing...")
        self.progress.start()
        
        # Start background thread
        thread = threading.Thread(target=self._process_file, daemon=True)
        thread.start()
    
    def _process_file(self) -> None:
        """Process file in background thread."""
        try:
            logger.info(f"Processing file: {self.current_file}")
            
            self._update_status(self.strings.get('Messages', 'loading_file', 'Loading file...'))
            
            # Initialize handler
            handler = VinafcoExcelHandler(self.config)
            
            # Read Excel file - returns tuple (df, header_row_idx)
            self._update_status(self.strings.get('Messages', 'reading_sheet', 'Reading sheet...'))
            raw_df, header_row_idx = handler.read_excel(str(self.current_file))
            
            logger.info(f"Read {len(raw_df)} rows from Excel, header at row {header_row_idx}")
            
            if raw_df is None or raw_df.empty:
                raise ValueError(self.strings.get('Messages', 'no_data', 'No data to process'))
            
            # Parse manifest data
            self._update_status(self.strings.get('Messages', 'analyzing_data', 'Analyzing data...'))
            parsed_data = handler.parse_manifest_data(raw_df, header_row_idx)
            
            logger.info(f"Parsed {len(parsed_data) if parsed_data else 0} container records")
            
            if not parsed_data:
                raise ValueError(self.strings.get('Messages', 'no_containers', 'No containers found'))
            
            # Build output DataFrame
            self._update_status(self.strings.get('Messages', 'writing_output', 'Building output...'))
            self.processed_df = handler.build_output_dataframe(parsed_data)
            
            logger.info(f"Built output DataFrame with {len(self.processed_df)} rows")
            
            # Calculate statistics
            stats = self._calculate_stats(self.processed_df)
            
            # Update UI on main thread
            self.root.after(0, lambda: self._processing_complete(stats))
            
        except Exception as e:
            logger.exception(f"Error processing file: {e}")
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self._processing_error(msg))
        
        finally:
            self.is_processing = False
            self.root.after(0, lambda: self.progress.stop())
    
    def _calculate_stats(self, df: pd.DataFrame) -> Dict[str, int]:
        """Calculate statistics from processed DataFrame."""
        stats = {
            'total': len(df),
            'full': 0,
            'empty': 0,
            'soc': 0,
            'coc': 0
        }
        
        logger.info(f"DataFrame columns: {df.columns.tolist()}")
        
        # Check F/E column
        if 'F/E' in df.columns:
            stats['full'] = len(df[df['F/E'] == 'F'])
            stats['empty'] = len(df[df['F/E'] == 'E'])
        
        # Check Operator column - Vietnamese name "Hãng khai thác"
        op_col = 'Hãng khai thác'
        if op_col in df.columns:
            soc_op = self.config.get('operator_soc', 'SVF')
            coc_op = self.config.get('operator_coc', 'VFC')
            unique_ops = df[op_col].unique().tolist()
            logger.info(f"Operator column unique values: {unique_ops}")
            stats['soc'] = len(df[df[op_col] == soc_op])
            stats['coc'] = len(df[df[op_col] == coc_op])
        
        logger.info(f"Stats calculated: {stats}")
        return stats
    
    def _update_stats(self, stats: Dict[str, int]) -> None:
        """Update statistics display."""
        self.container_count_var.set(str(stats['total']))
        self.full_count_var.set(str(stats['full']))
        self.empty_count_var.set(str(stats['empty']))
        self.soc_count_var.set(str(stats['soc']))
        self.coc_count_var.set(str(stats['coc']))
    
    def _reset_stats(self) -> None:
        """Reset statistics to zero."""
        self.container_count_var.set("0")
        self.full_count_var.set("0")
        self.empty_count_var.set("0")
        self.soc_count_var.set("0")
        self.coc_count_var.set("0")
    
    def _processing_complete(self, stats: Dict[str, int]) -> None:
        """Handle successful processing completion."""
        self._set_ui_state(processing=False)
        self._update_stats(stats)
        self.status_var.set(self.strings.get('Messages', 'processing_complete', 'Processing complete!'))
        
        self.preview_btn.config(state=tk.NORMAL)
        self.save_btn.config(state=tk.NORMAL)
        
        logger.info(f"Processing complete. {stats['total']} containers found.")
        
        # Auto show preview if enabled
        if self.settings.get('show_preview', True) and self.processed_df is not None:
            self._show_preview()
    
    def _processing_error(self, error_msg: str) -> None:
        """Handle processing error."""
        self._set_ui_state(processing=False)
        self.status_var.set(f"{self.strings.get('Labels', 'error', 'Error')}: {error_msg}")
        
        messagebox.showerror(
            self.strings.get('Dialogs', 'error_title', 'Error'),
            f"{self.strings.get('Messages', 'processing_error', 'Error during processing:')} {error_msg}"
        )
    
    def _set_ui_state(self, processing: bool) -> None:
        """Enable/disable UI elements during processing."""
        state = tk.DISABLED if processing else tk.NORMAL
        
        self.process_btn.config(state=state)
        
        if processing:
            self.status_var.set(self.strings.get('Labels', 'processing', 'Processing...'))
    
    def _update_status(self, message: str) -> None:
        """Update status from background thread."""
        self.root.after(0, lambda: self.status_var.set(message))
    
    def _show_preview(self) -> None:
        """Show preview window."""
        if self.processed_df is None or self.processed_df.empty:
            messagebox.showinfo(
                self.strings.get('Dialogs', 'info_title', 'Info'),
                self.strings.get('Messages', 'no_data', 'No data to preview')
            )
            return
        
        PreviewWindow(self.root, self.processed_df, self.strings)
    
    def _save_file(self) -> None:
        """Save processed file."""
        if self.processed_df is None or self.processed_df.empty:
            messagebox.showwarning(
                self.strings.get('Dialogs', 'warning_title', 'Warning'),
                self.strings.get('Messages', 'no_data', 'No data to save')
            )
            return
        
        # Generate output filename
        if self.current_file:
            default_name = f"{self.current_file.stem}_processed.xlsx"
            initial_dir = str(self.current_file.parent)
        else:
            default_name = "manifest_processed.xlsx"
            initial_dir = str(Path.home())
        
        file_types = [
            (self.strings.get('Dialogs', 'filter_excel', 'Excel Files'), "*.xlsx"),
        ]
        
        output_path = filedialog.asksaveasfilename(
            title=self.strings.get('Dialogs', 'title_save', 'Save processed file'),
            initialdir=initial_dir,
            initialfile=default_name,
            filetypes=file_types,
            defaultextension=".xlsx"
        )
        
        if not output_path:
            return
        
        try:
            self.status_var.set(self.strings.get('Messages', 'writing_output', 'Saving file...'))
            
            handler = VinafcoExcelHandler(self.config)
            handler.save_excel(self.processed_df, output_path)
            
            self.status_var.set(self.strings.get('Messages', 'save_success', 'File saved successfully!'))
            
            logger.info(f"File saved: {output_path}")
            
            # Auto open output folder
            if self.settings.get('auto_open_output', True):
                self._open_output_folder(output_path)
            
            messagebox.showinfo(
                self.strings.get('Dialogs', 'info_title', 'Success'),
                f"{self.strings.get('Messages', 'save_success', 'File saved successfully!')}\n{output_path}"
            )
            
        except Exception as e:
            logger.exception(f"Error saving file: {e}")
            self.status_var.set(f"{self.strings.get('Labels', 'error', 'Error')}: {e}")
            
            messagebox.showerror(
                self.strings.get('Dialogs', 'error_title', 'Error'),
                f"{self.strings.get('Messages', 'save_error', 'Error saving file:')} {e}"
            )
    
    def _open_output_folder(self, file_path: str) -> None:
        """Open the output folder in file explorer."""
        try:
            folder = Path(file_path).parent
            if sys.platform == 'win32':
                os.startfile(folder)
            elif sys.platform == 'darwin':
                subprocess.run(['open', folder])
            else:
                subprocess.run(['xdg-open', folder])
        except Exception as e:
            logger.warning(f"Could not open folder: {e}")
    
    def _show_settings(self) -> None:
        """Show settings dialog."""
        SettingsDialog(
            self.root,
            self.settings,
            self.strings,
            on_save=self._apply_settings
        )
    
    def _apply_settings(self) -> None:
        """Apply changed settings."""
        new_lang = self.settings.get('language', 'vi')
        if new_lang != self.strings.language:
            self._change_language(new_lang)
    
    def _change_language(self, language: str) -> None:
        """Change application language."""
        self.settings.set('language', language)
        self.strings.set_language(language)
        self.lang_var.set(language)
        
        # Notify user to restart for full effect
        messagebox.showinfo(
            self.strings.get('Dialogs', 'info_title', 'Info'),
            "Please restart the application for language changes to take full effect."
        )
    
    def _show_shortcuts(self) -> None:
        """Show keyboard shortcuts dialog."""
        shortcuts = [
            ("Ctrl+O", self.strings.get('Shortcuts', 'ctrl_o', 'Open file')),
            ("Ctrl+S", self.strings.get('Shortcuts', 'ctrl_s', 'Save file')),
            ("Ctrl+P", self.strings.get('Shortcuts', 'ctrl_p', 'Preview')),
            ("Ctrl+R", self.strings.get('Shortcuts', 'ctrl_r', 'Process')),
            ("Ctrl+Q", self.strings.get('Shortcuts', 'ctrl_q', 'Exit')),
            ("F1", self.strings.get('Shortcuts', 'f1', 'Help')),
            ("F5", self.strings.get('Shortcuts', 'f5', 'Refresh')),
        ]
        
        msg = "\n".join([f"{key}: {desc}" for key, desc in shortcuts])
        
        messagebox.showinfo(
            self.strings.get('Dialogs', 'shortcuts_title', 'Keyboard Shortcuts'),
            msg
        )
    
    def _show_about(self) -> None:
        """Show about dialog."""
        about_text = f"""{APP_TITLE}
{APP_VERSION}

{self.strings.get('About', 'description', 'VINAFCO Manifest Processing Tool')}

{self.strings.get('About', 'developer', 'Developed by: Tien-Tan Thuan Port')}

{self.strings.get('About', 'copyright', '© 2026 All rights reserved')}
"""
        
        messagebox.showinfo(
            self.strings.get('Dialogs', 'about_title', 'About'),
            about_text
        )
    
    def _refresh(self) -> None:
        """Refresh application state."""
        self._reset_stats()
        self.processed_df = None
        self.preview_btn.config(state=tk.DISABLED)
        self.save_btn.config(state=tk.DISABLED)
        self.status_var.set(self.strings.get('Labels', 'ready', 'Ready'))
    
    def _on_closing(self) -> None:
        """Handle window close event."""
        if self.settings.get('confirm_exit', True):
            if not messagebox.askyesno(
                self.strings.get('Dialogs', 'confirm_title', 'Confirm'),
                self.strings.get('Messages', 'confirm_exit', 'Are you sure you want to exit?')
            ):
                return
        
        # Save window geometry
        self.settings.set('window_geometry', self.root.geometry())
        
        logger.info("Application closed")
        self.root.destroy()
    
    def run(self) -> None:
        """Start the application main loop."""
        # Restore window geometry
        geometry = self.settings.get('window_geometry', '900x700')
        self.root.geometry(geometry)
        
        self.root.mainloop()


def create_app() -> ManifestCleanerApp:
    """Factory function to create application instance."""
    root = tk.Tk()
    app = ManifestCleanerApp(root)
    return app


if __name__ == "__main__":
    app = create_app()
    app.run()
