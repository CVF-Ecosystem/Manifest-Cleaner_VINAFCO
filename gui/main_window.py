"""Main Window GUI for VINAFCO Manifest Cleaner v2.0.

Clean, compact design based on VOSCO architecture.
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
import winsound
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd

from config.constants import APP_TITLE, APP_VERSION, UI_TEXT
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


# Color scheme - VOSCO style
COLORS = {
    "bg_main": "#f5f5f5",
    "bg_card": "#ffffff",
    "bg_hover": "#e8f4f8",
    "primary": "#1976d2",
    "primary_hover": "#1565c0",
    "success": "#4caf50",
    "danger": "#ef5350",
    "text": "#333333",
    "text_muted": "#666666",
    "border": "#e0e0e0",
}


class AppConfig:
    """Manages application configuration and recent files."""
    
    def __init__(self, config_file: str = "app_config.json"):
        self.config_path = get_resource_path(config_file)
        self.language = "vi"
        self.auto_open_output = True
        self.recent_files: List[Dict[str, Any]] = []
        self.max_recent = 5
        self._load()
    
    def _load(self) -> None:
        """Load config from JSON file."""
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.language = data.get('language', 'vi')
                    self.auto_open_output = data.get('auto_open_output', True)
                    self.recent_files = data.get('recent_files', [])
            except Exception as e:
                logger.error(f"Error loading config: {e}")
    
    def save(self) -> None:
        """Save config to JSON file."""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump({
                    'language': self.language,
                    'auto_open_output': self.auto_open_output,
                    'recent_files': self.recent_files
                }, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"Error saving config: {e}")
    
    def add_recent_file(self, file_path: str) -> None:
        """Add file to recent list."""
        entry = {
            'path': file_path,
            'name': Path(file_path).name,
            'time': pd.Timestamp.now().isoformat()
        }
        
        # Remove if exists - support both dict and string format
        new_list = []
        for f in self.recent_files:
            if isinstance(f, dict):
                if f.get('path') != file_path:
                    new_list.append(f)
            else:
                if str(f) != file_path:
                    pass  # Remove old string format entries
        self.recent_files = new_list
        
        # Add to front
        self.recent_files.insert(0, entry)
        
        # Trim
        self.recent_files = self.recent_files[:self.max_recent]
        self.save()


class StringManager:
    """Manages bilingual UI strings."""
    
    def __init__(self, language: str = 'vi'):
        self.language = language
        self.strings: Dict[str, str] = {}
        self._load_strings()
    
    def _load_strings(self) -> None:
        """Load strings from INI files."""
        lang_file = f"strings_{self.language}.ini"
        lang_path = get_resource_path(lang_file)
        
        if not lang_path.exists():
            lang_path = get_resource_path("strings_vi.ini")
        
        if lang_path.exists():
            parser = configparser.ConfigParser(interpolation=None)
            parser.read(lang_path, encoding='utf-8')
            
            for section in parser.sections():
                for key, value in parser[section].items():
                    self.strings[key] = value
    
    def get(self, key: str, default: str = "") -> str:
        """Get a string value."""
        return self.strings.get(key, default or key)
    
    def set_language(self, language: str) -> None:
        """Change language."""
        self.language = language
        self._load_strings()


# Default texts (Vietnamese)
DEFAULT_TEXTS = {
    "app_title": "VINAFCO Manifest Cleaner",
    "app_subtitle": "Xử lý nhanh manifest container",
    "btn_choose_file": "📁 Chọn file Excel",
    "btn_process": "🔄 Xử lý",
    "btn_save": "💾 Lưu file",
    "btn_preview": "👁️ Xem trước",
    "recent_files": "📂 Gần đây:",
    "no_recent": "Chưa có file nào",
    "status_ready": "Sẵn sàng",
    "status_processing": "Đang xử lý...",
    "status_done": "Hoàn thành!",
    "status_error": "Lỗi:",
    "stats_containers": "Container:",
    "stats_full": "Full:",
    "stats_empty": "Empty:",
    "stats_soc": "SOC:",
    "stats_coc": "COC:",
    "footer": "© 2025 Tiền Tấn - Thuận Port",
    "settings": "⚙️",
    "language": "Ngôn ngữ",
    "auto_open": "Tự động mở file",
    "about": "Thông tin",
    "shortcuts": "Phím tắt",
}


class ManifestCleanerApp:
    """Main application class - Compact design."""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_TITLE} {APP_VERSION}")
        self.root.configure(bg=COLORS["bg_main"])
        
        # Initialize
        self.config = AppConfig()
        self.texts = DEFAULT_TEXTS.copy()
        self._load_language_strings()
        
        # Load INI config
        try:
            self.ini_config = load_ini_config()
        except Exception as e:
            logger.error(f"Error loading config.ini: {e}")
            self.ini_config = {}
        
        # State
        self.current_file: Optional[Path] = None
        self.processed_df: Optional[pd.DataFrame] = None
        self.is_processing = False
        
        # Setup
        self._setup_styles()
        self._build_ui()
        self._bind_shortcuts()
        self._center_window(620, 580)
        
        setup_logging(self.ini_config.get('log_file_name', 'vinafco_manifest.log'))
        logger.info(f"App started: {APP_TITLE} {APP_VERSION}")
    
    def _load_language_strings(self) -> None:
        """Load language strings from INI file."""
        lang_file = f"strings_{self.config.language}.ini"
        lang_path = get_resource_path(lang_file)
        
        if lang_path.exists():
            parser = configparser.ConfigParser(interpolation=None)
            parser.read(lang_path, encoding='utf-8')
            
            # Load all UI strings from [UI] section
            if parser.has_section('UI'):
                for key, value in parser.items('UI'):
                    self.texts[key] = value
            
            logger.info(f"Loaded language: {self.config.language}")
    
    def _center_window(self, width: int, height: int) -> None:
        """Center window on screen."""
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(500, 400)
    
    def _setup_styles(self) -> None:
        """Configure ttk styles."""
        style = ttk.Style()
        style.theme_use("clam")
        
        # Card frame
        style.configure(
            "Card.TFrame",
            background=COLORS["bg_card"]
        )
        
        # Primary button
        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(20, 10)
        )
        
        # Secondary button
        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 9),
            padding=(12, 6)
        )
        
        # Title label
        style.configure(
            "Title.TLabel",
            font=("Segoe UI", 16, "bold"),
            background=COLORS["bg_card"],
            foreground=COLORS["text"]
        )
        
        # Subtitle label
        style.configure(
            "Subtitle.TLabel",
            font=("Segoe UI", 9),
            background=COLORS["bg_card"],
            foreground=COLORS["text_muted"]
        )
        
        # Footer label
        style.configure(
            "Footer.TLabel",
            font=("Segoe UI", 8),
            background=COLORS["bg_card"],
            foreground=COLORS["text_muted"]
        )
        
        # Stats label
        style.configure(
            "Stats.TLabel",
            font=("Segoe UI", 10),
            background=COLORS["bg_card"],
            foreground=COLORS["text"]
        )
        
        style.configure(
            "StatsValue.TLabel",
            font=("Segoe UI", 11, "bold"),
            background=COLORS["bg_card"],
            foreground=COLORS["primary"]
        )
    
    def _build_ui(self) -> None:
        """Build main UI - compact card layout."""
        # Toolbar
        self._build_toolbar()
        
        # Card frame
        self.card_frame = ttk.Frame(self.root, style="Card.TFrame", padding=20)
        self.card_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        
        # Header
        self._build_header()
        
        # Action section
        self._build_action_section()
        
        # Recent files
        self._build_recent_section()
        
        # Status
        self._build_status_section()
        
        # Statistics
        self._build_stats_section()
        
        # Footer
        self._build_footer()
    
    def _build_toolbar(self) -> None:
        """Build compact toolbar."""
        toolbar = tk.Frame(self.root, bg=COLORS["bg_main"])
        toolbar.pack(fill=tk.X, padx=15, pady=10)
        
        # App icon
        tk.Label(
            toolbar,
            text="📦",
            font=("Segoe UI", 14),
            bg=COLORS["bg_main"]
        ).pack(side=tk.LEFT)
        
        # VINAFCO text
        tk.Label(
            toolbar,
            text="VINAFCO",
            font=("Segoe UI", 11, "bold"),
            bg=COLORS["bg_main"],
            fg=COLORS["primary"]
        ).pack(side=tk.LEFT, padx=(5, 0))
        
        # Settings button (dropdown)
        self.settings_btn = tk.Menubutton(
            toolbar,
            text="⚙️",
            font=("Segoe UI", 12),
            bg=COLORS["bg_main"],
            relief=tk.FLAT,
            cursor="hand2"
        )
        self.settings_btn.pack(side=tk.RIGHT)
        
        # Settings menu
        settings_menu = tk.Menu(self.settings_btn, tearoff=0)
        self.settings_btn.config(menu=settings_menu)
        
        # Language submenu
        lang_menu = tk.Menu(settings_menu, tearoff=0)
        settings_menu.add_cascade(label="🌐 " + self.texts.get("language", "Language"), menu=lang_menu)
        
        self.lang_var = tk.StringVar(value=self.config.language)
        lang_menu.add_radiobutton(
            label="Tiếng Việt",
            variable=self.lang_var,
            value="vi",
            command=lambda: self._change_language("vi")
        )
        lang_menu.add_radiobutton(
            label="English",
            variable=self.lang_var,
            value="en",
            command=lambda: self._change_language("en")
        )
        
        # Auto-open toggle
        self.auto_open_var = tk.BooleanVar(value=self.config.auto_open_output)
        settings_menu.add_checkbutton(
            label="📂 " + self.texts.get("auto_open", "Auto open file"),
            variable=self.auto_open_var,
            command=self._toggle_auto_open
        )
        
        settings_menu.add_separator()
        settings_menu.add_command(
            label="⌨️ " + self.texts.get("shortcuts", "Shortcuts") + " (F1)",
            command=self._show_shortcuts
        )
        settings_menu.add_command(
            label="ℹ️ " + self.texts.get("about", "About"),
            command=self._show_about
        )
    
    def _build_header(self) -> None:
        """Build header with icon and title."""
        header_frame = ttk.Frame(self.card_frame, style="Card.TFrame")
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # App icon
        ttk.Label(
            header_frame,
            text="📦",
            font=("Segoe UI", 28),
            background=COLORS["bg_card"]
        ).pack()
        
        # Subtitle only - cleaner look
        self.subtitle_label = ttk.Label(
            header_frame,
            text=self.texts.get("app_subtitle", "Manifest Cleaner"),
            style="Subtitle.TLabel"
        )
        self.subtitle_label.pack(pady=(5, 0))
    
    def _build_action_section(self) -> None:
        """Build action buttons."""
        action_frame = ttk.Frame(self.card_frame, style="Card.TFrame")
        action_frame.pack(fill=tk.X, pady=10)
        
        # Choose file button - centered
        self.btn_choose = ttk.Button(
            action_frame,
            text=self.texts.get("btn_choose_file", "📁 Chọn file Excel"),
            style="Primary.TButton",
            command=self._on_choose_file
        )
        self.btn_choose.pack()
        
        # File name display (hidden initially)
        self.file_label_var = tk.StringVar(value="")
        self.file_label = tk.Label(
            action_frame,
            textvariable=self.file_label_var,
            font=("Segoe UI", 9),
            bg=COLORS["bg_card"],
            fg=COLORS["text_muted"]
        )
        # Don't pack yet - show when file selected
        
        # Action buttons row (hidden initially)
        self.action_btns_frame = ttk.Frame(action_frame, style="Card.TFrame")
        
        self.btn_process = ttk.Button(
            self.action_btns_frame,
            text=self.texts.get("btn_process", "🔄 Xử lý"),
            style="Secondary.TButton",
            command=self._on_process
        )
        self.btn_process.pack(side=tk.LEFT, padx=3)
        
        self.btn_preview = ttk.Button(
            self.action_btns_frame,
            text=self.texts.get("btn_preview", "👁️ Xem trước"),
            style="Secondary.TButton",
            command=self._on_preview,
            state=tk.DISABLED
        )
        self.btn_preview.pack(side=tk.LEFT, padx=3)
        
        self.btn_save = ttk.Button(
            self.action_btns_frame,
            text=self.texts.get("btn_save", "💾 Lưu file"),
            style="Secondary.TButton",
            command=self._on_save,
            state=tk.DISABLED
        )
        self.btn_save.pack(side=tk.LEFT, padx=3)
    
    def _build_recent_section(self) -> None:
        """Build collapsible recent files section."""
        self.recent_frame = ttk.Frame(self.card_frame, style="Card.TFrame")
        self.recent_frame.pack(fill=tk.X, pady=8)
        
        self.recent_expanded = False
        
        # Header row
        header_frame = ttk.Frame(self.recent_frame, style="Card.TFrame")
        header_frame.pack(fill=tk.X)
        
        self.recent_toggle = tk.Label(
            header_frame,
            text="▶",
            fg=COLORS["text_muted"],
            bg=COLORS["bg_card"],
            font=("Segoe UI", 9),
            cursor="hand2"
        )
        self.recent_toggle.pack(side=tk.LEFT)
        
        self.recent_title = tk.Label(
            header_frame,
            text=self.texts.get("recent_files", "📂 Gần đây:"),
            fg=COLORS["text_muted"],
            bg=COLORS["bg_card"],
            font=("Segoe UI", 9),
            cursor="hand2"
        )
        self.recent_title.pack(side=tk.LEFT)
        
        # Count badge
        count = len(self.config.recent_files)
        self.recent_count = tk.Label(
            header_frame,
            text=f"({count})" if count else "",
            fg=COLORS["text_muted"],
            bg=COLORS["bg_card"],
            font=("Segoe UI", 8)
        )
        self.recent_count.pack(side=tk.LEFT, padx=(3, 0))
        
        # Bind clicks
        for widget in [self.recent_toggle, self.recent_title, self.recent_count]:
            widget.bind("<Button-1>", lambda e: self._toggle_recent())
        
        # File list (hidden)
        self.recent_list_frame = ttk.Frame(self.recent_frame, style="Card.TFrame")
        self._update_recent_display()
    
    def _toggle_recent(self) -> None:
        """Toggle recent files section."""
        self.recent_expanded = not self.recent_expanded
        
        if self.recent_expanded:
            self.recent_toggle.config(text="▼")
            self.recent_list_frame.pack(fill=tk.X, pady=(5, 0))
        else:
            self.recent_toggle.config(text="▶")
            self.recent_list_frame.pack_forget()
    
    def _update_recent_display(self) -> None:
        """Update recent files display."""
        # Clear
        for widget in self.recent_list_frame.winfo_children():
            widget.destroy()
        
        count = len(self.config.recent_files)
        self.recent_count.config(text=f"({count})" if count else "")
        
        if not self.config.recent_files:
            tk.Label(
                self.recent_list_frame,
                text=self.texts.get("no_recent", "Chưa có file nào"),
                fg=COLORS["text_muted"],
                bg=COLORS["bg_card"],
                font=("Segoe UI", 8)
            ).pack(anchor=tk.W, padx=(15, 0))
            return
        
        for entry in self.config.recent_files[:5]:
            # Support both dict and string format
            if isinstance(entry, dict):
                file_path = entry.get('path', '')
                file_name = entry.get('name', Path(file_path).name)
            else:
                file_path = str(entry)
                file_name = Path(file_path).name
            
            file_btn = tk.Label(
                self.recent_list_frame,
                text=f"  📄 {file_name}",
                fg=COLORS["primary"],
                bg=COLORS["bg_card"],
                font=("Segoe UI", 9),
                cursor="hand2"
            )
            file_btn.pack(anchor=tk.W, pady=1)
            file_btn.bind("<Button-1>", lambda e, p=file_path: self._open_recent(p))
            file_btn.bind("<Enter>", lambda e, w=file_btn: w.config(fg=COLORS["primary_hover"]))
            file_btn.bind("<Leave>", lambda e, w=file_btn: w.config(fg=COLORS["primary"]))
    
    def _build_status_section(self) -> None:
        """Build status display."""
        status_frame = ttk.Frame(self.card_frame, style="Card.TFrame")
        status_frame.pack(fill=tk.X, pady=10)
        
        self.status_var = tk.StringVar(value=self.texts.get("status_ready", "Sẵn sàng"))
        self.status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            fg=COLORS["text_muted"],
            bg=COLORS["bg_card"],
            font=("Segoe UI", 9)
        )
        self.status_label.pack()
        
        # Progress bar (hidden)
        self.progress = ttk.Progressbar(
            status_frame,
            mode='indeterminate',
            length=200
        )
    
    def _build_stats_section(self) -> None:
        """Build statistics display."""
        self.stats_frame = ttk.Frame(self.card_frame, style="Card.TFrame")
        # Don't pack yet - show after processing
        
        # Stats grid
        stats_grid = ttk.Frame(self.stats_frame, style="Card.TFrame")
        stats_grid.pack()
        
        # Row 1: Containers, Full, Empty
        labels_row1 = [
            ("stats_containers", "container_count"),
            ("stats_full", "full_count"),
            ("stats_empty", "empty_count")
        ]
        
        self.stats_vars = {}
        
        for col, (label_key, var_key) in enumerate(labels_row1):
            ttk.Label(
                stats_grid,
                text=self.texts.get(label_key, label_key),
                style="Stats.TLabel"
            ).grid(row=0, column=col*2, padx=(10, 2))
            
            self.stats_vars[var_key] = tk.StringVar(value="0")
            ttk.Label(
                stats_grid,
                textvariable=self.stats_vars[var_key],
                style="StatsValue.TLabel"
            ).grid(row=0, column=col*2+1, padx=(0, 10))
        
        # Row 2: SOC, COC
        labels_row2 = [
            ("stats_soc", "soc_count"),
            ("stats_coc", "coc_count")
        ]
        
        for col, (label_key, var_key) in enumerate(labels_row2):
            ttk.Label(
                stats_grid,
                text=self.texts.get(label_key, label_key),
                style="Stats.TLabel"
            ).grid(row=1, column=col*2, padx=(10, 2), pady=(5, 0))
            
            self.stats_vars[var_key] = tk.StringVar(value="0")
            ttk.Label(
                stats_grid,
                textvariable=self.stats_vars[var_key],
                style="StatsValue.TLabel"
            ).grid(row=1, column=col*2+1, padx=(0, 10), pady=(5, 0))
    
    def _build_footer(self) -> None:
        """Build footer."""
        self.footer_label = ttk.Label(
            self.card_frame,
            text=self.texts.get("footer_author", "Developed by: Tien-Tan Thuan Port"),
            style="Footer.TLabel"
        )
        self.footer_label.pack(side=tk.BOTTOM, pady=(10, 0))
    
    def _bind_shortcuts(self) -> None:
        """Bind keyboard shortcuts."""
        self.root.bind('<Control-o>', lambda e: self._on_choose_file())
        self.root.bind('<Control-O>', lambda e: self._on_choose_file())
        self.root.bind('<Control-r>', lambda e: self._on_process())
        self.root.bind('<Control-R>', lambda e: self._on_process())
        self.root.bind('<Control-s>', lambda e: self._on_save())
        self.root.bind('<Control-S>', lambda e: self._on_save())
        self.root.bind('<Control-p>', lambda e: self._on_preview())
        self.root.bind('<Control-P>', lambda e: self._on_preview())
        self.root.bind('<F1>', lambda e: self._show_shortcuts())
        self.root.bind('<Control-q>', lambda e: self._on_close())
        self.root.bind('<Control-Q>', lambda e: self._on_close())
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
    
    # ========== Event Handlers ==========
    
    def _on_choose_file(self) -> None:
        """Handle file selection."""
        initial_dir = ""
        if self.config.recent_files:
            entry = self.config.recent_files[0]
            # Support both dict and string format
            if isinstance(entry, dict):
                last_path = entry.get('path', '')
            else:
                last_path = str(entry)
            if last_path:
                initial_dir = str(Path(last_path).parent)
        
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialdir=initial_dir
        )
        
        if file_path:
            self._load_file(file_path)
    
    def _load_file(self, file_path: str) -> None:
        """Load selected file and auto-process."""
        self.current_file = Path(file_path)
        
        if not self.current_file.exists():
            messagebox.showerror("Error", f"File not found: {file_path}")
            return
        
        # Update UI
        self.file_label_var.set(f"📄 {self.current_file.name}")
        self.file_label.pack(pady=(5, 0))
        self.action_btns_frame.pack(pady=(10, 0))
        
        # Reset state
        self.processed_df = None
        self.btn_preview.config(state=tk.DISABLED)
        self.btn_save.config(state=tk.DISABLED)
        self.stats_frame.pack_forget()
        
        # Update status
        size = format_file_size(self.current_file.stat().st_size)
        self.status_var.set(f"Đã chọn: {self.current_file.name} ({size})")
        self.status_label.config(fg=COLORS["text"])
        
        # Add to recent
        self.config.add_recent_file(file_path)
        self._update_recent_display()
        
        logger.info(f"File loaded: {file_path}")
        
        # Auto-process after loading (like VIMC)
        self.root.after(100, self._on_process)
    
    def _open_recent(self, file_path: str) -> None:
        """Open recent file."""
        if Path(file_path).exists():
            self._load_file(file_path)
        else:
            messagebox.showwarning("Warning", f"File not found:\n{file_path}")
            # Remove from recent
            self.config.recent_files = [
                f for f in self.config.recent_files 
                if f.get('path') != file_path
            ]
            self.config.save()
            self._update_recent_display()
    
    def _on_process(self) -> None:
        """Handle process button click."""
        if not self.current_file:
            messagebox.showwarning("Warning", "Please select a file first")
            return
        
        if self.is_processing:
            return
        
        self.is_processing = True
        self.btn_process.config(state=tk.DISABLED)
        self.status_var.set(self.texts.get("status_processing", "Đang xử lý..."))
        self.status_label.config(fg=COLORS["primary"])
        self.progress.pack(pady=(5, 0))
        self.progress.start(10)
        
        # Process in thread
        thread = threading.Thread(target=self._process_file)
        thread.start()
    
    def _process_file(self) -> None:
        """Process file in background thread."""
        try:
            # Pre-validation checks
            file_path = Path(self.current_file)
            
            # Check file exists
            if not file_path.exists():
                raise FileNotFoundError(f"File not found: {file_path.name}")
            
            # Check file extension
            valid_extensions = ['.xlsx', '.xls', '.xlsm']
            if file_path.suffix.lower() not in valid_extensions:
                raise ValueError(f"Invalid file format. Supported: {', '.join(valid_extensions)}")
            
            # Check file size (warn if > 10MB, reject if > 50MB)
            file_size_mb = file_path.stat().st_size / (1024 * 1024)
            if file_size_mb > 50:
                raise ValueError(f"File too large ({file_size_mb:.1f}MB). Maximum: 50MB")
            
            # Check if file is readable (not locked by another process)
            try:
                with open(file_path, 'rb') as f:
                    f.read(1)  # Try reading 1 byte
            except PermissionError:
                raise PermissionError("File is locked by another program. Please close it and try again.")
            except IOError as e:
                raise IOError(f"Cannot read file: {e}")
            
            handler = VinafcoExcelHandler(self.ini_config)
            
            # Read file
            result = handler.read_excel(self.current_file)
            
            if isinstance(result, tuple):
                df, header_row_idx = result
            else:
                df = result
                header_row_idx = 0
            
            if df is None or df.empty:
                raise ValueError("Cannot read data from file")
            
            # Parse manifest data
            intermediate_data = handler.parse_manifest_data(df, header_row_idx)
            
            if not intermediate_data:
                raise ValueError("No valid data found in file")
            
            # Build output DataFrame
            self.processed_df = handler.build_output_dataframe(intermediate_data)
            
            if self.processed_df is None or self.processed_df.empty:
                raise ValueError("Failed to build output table")
            
            # Update UI
            self.root.after(0, self._on_process_complete)
            
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Processing error: {error_msg}")
            self.root.after(0, lambda msg=error_msg: self._on_process_error(msg))
    
    def _play_sound(self, sound_type: str) -> None:
        """Play notification sound."""
        try:
            if sound_type == "success":
                winsound.MessageBeep(winsound.MB_OK)
            elif sound_type == "error":
                winsound.MessageBeep(winsound.MB_ICONHAND)
        except Exception:
            pass
    
    def _on_process_complete(self) -> None:
        """Handle processing complete - auto show save dialog."""
        self.is_processing = False
        self.progress.stop()
        self.progress.pack_forget()
        
        self.btn_process.config(state=tk.NORMAL)
        self.btn_preview.config(state=tk.NORMAL)
        self.btn_save.config(state=tk.NORMAL)
        
        self.status_var.set(self.texts.get("status_done", "Hoàn thành!"))
        self.status_label.config(fg=COLORS["success"])
        
        # Calculate stats
        self._calculate_stats()
        self.stats_frame.pack(fill=tk.X, pady=10)
        
        # Play success sound
        self._play_sound("success")
        
        logger.info("Processing complete")
        
        # Auto show save dialog (like VIMC)
        self.root.after(100, self._on_save)
    
    def _on_process_error(self, error_msg: str) -> None:
        """Handle processing error."""
        self.is_processing = False
        self.progress.stop()
        self.progress.pack_forget()
        
        self.btn_process.config(state=tk.NORMAL)
        
        self.status_var.set(f"{self.texts.get('status_error', 'Lỗi:')} {error_msg}")
        self.status_label.config(fg=COLORS["danger"])
        
        # Play error sound
        self._play_sound("error")
        
        messagebox.showerror("Error", f"Processing failed:\n{error_msg}")
    
    def _calculate_stats(self) -> None:
        """Calculate and display statistics."""
        if self.processed_df is None:
            return
        
        df = self.processed_df
        
        # Container count
        total = len(df)
        self.stats_vars["container_count"].set(str(total))
        
        # Full/Empty
        fe_col = None
        for col in df.columns:
            if 'F/E' in col.upper() or col.upper() in ['F/E', 'FE', 'FULL/EMPTY']:
                fe_col = col
                break
        
        full_count = 0
        empty_count = 0
        if fe_col:
            full_count = len(df[df[fe_col].astype(str).str.upper() == 'F'])
            empty_count = len(df[df[fe_col].astype(str).str.upper() == 'E'])
        
        self.stats_vars["full_count"].set(str(full_count))
        self.stats_vars["empty_count"].set(str(empty_count))
        
        # SOC/COC - find operator column
        operator_col = None
        for col in df.columns:
            col_upper = col.upper()
            if 'HÃNG KHAI THÁC' in col.upper() or 'HANG KHAI THAC' in col.upper():
                operator_col = col
                break
            elif 'OPERATOR' in col_upper or 'OPR' in col_upper:
                operator_col = col
                break
        
        soc_count = 0
        coc_count = 0
        
        if operator_col:
            # SVF = SOC, VFC = COC
            for val in df[operator_col].astype(str).str.upper():
                if 'SVF' in val:
                    soc_count += 1
                elif 'VFC' in val:
                    coc_count += 1
        
        self.stats_vars["soc_count"].set(str(soc_count))
        self.stats_vars["coc_count"].set(str(coc_count))
    
    def _on_preview(self) -> None:
        """Show preview window."""
        if self.processed_df is None:
            messagebox.showwarning("Warning", "No data to preview")
            return
        
        PreviewWindow(self.root, self.processed_df)
    
    def _on_save(self) -> None:
        """Save processed file."""
        if self.processed_df is None:
            messagebox.showwarning("Warning", "No data to save")
            return
        
        # Generate output filename
        if self.current_file:
            output_name = f"{self.current_file.stem}_processed.xlsx"
            output_dir = self.current_file.parent
        else:
            output_name = "manifest_processed.xlsx"
            output_dir = Path.cwd()
        
        output_path = filedialog.asksaveasfilename(
            title="Save Processed File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=output_name,
            initialdir=output_dir
        )
        
        if not output_path:
            return
        
        try:
            handler = VinafcoExcelHandler(self.ini_config)
            handler.save_excel(self.processed_df, Path(output_path))
            
            self.status_var.set(f"Đã lưu: {Path(output_path).name}")
            self.status_label.config(fg=COLORS["success"])
            
            # Auto open
            if self.config.auto_open_output:
                if sys.platform == 'win32':
                    os.startfile(output_path)
                elif sys.platform == 'darwin':
                    subprocess.run(['open', output_path])
                else:
                    subprocess.run(['xdg-open', output_path])
            
            logger.info(f"File saved: {output_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Cannot save file:\n{e}")
            logger.error(f"Save error: {e}")
    
    def _change_language(self, lang: str) -> None:
        """Change UI language - apply immediately."""
        if lang == self.config.language:
            return
        
        self.config.language = lang
        self.config.save()
        
        # Reload language strings
        self._load_language_strings()
        
        # Update UI elements
        self._refresh_ui_texts()
    
    def _refresh_ui_texts(self) -> None:
        """Refresh all UI texts after language change."""
        # Update subtitle
        if hasattr(self, 'subtitle_label'):
            self.subtitle_label.config(text=self.texts.get("app_subtitle", "Manifest Cleaner"))
        
        # Update buttons
        if hasattr(self, 'btn_choose'):
            self.btn_choose.config(text=self.texts.get("btn_choose_file", "📂 Chọn File Manifest"))
        if hasattr(self, 'btn_process'):
            self.btn_process.config(text=self.texts.get("btn_process", "🔄 Xử lý"))
        if hasattr(self, 'btn_preview'):
            self.btn_preview.config(text=self.texts.get("btn_preview", "👁️ Xem trước"))
        if hasattr(self, 'btn_save'):
            self.btn_save.config(text=self.texts.get("btn_save", "💾 Lưu file"))
        
        # Update recent files section
        if hasattr(self, 'recent_title'):
            self.recent_title.config(text=self.texts.get("recent_files", "📂 Gần đây:"))
        
        # Update status if ready
        if hasattr(self, 'status_var') and not self.is_processing:
            if self.processed_df is None:
                self.status_var.set(self.texts.get("status_ready", "Sẵn sàng"))
        
        # Update footer
        if hasattr(self, 'footer_label'):
            self.footer_label.config(text=self.texts.get("footer_author", "Developed by: Tien-Tan Thuan Port"))
        
        # Rebuild settings menu with new language
        self._rebuild_settings_menu()
    
    def _rebuild_settings_menu(self) -> None:
        """Rebuild settings menu with current language."""
        if not hasattr(self, 'settings_btn'):
            return
        
        # Create new menu
        settings_menu = tk.Menu(self.settings_btn, tearoff=0)
        self.settings_btn.config(menu=settings_menu)
        
        # Language submenu
        lang_menu = tk.Menu(settings_menu, tearoff=0)
        settings_menu.add_cascade(
            label="🌐 " + self.texts.get("menu_language", "Ngôn ngữ"), 
            menu=lang_menu
        )
        
        self.lang_var = tk.StringVar(value=self.config.language)
        lang_menu.add_radiobutton(
            label="Tiếng Việt",
            variable=self.lang_var,
            value="vi",
            command=lambda: self._change_language("vi")
        )
        lang_menu.add_radiobutton(
            label="English",
            variable=self.lang_var,
            value="en",
            command=lambda: self._change_language("en")
        )
        
        # Auto-open toggle
        self.auto_open_var = tk.BooleanVar(value=self.config.auto_open_output)
        settings_menu.add_checkbutton(
            label="📂 " + self.texts.get("menu_auto_open", "Tự động mở file"),
            variable=self.auto_open_var,
            command=self._toggle_auto_open
        )
        
        settings_menu.add_separator()
        settings_menu.add_command(
            label="⌨️ " + self.texts.get("menu_shortcuts", "Phím tắt") + " (F1)",
            command=self._show_shortcuts
        )
        settings_menu.add_command(
            label="ℹ️ " + self.texts.get("menu_about", "Thông tin"),
            command=self._show_about
        )
    
    def _toggle_auto_open(self) -> None:
        """Toggle auto-open setting."""
        self.config.auto_open_output = self.auto_open_var.get()
        self.config.save()
    
    def _show_shortcuts(self) -> None:
        """Show keyboard shortcuts."""
        shortcuts = """
⌨️ Keyboard Shortcuts:

Ctrl+O  -  Open file
Ctrl+R  -  Process file
Ctrl+P  -  Preview
Ctrl+S  -  Save file
Ctrl+Q  -  Quit
F1      -  Show shortcuts
        """
        messagebox.showinfo("Shortcuts", shortcuts.strip())
    
    def _show_about(self) -> None:
        """Show about dialog."""
        about_text = f"""
{APP_TITLE}
Version: {APP_VERSION}

Xử lý manifest container cho VINAFCO
Operators: VFC (COC), SVF (SOC)

© 2026 Tiền Tấn - Thuận Port
        """
        messagebox.showinfo("About", about_text.strip())
    
    def _on_close(self) -> None:
        """Handle window close."""
        self.root.destroy()


class PreviewWindow(tk.Toplevel):
    """Preview window for processed data."""
    
    def __init__(self, parent: tk.Tk, df: pd.DataFrame):
        super().__init__(parent)
        self.df = df
        
        self.title("Preview - Processed Data")
        self.geometry("1000x600")
        self.minsize(800, 400)
        
        self._setup_ui()
        self._populate_data()
        
        self.transient(parent)
        self.grab_set()
    
    def _setup_ui(self) -> None:
        """Setup preview window UI."""
        # Header
        header_frame = ttk.Frame(self)
        header_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(
            header_frame,
            text=f"Total rows: {len(self.df)}",
            font=('Segoe UI', 10, 'bold')
        ).pack(side=tk.LEFT)
        
        ttk.Button(
            header_frame,
            text="Close",
            command=self.destroy
        ).pack(side=tk.RIGHT)
        
        # Treeview
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        y_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        x_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
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
        
        for col in columns:
            self.tree.heading(col, text=col, anchor=tk.W)
            self.tree.column(col, width=100, minwidth=50, anchor=tk.W)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
    
    def _populate_data(self) -> None:
        """Populate treeview."""
        display_df = self.df.head(500)
        
        for idx, row in display_df.iterrows():
            values = [str(v) if pd.notna(v) else '' for v in row.values]
            self.tree.insert('', tk.END, values=values)


def main():
    """Main entry point."""
    root = tk.Tk()
    app = ManifestCleanerApp(root)
    root.mainloop()


class Application:
    """Application wrapper for compatibility with vinafco_app.py."""
    
    def __init__(self):
        self.root = tk.Tk()
        self.app = ManifestCleanerApp(self.root)
    
    def run(self):
        """Run the application."""
        self.root.mainloop()


def create_app() -> Application:
    """Create and return application instance."""
    return Application()


if __name__ == "__main__":
    main()
