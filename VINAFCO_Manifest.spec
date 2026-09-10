# -*- mode: python ; coding: utf-8 -*-
# VINAFCO Manifest Cleaner v3.0.0
# PyInstaller spec - Updated 2026-09-10

a = Analysis(
    ['vinafco_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Config files
        ('config.ini', '.'),
        ('app_config.json', '.'),
        ('strings_vi.ini', '.'),
        ('strings_en.ini', '.'),
        # Python packages (source modules)
        ('config', 'config'),
        ('core', 'core'),
        ('gui', 'gui'),
        ('utils', 'utils'),
    ],
    hiddenimports=[
        # numpy (required by pandas)
        'numpy',
        'numpy.core',
        'numpy.core._multiarray_umath',
        'numpy.core._multiarray_tests',
        'numpy.core.multiarray',
        'numpy.core.numeric',
        'numpy.core._dtype_ctypes',
        'numpy.lib.format',
        # pandas + excel engines
        'pandas',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.missing',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.styles.fonts',
        'openpyxl.styles.fills',
        'openpyxl.styles.borders',
        'openpyxl.utils',
        'openpyxl.utils.dataframe',
        'openpyxl.reader.excel',
        'openpyxl.workbook',
        'xlrd',
        'xlsxwriter',
        'xlsxwriter.utility',
        # tkinter
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        # stdlib
        'configparser',
        'logging',
        'logging.handlers',
        'threading',
        'pathlib',
        'json',
        're',
        'winsound',
        'subprocess',
        'os',
        'sys',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        # NOTE: numpy must NOT be excluded - pandas requires it
        'PIL',
        'cv2',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
        'IPython',
        'jupyter',
        'notebook',
        'pytest',
        'doctest',
        'unittest',
    ],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='VINAFCO_Manifest_Cleaner',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # version info
    version=None,
)
