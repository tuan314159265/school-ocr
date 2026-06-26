# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec — School Registration OCR to Excel.

Build command:
    pyinstaller school_ocr.spec

Output: dist/school_ocr.exe (single file, console giữ lại để debug).
"""

import sys
from pathlib import Path

# Thư mục gốc của project — spec file luôn ở thư mục gốc
base_dir = Path.cwd()

a = Analysis(
    # Entry point — run.py là file chính
    [str(base_dir / "run.py")],
    pathex=[],
    binaries=[],
    datas=[
        # Package chính
        (str(base_dir / "app.py"), "."),
        (str(base_dir / "ocr_engine.py"), "."),
        # Hook directory (PyInstaller tự động đọc hook-*.py)
        (str(base_dir / "hooks"), "hooks"),
    ],
    hiddenimports=[
        # Streamlit core
        "streamlit",
        "streamlit.runtime",
        "streamlit.runtime.caching",
        "streamlit.runtime.state",
        "streamlit.runtime.scriptrunner",
        "streamlit.runtime.uploaded_file_manager",
        "streamlit.web",
        "streamlit.web.server",
        "streamlit.web.server.server",
        "streamlit.web.server.websocket_headers",
        "streamlit.web.server.server_util",
        "streamlit.elements",
        "streamlit.elements.lib",
        "streamlit.elements.widgets",
        "streamlit.elements.text_widgets",
        "streamlit.proto",
        "streamlit.proto.Common",
        # Altair / Vega
        "altair",
        "altair.vegalite",
        "altair.vegalite.v5",
        "altair.vegalite.v5.schema",
        "altair.vegalite.v5.schema.channels",
        "altair.vegalite.v5.compiler",
        # PyArrow (Altair dependency)
        "pyarrow",
        "pyarrow.lib",
        # Pandas
        "pandas",
        "pandas._libs",
        "pandas._libs.tslibs",
        "pandas._libs.parsers",
        "pandas.io",
        "pandas.io.formats",
        "pandas.io.formats.excel",
        "pandas.io.formats.style",
        # OpenPyXL
        "openpyxl",
        "openpyxl.cell",
        "openpyxl.styles",
        "openpyxl.worksheet",
        "openpyxl.reader",
        "openpyxl.writer",
        # Pillow
        "PIL",
        "PIL.Image",
        "PIL.ImageEnhance",
        "PIL.ImageOps",
        # Google Generative AI (google-genai)
        "google",
        "google.genai",
        "google.genai.models",
        # Socket / urllib (cho run.py)
        "socket",
        "urllib",
        "urllib.request",
        "urllib.error",
    ],
    hookspath=[str(base_dir / "hooks")],
    hooksconfig={},
    excludes=[
        # Loại bỏ các module không cần thiết (giảm dung lượng file)
        "matplotlib",
        "scipy",
        "sympy",
        "notebook",
        "jupyter",
        "jupyter_client",
        "jupyter_core",
        "ipython",
        "tensorflow",
        "torch",
        "tkinter",
        "PyQt5",
        "PySide2",
        "PySide6",
        "sphinx",
        "setuptools._distutils",
    ],
    cipher=None,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

# Tạo single-file executable, giữ console để debug lần đầu
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="school_ocr",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,          # Giữ console để user thấy log
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
