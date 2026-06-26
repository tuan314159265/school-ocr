"""
PyInstaller hook cho Streamlit — thu thập static files và hidden imports.

Streamlit cần các static assets (frontend build) và nhiều hidden import
mà PyInstaller không tự detect được. Hook này đảm bảo chúng được gom vào .exe.
"""

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Thu thập toàn bộ static files của streamlit (frontend build, vendor, ...)
datas = collect_data_files("streamlit")

# Hidden imports thường bị miss
hiddenimports = collect_submodules("streamlit")

# Các module bổ sung mà streamlit cần runtime
extra_hidden = [
    "streamlit.runtime",
    "streamlit.runtime.caching",
    "streamlit.runtime.state",
    "streamlit.runtime.scriptrunner",
    "streamlit.runtime.uploaded_file_manager",
    "streamlit.web",
    "streamlit.web.server",
    "streamlit.web.server.server",
    "streamlit.web.server.websocket_headers",
    "streamlit.web.server.routes",
    "streamlit.elements",
    "streamlit.elements.lib",
    "streamlit.elements.widgets",
    "streamlit.elements.text_widgets",
    "streamlit.proto",
    "altair",
    "altair.vegalite",
    "pyarrow",
    "pyarrow.lib",
    "pandas",
    "pandas._libs",
    "pandas.io",
    "pandas.io.formats",
    "pandas.io.formats.excel",
    "openpyxl",
    "openpyxl.cell",
    "openpyxl.styles",
    "openpyxl.worksheet",
    "openpyxl.reader",
    "openpyxl.writer",
    "PIL",
    "PIL.Image",
    "PIL.ImageEnhance",
    "PIL.ImageOps",
    "google.generativeai",
]

# Merge vào hiddenimports (tránh trùng)
for mod in extra_hidden:
    if mod not in hiddenimports:
        hiddenimports.append(mod)
