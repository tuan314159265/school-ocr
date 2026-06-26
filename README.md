# School Registration OCR

Công cụ OCR tự động đọc **Giấy Khai Sinh / Trích lục Khai Sinh** thành file Excel danh sách học sinh.

## Chạy local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Tạo file `config.json` cùng thư mục:

```json
{
  "api_key": "AIzaSy_...",
  "model": "gemini-2.0-flash"
}
```

## Build .exe (Windows)

### Cách 1 — GitHub Actions (tự động)

Push tag `v1.0`, `v2.0`,... lên GitHub:

```bash
git tag v1.0
git push origin v1.0
```

Vào GitHub repo → Actions → tải file `.exe` từ artifact.

### Cách 2 — Build thủ công

Trên máy Windows:

```bash
pip install -r requirements.txt pyinstaller
pyinstaller school_ocr.spec
```

File `.exe` ở `dist/school_ocr.exe`, đặt `config.json` cạnh nó.

## Deploy lên Streamlit Cloud

1. Push code lên GitHub
2. Vào https://share.streamlit.io , deploy repo
3. Vào Settings → Secrets, thêm:
   - `api_key`: Gemini API key
   - `model`: `gemini-2.0-flash`
