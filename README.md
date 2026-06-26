# School Registration OCR

Công cụ OCR tự động đọc **Giấy Khai Sinh / Trích lục Khai Sinh** thành file Excel danh sách học sinh.

## Cách dùng (Streamlit Cloud)

**Bước 1 — Deploy:**

1. Fork repo này về GitHub của bạn (hoặc push lên repo private)
2. Vào https://share.streamlit.io , đăng nhập bằng GitHub
3. Chọn repo `school-ocr`, bấm **Deploy**
4. Sau khi deploy xong, vào **Settings → Secrets** và thêm:

```
api_key = "Gemini_API_key_của_bạn"
model = "gemini-2.0-flash"
```

5. Bấm **Save**, app tự chạy lại — xong!

**Bước 2 — Sử dụng:**

1. Mở link web app (dạng `https://school-ocr-xxx.streamlit.app`)
2. **Mở file Excel cũ** (nếu có) để nạp danh sách học sinh hiện tại
3. **Kéo thả ảnh** Giấy Khai Sinh vào khung upload
4. Bấm **Bắt đầu quét dữ liệu** — app tự OCR và thêm vào danh sách
5. **Lưu file Excel** — tải về danh sách tổng hợp

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
