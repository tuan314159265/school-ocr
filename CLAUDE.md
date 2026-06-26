# CLAUDE.md — Project Context: School Registration OCR to Excel

## Mục tiêu dự án

Xây dựng công cụ desktop (`.exe`) giúp nhân viên kế toán trường tiểu học số hóa **Giấy Khai Sinh** và **Trích lục Khai Sinh** (bản giấy) thành danh sách học sinh nhập học dạng file Excel. Ứng dụng chạy local trên Windows, không cần cài thêm phần mềm.

---

## Tech Stack

| Layer | Công nghệ |
|---|---|
| UI | Streamlit |
| Backend | Python 3.10+ |
| OCR/AI | Google Gemini API (`gemini-1.5-pro`) |
| Data | `pandas`, `openpyxl` |
| Image pre-processing | `Pillow` |
| Config | `config.json` (lưu API key local) |
| Packaging | `PyInstaller` (single `.exe`) |
| Entry point | `run.py` (tìm port, spawn Streamlit, mở browser) |

---

## Cấu trúc thư mục mục tiêu

```
project/
├── app.py              # Streamlit UI chính
├── ocr_engine.py       # Logic gọi Gemini API, trả JSON
├── run.py              # Entry point: tìm port, chạy app, mở browser
├── config.json         # Lưu API key (tạo lần đầu, KHÔNG commit)
├── requirements.txt
├── school_ocr.spec     # PyInstaller spec file
└── hooks/
    └── hook-streamlit.py  # PyInstaller hook cho Streamlit
```

---

## Schema cố định — Giấy Khai Sinh / Trích lục Khai Sinh

Nguồn dữ liệu đầu vào là **Giấy Khai Sinh** hoặc **Trích lục Khai Sinh** (in sẵn, có thể có chữ điền tay). Mục tiêu output là danh sách học sinh nhập học (Phase A1 — chưa có SĐT phụ huynh).

Gemini **phải** trả về JSON với đúng 8 key sau (không được tự ý đặt tên khác):

```json
{
  "ho_ten_hoc_sinh": "",
  "ngay_sinh": "",
  "gioi_tinh": "",
  "dan_toc": "",
  "noi_sinh": "",
  "ho_ten_cha": "",
  "ho_ten_me": "",
  "noi_cu_tru": ""
}
```

**Mapping từ Giấy Khai Sinh sang schema:**
| Key | Lấy từ trường nào trên giấy |
|---|---|
| `ho_ten_hoc_sinh` | "Họ, chữ đệm, tên" (của trẻ) |
| `ngay_sinh` | "Ngày, tháng, năm sinh" — format `DD/MM/YYYY` |
| `gioi_tinh` | "Giới tính" |
| `dan_toc` | "Dân tộc" (của trẻ) |
| `noi_sinh` | "Nơi sinh" |
| `so_dinh_danh` | "Số định danh cá nhân" |
| `ho_ten_cha` | "Họ, chữ đệm, tên người cha" |
| `ho_ten_me` | "Họ, chữ đệm, tên người mẹ" |
| `noi_cu_tru` | Nơi cư trú của cha hoặc mẹ (ưu tiên cha, nếu không có thì lấy mẹ) |

Nếu một trường không đọc được hoặc không có trên giấy, để giá trị là chuỗi rỗng `""`. Tuyệt đối không bịa thêm key.

---

## Các file cần viết — theo thứ tự

### Step 0 (tiền đề — không cần file mới)
- Nguồn ảnh đầu vào: **Giấy Khai Sinh** hoặc **Trích lục Khai Sinh** (Phase A1).
- Schema JSON 8 trường ở trên là cố định cho toàn bộ project phase này.
- Mọi xử lý DataFrame phải dựa trên đúng 8 key này.

### Step 1 — `requirements.txt` + `ocr_engine.py`

**`requirements.txt`** liệt kê đủ dependencies, pin version hợp lý:
```
streamlit
google-generativeai
pandas
openpyxl
Pillow
```

**`ocr_engine.py`** phải có:
- Class `OCREngine` nhận `api_key: str` lúc khởi tạo.
- Method `process_image(image_path_or_pil: str | PIL.Image) -> dict` :
  - Pre-process ảnh bằng Pillow: auto-rotate (EXIF), tăng contrast nhẹ, resize nếu > 4MB.
  - Gọi Gemini API với prompt tiếng Việt, yêu cầu trả về JSON strict theo schema ở trên.
  - Parse JSON response, validate đủ key, fill key thiếu bằng `""`.
  - Trả về `dict` theo schema + thêm key `"_confidence"` (`"ok"` / `"low"` / `"error"`) và `"_filename"`.
- Xử lý exception: network error, quota exceeded, invalid JSON response, ảnh không đọc được.
- Có comment tiếng Việt rõ ràng.

**Prompt mẫu gửi Gemini** (nhúng trong `ocr_engine.py`):
```
Bạn là công cụ OCR chuyên nghiệp. Hãy đọc thông tin từ Giấy Khai Sinh hoặc Trích lục 
Khai Sinh trong ảnh (tiếng Việt, có thể có chữ điền tay). Trả về DUY NHẤT một JSON object 
với các key sau, không thêm bất kỳ text hay markdown nào khác:
{schema_json}
Lưu ý:
- "noi_cu_tru" lấy từ nơi cư trú của cha, nếu không có thì lấy của mẹ.
- "ngay_sinh" định dạng DD/MM/YYYY.
- Nếu không đọc được một trường, để giá trị là "". Không được bịa thông tin.
```

### Step 2 — `app.py`

Giao diện Streamlit gồm các phần:

**A. Settings (sidebar):**
- Input nhập Gemini API key, có nút Save → lưu vào `config.json`.
- Dropdown chọn model (`gemini-1.5-pro` mặc định, `gemini-1.5-flash` tùy chọn).
- Hiển thị trạng thái API key (đã lưu / chưa có).

**B. Upload & Scan:**
- `st.file_uploader` cho phép chọn nhiều file `.jpg`, `.jpeg`, `.png`.
- Nút **"Bắt đầu quét dữ liệu"** → xử lý tuần tự từng ảnh.
- Progress bar + spinner hiển thị tên file đang xử lý.
- Sau mỗi ảnh append kết quả vào `st.session_state["records"]`.

**C. Data Editor:**
- Dùng `st.data_editor` hiển thị DataFrame từ `st.session_state["records"]`.
- Cột `_confidence` hiển thị emoji: ✅ (ok), ⚠️ (low), ❌ (error) — read-only.
- Cột `_filename` read-only, các cột còn lại editable.
- Cho phép xóa dòng.

**D. Export:**
- Nút **"Tải xuống file Excel"** → dùng `openpyxl` xuất file, bỏ cột `_confidence` và `_filename` khỏi output, format header tiếng Việt đẹp.
- Tên file mặc định: `danh_sach_hoc_sinh_YYYYMMDD.xlsx`.
- Header tiếng Việt đẹp, mapping từ key sang tên cột:

| Key | Tên cột Excel |
|---|---|
| `ho_ten_hoc_sinh` | Họ và tên học sinh |
| `ngay_sinh` | Ngày sinh |
| `gioi_tinh` | Giới tính |
| `dan_toc` | Dân tộc |
| `que_quan` | Quê quán |
| `noi_sinh` | Nơi sinh |
| `so_dinh_danh` | Số định danh |
| `ho_ten_cha` | Họ tên cha |
| `ho_ten_me` | Họ tên mẹ |
| `noi_cu_tru` | Nơi cư trú |

**Lưu ý:**
- Toàn bộ state (records, uploaded files) giữ qua `st.session_state`.
- Nếu chưa có API key thì disable nút scan và hiển thị cảnh báo.

### Step 3 — `run.py`

Entry point khi chạy `.exe`:
1. Tìm port trống trong dải 8500–9000 bằng `socket`.
2. Spawn subprocess: `streamlit run app.py --server.port <port> --server.headless true`.
3. Handle `sys._MEIPASS` để tìm đúng đường dẫn `app.py` khi chạy từ `.exe`.
4. Đợi tối đa 10 giây (poll thử kết nối) rồi `webbrowser.open(...)`.
5. Giữ process cha sống cho đến khi user đóng cửa sổ terminal / tray.

### Step 4 — `school_ocr.spec` + `hooks/hook-streamlit.py`

**`hooks/hook-streamlit.py`** — PyInstaller hook xử lý Streamlit static files:
- Thu thập toàn bộ `streamlit` package data (static assets, vendor).
- Xử lý các hidden imports thường bị miss: `streamlit.runtime`, `altair`, `pyarrow`.

**`school_ocr.spec`:**
- `onefile=True`, `windowed=False` (giữ console để debug lần đầu).
- Khai báo `datas` bao gồm Streamlit static files, `app.py`, `ocr_engine.py`.
- Khai báo `hiddenimports` đầy đủ.
- `icon` để trống hoặc dùng icon tùy chọn.

Cung cấp lệnh build cuối cùng:
```bash
pyinstaller school_ocr.spec
```

---

## Constraints & Rules cho Claude Code

1. **Không hardcode API key** bất cứ đâu. Key chỉ tồn tại trong `config.json` (gitignored).
2. **Schema JSON là bất biến** — không tự thêm/đổi tên key ở bất kỳ file nào.
3. **Comment bằng tiếng Việt** cho các đoạn logic quan trọng.
4. **Xử lý exception** cho mọi lời gọi API và IO.
5. **Hỏi trước khi chuyển step** — sau mỗi step, dừng lại để confirm.
6. Ưu tiên **đơn giản, dễ bảo trì** hơn clever code.
7. Mỗi file phải chạy được độc lập để test (không circular import).

---

## Thứ tự thực thi

```
Step 1 → confirm → Step 2 → confirm → Step 3 → confirm → Step 4
```

Bắt đầu bằng Step 1: viết `requirements.txt` và `ocr_engine.py`.
