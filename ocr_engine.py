"""
Module OCR Engine — Gọi Gemini API để đọc Giấy Khai Sinh / Trích lục Khai Sinh.

Class chính:
    OCREngine        — Nhận API key, xử lý ảnh, trả JSON theo schema cố định.

Schema output (10 keys):
    ho_ten_hoc_sinh, ngay_sinh, gioi_tinh, dan_toc, que_quan,
    noi_sinh, so_dinh_danh, ho_ten_cha, ho_ten_me, noi_cu_tru
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any

from google import genai
from PIL import Image, ImageEnhance, ImageOps

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Schema JSON cố định — 8 key, không thêm bớt.
SCHEMA_KEYS: list[str] = [
    "ho_ten_hoc_sinh",
    "ngay_sinh",
    "gioi_tinh",
    "dan_toc",
    "noi_sinh",
    "ho_ten_cha",
    "ho_ten_me",
    "noi_cu_tru",
]

# Prompt gửi cho Gemini — tiếng Việt, yêu cầu JSON thuần.
PROMPT_TEMPLATE: str = """\
Bạn là công cụ OCR chuyên nghiệp. Hãy đọc thông tin từ Giấy Khai Sinh hoặc Trích lục
Khai Sinh trong ảnh (tiếng Việt, có thể có chữ điền tay). Trả về DUY NHẤT một JSON object
với các key sau, không thêm bất kỳ text hay markdown nào khác:
{schema_json}
Lưu ý:
- "noi_cu_tru" lấy từ nơi cư trú của cha, nếu không có thì lấy của mẹ.
- "ngay_sinh" định dạng DD/MM/YYYY.
"""


def _preprocess_image(image: Image.Image) -> Image.Image:
    """
    Tiền xử lý ảnh trước khi gửi Gemini:
      - Tự động xoay theo EXIF.
      - Tăng độ tương phản nhẹ.
      - Resize nếu ảnh quá lớn (> 4MB ước lượng).
    """
    # Auto-rotate theo EXIF
    image = ImageOps.exif_transpose(image) or image

    # Chuyển về RGB nếu là ảnh trong suốt (RGBA) hoặc ảnh đơn sắc
    if image.mode in ("RGBA", "P"):
        image = image.convert("RGB")

    # Tăng contrast nhẹ giúp chữ dễ đọc hơn
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(1.3)

    # Resize nếu ảnh quá lớn — Gemini giới hạn dung lượng
    # Ước lượng: ảnh RGB ở 72DPI, ~3 byte/pixel → 4MB ≈ 1.4M pixels
    max_pixels = 4_000_000
    width, height = image.size
    if width * height > max_pixels:
        ratio = (max_pixels / (width * height)) ** 0.5
        new_width = int(width * ratio)
        new_height = int(height * ratio)
        image = image.resize((new_width, new_height), Image.LANCZOS)
        logger.info("Đã resize ảnh từ %dx%d xuống %dx%d", width, height, new_width, new_height)

    return image


def _validate_and_fill(data: dict[str, Any]) -> dict[str, Any]:
    """
    Kiểm tra dict trả về từ Gemini có đủ 10 key không.
    Key nào thiếu hoặc không phải string thì fill bằng "".
    """
    result: dict[str, Any] = {}
    for key in SCHEMA_KEYS:
        value = data.get(key)
        if isinstance(value, str):
            result[key] = value.strip()
        else:
            result[key] = ""
    return result


def _calculate_confidence(data: dict[str, Any]) -> str:
    """
    Tính độ tin cậy dựa trên số trường trống:
      - "ok":   không có trường nào rỗng.
      - "low":  1-3 trường rỗng.
      - "error": >=4 trường rỗng hoặc toàn bộ rỗng.
    """
    empty_count = sum(1 for k in SCHEMA_KEYS if not data.get(k))
    if empty_count == 0:
        return "ok"
    elif empty_count <= 3:
        return "low"
    else:
        return "error"


class OCREngine:
    """
    Engine OCR sử dụng Gemini API.

    Cách dùng:
        engine = OCREngine(api_key="...")
        result = engine.process_image("path/to/image.jpg")
    """

    def __init__(self, api_key: str) -> None:
        """
        Khởi tạo engine với API key.
        Model mặc định: gemini-2.0-flash (quota 1500 req/ngày, free).
        """
        self.api_key = api_key
        self._client: genai.Client | None = None
        self.model_name: str = "gemini-3.1-flash-lite"
        logger.info("OCREngine khởi tạo với model %s", self.model_name)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def process_image(self, image_source: str | Path | Image.Image) -> dict[str, Any]:
        """
        Xử lý một ảnh (đường dẫn hoặc PIL.Image) và trả về dict theo schema.

        Returns:
            dict với 10 key schema + "_confidence" + "_filename".
        """
        try:
            # Load ảnh nếu là đường dẫn
            if isinstance(image_source, (str, Path)):
                image = Image.open(image_source)
                filename = Path(image_source).name
            else:
                image = image_source
                filename = getattr(image_source, "filename", "unknown")

            # Tiền xử lý
            processed = _preprocess_image(image)

            # Gửi lên Gemini
            raw_response = self._call_gemini(processed)

            # Parse JSON
            parsed = self._parse_response(raw_response)

            # Validate & fill thiếu
            cleaned = _validate_and_fill(parsed)

            # Tính confidence
            confidence = _calculate_confidence(cleaned)
        except Exception as exc:
            logger.error("Lỗi xử lý ảnh %s: %s", image_source, exc)
            cleaned = {k: "" for k in SCHEMA_KEYS}
            confidence = "error"

        # Gắn metadata
        cleaned["_confidence"] = confidence
        cleaned["_filename"] = filename
        return cleaned

    def update_model(self, model_name: str) -> None:
        """Cập nhật model Gemini (vd: gemini-2.0-flash, gemini-1.5-pro)."""
        self.model_name = model_name
        logger.info("Đã chuyển model sang %s", model_name)

    def test_connection(self) -> bool:
        """
        Kiểm tra API key có hoạt động không bằng cách list models.
        Không tốn quota generate_content.
        """
        try:
            client = self._get_client()
            # Chỉ lấy 1 model, không generate — không tốn quota
            for _ in client.models.list():
                return True
            return False
        except Exception as exc:
            logger.warning("Test connection thất bại: %s", exc)
            return False

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _get_client(self) -> genai.Client:
        """Lấy hoặc khởi tạo Gemini client."""
        if self._client is None:
            self._client = genai.Client(api_key=self.api_key)
        return self._client

    def _call_gemini(self, image: Image.Image, max_retries: int = 3) -> str:
        """
        Gửi ảnh + prompt lên Gemini, trả về text response.
        Tự động retry khi bị rate limit (429) hoặc lỗi tạm thời.
        """
        import time as _time

        client = self._get_client()

        schema_str = json.dumps({k: "" for k in SCHEMA_KEYS}, ensure_ascii=False, indent=2)
        prompt = PROMPT_TEMPLATE.format(schema_json=schema_str)

        last_error: Exception | None = None

        for attempt in range(max_retries):
            try:
                response = client.models.generate_content(
                    model=self.model_name,
                    contents=[prompt, image],
                )

                if not response.text:
                    raise ValueError("Gemini trả về response rỗng (text = None)")

                return response.text.strip()

            except Exception as exc:
                last_error = exc
                error_str = str(exc)

                # Kiểm tra nếu hết quota ngày — không retry
                if "limit: 0" in error_str:
                    logger.error("API key đã hết quota ngày. Không thể xử lý thêm.")
                    break

                # Retry nếu bị rate limit tạm thời (429) hoặc server quá tải (503)
                if "429" in error_str or "RESOURCE_EXHAUSTED" in error_str or "503" in error_str or "UNAVAILABLE" in error_str:
                    if attempt < max_retries - 1:
                        delay = (10 * (attempt + 1)) + 5  # 15s, 25s, 35s
                        logger.warning(
                            "Gemini quá tải (lần %d/%d), đợi %ds rồi thử lại...",
                            attempt + 1, max_retries, delay,
                        )
                        _time.sleep(delay)
                        continue
                else:
                    # Lỗi khác (4xx, 5xx) — không retry
                    break

        # Hết retry hoặc lỗi không retry được
        raise last_error  # type: ignore[misc]

    def _parse_response(self, text: str) -> dict[str, Any]:
        """
        Parse text response từ Gemini thành dict.
        Xử lý các trường hợp Gemini wrap JSON trong ```json ... ```.
        """
        # Nếu Gemini trả về markdown code block
        if text.startswith("```"):
            start = text.find("{")
            end = text.rfind("}")
            if start != -1 and end != -1:
                text = text[start : end + 1]

        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ValueError(
                f"Không parse được JSON từ response Gemini.\n"
                f"Response: {text[:500]}\n"
                f"Lỗi: {exc}"
            ) from exc

        if not isinstance(data, dict):
            raise ValueError(
                f"Response Gemini không phải object. "
                f"Kiểu nhận được: {type(data).__name__}"
            )

        return data
