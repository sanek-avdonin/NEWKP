from __future__ import annotations

from pathlib import Path
from typing import Optional
import io

import fitz  # PyMuPDF
import pytesseract
from PIL import Image

from kp_generator.config import TESSERACT_CANDIDATES  # абсолютный импорт


def configure_tesseract() -> Optional[Path]:
    for p in TESSERACT_CANDIDATES:
        if p.exists():
            pytesseract.pytesseract.tesseract_cmd = str(p)
            return p
    return None


def extract_text_from_pdf(pdf_path: str, ocr_lang: str = "rus") -> str:
    # закрытие PDF через контекстный менеджер
    with fitz.open(pdf_path) as doc:
        # 1) Пробуем цифровой текст
        text_parts = []
        for page in doc:
            text_parts.append(page.get_text("text"))
        text = "\n".join(text_parts).strip()

        # эвристика: если текста достаточно — OCR не нужен
        if len(text) >= 500:
            return text

        # 2) OCR fallback
        configured = configure_tesseract()
        if configured is None:
            raise RuntimeError(
                "PDF выглядит как скан/фото, нужен OCR, но tesseract.exe не найден. "
                "Добавьте tesseract рядом с программой или в папку tesseract/."
            )

        ocr_parts = []
        for page in doc:
            pix = page.get_pixmap(dpi=300)
            img_bytes = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_bytes))
            ocr_parts.append(pytesseract.image_to_string(img, lang=ocr_lang))

        return "\n".join(ocr_parts)
