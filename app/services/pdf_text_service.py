from __future__ import annotations

import logging
from io import BytesIO

from pdfminer.high_level import extract_text

logger = logging.getLogger(__name__)


def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    """Extract embedded text layer from a PDF.

    Returns empty string if extraction fails or no text exists.
    """
    try:
        text = extract_text(BytesIO(pdf_bytes)) or ""
        return text.strip()
    except Exception as e:
        logger.warning(f"PDFテキスト抽出に失敗: {e}")
        return ""
