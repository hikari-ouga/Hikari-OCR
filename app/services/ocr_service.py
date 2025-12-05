from __future__ import annotations

import os
from typing import Optional, Dict, Any

from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential

from ..domain.invoice import Invoice


class OcrService:
    """
    Azure Document Intelligence (Form Recognizer) を使って
    PDF を解析し、Invoice ドメインオブジェクトを返すサービス。
    """

    def __init__(self, cfg: Dict[str, Any]) -> None:
        self.cfg = cfg
        self._client: Optional[DocumentAnalysisClient] = None

    def _get_client(self) -> DocumentAnalysisClient:
        if self._client is not None:
            return self._client

        endpoint = os.environ.get("AZURE_FORMREC_ENDPOINT")
        key = os.environ.get("AZURE_FORMREC_KEY")

        if (not endpoint or not key) and self.cfg is not None:
            azure_cfg = self.cfg.get("azure", {})
            endpoint = endpoint or azure_cfg.get("endpoint")
            key = key or azure_cfg.get("key")

        if not endpoint or not key:
            raise RuntimeError(
                "Azure Form Recognizer の接続情報が不足しています。\n"
                "環境変数 AZURE_FORMREC_ENDPOINT / AZURE_FORMREC_KEY "
                "もしくは config.json の azure.endpoint / azure.key を設定してください。"
            )

        self._client = DocumentAnalysisClient(
            endpoint=endpoint,
            credential=AzureKeyCredential(key),
        )
        return self._client

    def analyze_invoice(self, pdf_bytes: bytes, mode: str = "single") -> Invoice:
        """
        PDF バイト列を解析し、Invoice オブジェクトを返す。

        mode:
          - "single": 1PDF = 1ヶ月分
          - "multi":  1PDF内に複数月が含まれている
        """
        client = self._get_client()
        poller = client.begin_analyze_document("prebuilt-invoice", pdf_bytes)
        result = poller.result()

        full_text: str = result.content or ""
        invoice = Invoice.from_text(full_text, self.cfg, mode=mode)
        return invoice
