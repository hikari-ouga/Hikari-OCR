from __future__ import annotations

import os
from typing import Dict, Any, List, Optional

from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential

from ..domain.invoice import Invoice


class OcrService:
    """
    Azure Document Intelligence を使って PDF を解析し、

    - 単月モード: 1PDF = 1ヶ月分
    - 複数月モード: 12 / 24 / 36 ページ → 開始月から割り当てて12ヶ月分抽出

    の Invoice を生成するサービス。
    """

    def __init__(self, cfg: Dict[str, Any]) -> None:
        self.cfg = cfg

        # 環境変数固定
        endpoint = os.getenv("AZURE_FORMREC_ENDPOINT")
        key = os.getenv("AZURE_FORMREC_KEY")

        if not endpoint or not key:
            raise ValueError(
                "Form Recognizer の endpoint / key が見つかりません。\n"
                "環境変数 AZURE_FORMREC_ENDPOINT と AZURE_FORMREC_KEY を確認してください。"
            )

        self.client = DocumentAnalysisClient(
            endpoint=endpoint,
            credential=AzureKeyCredential(key),
        )

        # prebuilt-invoice をデフォルトで使用
        self.model_id: str = cfg.get("FORM_RECOGNIZER_MODEL_ID", "prebuilt-invoice")

    # --------------------------------------------------------
    # 公開メソッド：単月 / 複数月モードの切り替え
    # --------------------------------------------------------
    def analyze_invoice(
        self,
        content: bytes,
        mode: str = "single",
        start_month: Optional[int] = None,
    ) -> Invoice:

        if mode == "multi":
            if start_month is None:
                raise ValueError("複数月モードでは start_month が必須です。")
            return self._analyze_multi(content, start_month)

        # デフォルトは単月モード
        return self._analyze_single(content)

    # --------------------------------------------------------
    # 単月モード：1PDF = 1ヶ月分
    # --------------------------------------------------------
    def _analyze_single(self, content: bytes) -> Invoice:
        poller = self.client.begin_analyze_document(
            model_id=self.model_id,
            document=content,
        )
        result = poller.result()

        # ✅ SDK v3 系: 全文テキストは result.content に入っている
        full_text = result.content or ""

        invoice = Invoice.from_text(full_text, self.cfg, mode="single")
        return invoice

    # --------------------------------------------------------
    # 複数月モード：12 / 24 / 36ページを開始月から割り当てて12ヶ月分生成
    # --------------------------------------------------------
    def _analyze_multi(self, content: bytes, start_month: int) -> Invoice:

        poller = self.client.begin_analyze_document(
            model_id=self.model_id,
            document=content,
        )
        result = poller.result()

        # ✅ ページごとのテキストは result.content から spans で切り出す
        page_texts: List[str] = []
        full_content = result.content or ""

        for page in result.pages:
            if page.spans:
                # 通常1ページ1spanなので先頭だけ取ればOK
                span = page.spans[0]
                start = span.offset
                end = span.offset + span.length
                page_texts.append(full_content[start:end])
            else:
                page_texts.append("")

        num_pages = len(page_texts)

        if num_pages not in (12, 24, 36):
            raise ValueError(
                f"複数月モードは 12 / 24 / 36 ページのみ対応しています（実際: {num_pages}ページ）"
            )

        pages_per_month = num_pages // 12  # 12→1、24→2、36→3
        fields: Dict[str, str] = {}

        current_month = start_month

        for i in range(12):
            start_idx = i * pages_per_month
            end_idx = start_idx + pages_per_month
            month_text = "\n".join(page_texts[start_idx:end_idx])

            # kWh 抽出（単月と同じロジック）
            kwh_value = self._extract_kwh_from_text(month_text)
            if kwh_value:
                mapped_month = 12 if current_month == 1 else current_month - 1
                fields[f"{mapped_month}月値"] = kwh_value

            current_month = self._next_month(current_month)

        full_text = "\n".join(page_texts)
        return Invoice(fields=fields, raw_text=full_text)

    # --------------------------------------------------------
    # kWh 抽出（単月と同じ）
    # --------------------------------------------------------
    @staticmethod
    def _extract_kwh_from_text(text: str) -> str:
        import re

        matches = re.findall(r"([\d,]+)\s*kWh", text, flags=re.IGNORECASE)
        nums = []

        for raw in matches:
            raw = raw.replace(",", "")
            try:
                v = int(raw)
                if v > 0:
                    nums.append(v)
            except:
                continue

        if not nums:
            return ""
        return str(max(nums))

    @staticmethod
    def _next_month(month: int) -> int:
        return 1 if month == 12 else month + 1

