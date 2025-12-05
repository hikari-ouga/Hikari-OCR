from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, Any, Optional


def extract_month(text: str) -> Optional[int]:
    """
    テキスト全体から「このPDFが何月分か」を推定する。

    優先順位:
      1. 「○月ご使用分」「○月分」「○月ご請求分」などのパターン
      2. それが無ければ、単純に最初に出てくる「○月」
    """
    # ① 「○月ご使用分」「○月ご請求分」「○月分」系を優先して探す
    patterns = [
        r"(\d{1,2})月[^\n]{0,5}ご使用分",
        r"(\d{1,2})月[^\n]{0,5}ご請求分",
        r"(\d{1,2})月分",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            try:
                month = int(m.group(1))
                if 1 <= month <= 12:
                    return month
            except ValueError:
                pass

    # ② それでも見つからなければ、単純な「○月」を見る
    m = re.search(r"(\d{1,2})月", text)
    if not m:
        return None

    try:
        month = int(m.group(1))
        if 1 <= month <= 12:
            return month
    except ValueError:
        return None

    return None


def extract_kwh_value(text: str) -> str:
    """
    テキスト全体から「kWh が付いている 0 以外の数値」を探す。

    - 大文字/小文字を区別しない (kWh, KWH など)
    - カンマ付きもOK
    - 複数あった場合は「一番大きい値」を採用
    """
    # 例: "12,345 kWh", "12345kwh", "3,210KWH" など
    matches = re.findall(r"([\d,]+)\s*kWh", text, flags=re.IGNORECASE)
    values: list[int] = []

    for raw in matches:
        num_str = raw.replace(",", "")
        try:
            v = int(num_str)
        except ValueError:
            continue
        if v == 0:
            continue
        values.append(v)

    if not values:
        return ""

    # 一番大きい値を採用（合計使用量などを想定）
    return str(max(values))


@dataclass
class Invoice:
    """
    1枚の請求書を表すドメインオブジェクト。

    fields:
      - "1月値"〜"12月値":
          「このPDFは何月分か」＋「このPDF内の kWh の値」を
          月-1シフトした位置に 1 つだけ入れる。
          例）2月分の請求書で 12345kWh があれば:
              fields = {"1月値": "12345"}
    raw_text:
      - Azure OCR から取得したテキスト全体
    """

    fields: Dict[str, str] = field(default_factory=dict)
    raw_text: str = ""

    @classmethod
    def from_text(cls, text: str, cfg: Dict[str, Any]) -> "Invoice":
        fields: Dict[str, str] = {}

        # このPDFが「何月分の請求書か」
        month = extract_month(text)
        # このPDF内の「0じゃない kWh 付き数値」を一つ（最大値）
        kwh_value = extract_kwh_value(text)

        if month is not None and kwh_value:
            # 1月→12月、2月→1月、… という形で -1 シフト
            mapped_month = 12 if month == 1 else month - 1
            key = f"{mapped_month}月値"
            fields[key] = kwh_value

        return cls(fields=fields, raw_text=text)
