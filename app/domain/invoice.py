from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict


# ----------------------
# Invoice 本体
# ----------------------
@dataclass
class Invoice:
    """
    1つの請求を表すオブジェクト。

    fields:
      - "1月値"〜"12月値" をキーに、kWh の数値文字列を入れる。
        単月モード: 通常は1つだけ入る（例: {"1月値": "12345"}）
        複数月モード: 1PDF内に複数月あれば複数キーが入る
    raw_text:
      - Azure OCR から取得したテキスト全体
    """

    fields: Dict[str, str] = field(default_factory=dict)
    raw_text: str = ""
