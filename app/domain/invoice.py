from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, Any, Optional, List, Tuple


# ----------------------
# ヘルパー関数
# ----------------------
def _safe_int(num_str: str) -> Optional[int]:
    try:
        return int(num_str)
    except ValueError:
        return None


def extract_month_for_single(text: str) -> Optional[int]:
    """
    【単月モード】
    テキスト全体から「このPDFが何月分か」を推定する。

    優先順位:
      1. 「○月ご使用分」「○月分」「○月ご請求分」などのパターン
      2. それが無ければ、単純に最初に出てくる「○月」
    """
    patterns = [
        r"(\d{1,2})月[^\n]{0,5}ご使用分",
        r"(\d{1,2})月[^\n]{0,5}ご請求分",
        r"(\d{1,2})月分",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            month = _safe_int(m.group(1))
            if month and 1 <= month <= 12:
                return month

    m = re.search(r"(\d{1,2})月", text)
    if not m:
        return None

    month = _safe_int(m.group(1))
    if month and 1 <= month <= 12:
        return month

    return None


def extract_kwh_single(text: str) -> str:
    """
    【単月モード】
    テキスト全体から「kWh が付いている 0 以外の数値」を探す。
    複数あった場合は一番大きい値を返す。
    """
    matches = re.findall(r"([\d,]+)\s*kWh", text, flags=re.IGNORECASE)
    values: List[int] = []

    for raw in matches:
        num_str = raw.replace(",", "")
        v = _safe_int(num_str)
        if v is None or v == 0:
            continue
        values.append(v)

    if not values:
        return ""

    return str(max(values))


def extract_month_segments(text: str) -> List[Tuple[int, int, int]]:
    """
    【複数月モード】
    テキスト中の「○月」の位置を全て探し、
    (month, start_index, end_index_of_segment) のリストを返す。

    セグメント:
      - i番目の「○月」の位置から、(i+1)番目の「○月」の直前まで
      - 最後の月はテキストの末尾まで
    """
    # すべての「○月」の出現位置を取得
    matches = list(re.finditer(r"(\d{1,2})月", text))
    segments: List[Tuple[int, int, int]] = []

    for idx, m in enumerate(matches):
        month = _safe_int(m.group(1))
        if month is None or not (1 <= month <= 12):
            continue
        start = m.start()  # この「○月」の位置

        if idx + 1 < len(matches):
            end = matches[idx + 1].start()  # 次の「○月」の直前まで
        else:
            end = len(text)  # 最後は末尾まで

        segments.append((month, start, end))

    return segments


def extract_kwh_from_segment(segment: str) -> Optional[int]:
    """
    セグメント内から「kWh がつく 0 以外の数値」を探し、
    一番大きい値を返す（なければ None）。
    """
    matches = re.findall(r"([\d,]+)\s*kWh", segment, flags=re.IGNORECASE)
    values: List[int] = []

    for raw in matches:
        num_str = raw.replace(",", "")
        v = _safe_int(num_str)
        if v is None or v == 0:
            continue
        values.append(v)

    if not values:
        return None

    return max(values)


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

    @classmethod
    def from_text(cls, text: str, cfg: Dict[str, Any], mode: str = "single") -> "Invoice":
        """
        OCR済みテキストをもとに Invoice を生成する。

        mode:
          - "single": 1PDF = 1ヶ月分（これまで通り）
          - "multi":  1PDF内に複数月が含まれている前提で、
                      各「○月」ごとに kWh を探して埋める
        """
        if mode == "multi":
            fields = cls._from_text_multi(text)
        else:
            fields = cls._from_text_single(text)

        return cls(fields=fields, raw_text=text)

    # ----------------------
    # 単月モード
    # ----------------------
    @staticmethod
    def _from_text_single(text: str) -> Dict[str, str]:
        fields: Dict[str, str] = {}

        month = extract_month_for_single(text)
        kwh_value = extract_kwh_single(text)

        if month is not None and kwh_value:
            # 1月→12月、2月→1月、… という形で -1 シフト
            mapped_month = 12 if month == 1 else month - 1
            key = f"{mapped_month}月値"
            fields[key] = kwh_value

        return fields

    # ----------------------
    # 複数月モード
    # ----------------------
    @staticmethod
    def _from_text_multi(text: str) -> Dict[str, str]:
        """
        テキスト中に 10月, 11月, 12月... と複数月が出てくる前提で、

        10月 ... (この区間内) ... 1234kWh
        11月 ... (この区間内) ... 2345kWh
        ...

        のように、
        「○月 〜 次の○月の直前まで」を1セグメントとして、その中の kWh を拾う。
        """
        fields: Dict[str, str] = {}

        segments = extract_month_segments(text)
        if not segments:
            return fields

        for month, start, end in segments:
            segment_text = text[start:end]
            v = extract_kwh_from_segment(segment_text)
            if v is None:
                continue

            mapped_month = 12 if month == 1 else month - 1
            key = f"{mapped_month}月値"
            # 同じ月が複数回出てきた場合は「後勝ち」とする
            fields[key] = str(v)

        return fields
