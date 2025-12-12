from __future__ import annotations

import re
from typing import Dict, Any, List, Optional
from pathlib import Path

import streamlit as st

from ..domain.invoice import Invoice
from ..services.ocr_service import OcrService
from ..services.excel_service import ExcelService


# ------------------------------------------------------------
# ファイル名から月を自動検出
# ------------------------------------------------------------
def _detect_month_from_filename(filename: str) -> Optional[int]:
    """
    ファイル名から月を自動検出する。
    
    例:
    - "2025年1月_電気料金.pdf" → 1
    - "01_請求書.pdf" → 1
    - "電気_2025_01.pdf" → 1
    - "invoice_jan.pdf" → 1
    - "2025-01-15.pdf" → 1
    
    Returns:
        検出された月（1-12）、検出できない場合はNone
    """
    # パターン1: "1月" "01月" "１月"などの形式
    match = re.search(r'([0-9０-９]{1,2})\s*月', filename)
    if match:
        month_str = match.group(1)
        # 全角数字を半角に変換
        month_str = month_str.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
        month = int(month_str)
        if 1 <= month <= 12:
            return month
    
    # パターン2: "_01_" "2025-01" "-01." などの形式
    match = re.search(r'[_\-]([0-9]{2})[_\-\.]', filename)
    if match:
        month = int(match.group(1))
        if 1 <= month <= 12:
            return month
    
    # パターン3: 英語の月名
    month_names = {
        'jan': 1, 'january': 1,
        'feb': 2, 'february': 2,
        'mar': 3, 'march': 3,
        'apr': 4, 'april': 4,
        'may': 5,
        'jun': 6, 'june': 6,
        'jul': 7, 'july': 7,
        'aug': 8, 'august': 8,
        'sep': 9, 'september': 9,
        'oct': 10, 'october': 10,
        'nov': 11, 'november': 11,
        'dec': 12, 'december': 12,
    }
    
    filename_lower = filename.lower()
    for name, month in month_names.items():
        if name in filename_lower:
            return month
    
    return None


# ------------------------------------------------------------
# セッション初期化
# ------------------------------------------------------------
def _init_session_state() -> None:
    defaults = {
        "pdf_files": [],
        "output_file": "",
        "corp_name": "",          # 法人名
        "parse_mode": "single",   # "single" or "multi"
        "start_month": 10,        # 複数月PDFの開始月
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def _build_excel_filename() -> str:
    corp_name = (st.session_state.get("corp_name") or "").strip()
    if not corp_name:
        return "output.xlsx"

    invalid_chars = r'\\/:*?"<>|'
    for ch in invalid_chars:
        corp_name = corp_name.replace(ch, "")
    if not corp_name:
        corp_name = "output"

    return f"{corp_name}.xlsx"


# ------------------------------------------------------------
# メインページ
# ------------------------------------------------------------
def render_main_page(cfg: Dict[str, Any]) -> None:
    _init_session_state()

    st.markdown(
        """
        <style>
        .block-container {
            max-width: 100% !important;
            padding-left: 2rem !important;
            padding-right: 2rem !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("見積プロトタイプ｜PDF 明細 → テンプレExcelへ自動反映")

    left, mid, right = st.columns([4, 1.5, 4])

    ocr_service = OcrService(cfg)
    excel_service = ExcelService(cfg)

    # ① 左：法人名・モード・アップロード
    with left:
        st.subheader("① 法人名入力 & モード選択 & PDF アップロード")

        st.session_state.corp_name = st.text_input(
            "法人名（テンプレ B1 セルに反映）",
            value=st.session_state.get("corp_name", ""),
            placeholder="例：〇〇株式会社",
        )

        mode_label = st.radio(
            "PDFの構造",
            options=["1PDF = 1ヶ月分", "1PDFの中に複数月が含まれている"],
            horizontal=False,
        )
        if mode_label == "1PDF = 1ヶ月分":
            st.session_state.parse_mode = "single"
        else:
            st.session_state.parse_mode = "multi"

        if st.session_state.parse_mode == "multi":
            st.session_state.start_month = st.selectbox(
                "開始月（このPDFの最初のページが何月分か）",
                options=list(range(1, 13)),
                index=(st.session_state.get("start_month", 10) - 1),
                format_func=lambda m: f"{m}月",
            )

        pdf_files = st.file_uploader(
            "PDFをアップロード（複数選択可 / 1つずつでもOK）",
            type=["pdf"],
            accept_multiple_files=True,
            key="pdf_uploader",
        )

        if pdf_files is not None and len(pdf_files) > 0:
            st.session_state.pdf_files = []
            st.session_state.output_file = ""

            for f in pdf_files:
                # ファイル名から月を自動推定
                detected_month = _detect_month_from_filename(f.name)
                
                st.session_state.pdf_files.append(
                    {
                        "name": f.name,
                        "status": "未処理",
                        "invoice": None,
                        "text": "",
                        "bytes": f.read(),
                        "detected_month": detected_month,  # 自動検出した月
                        "selected_month": detected_month,  # ユーザーが選択する月
                    }
                )
        else:
            st.session_state.pdf_files = []
            st.session_state.output_file = ""
        
        # 単月モードの場合、アップロードしたファイルの月を選択
        if st.session_state.parse_mode == "single" and st.session_state.pdf_files:
            st.markdown("---")
            st.markdown("**📅 各PDFの月を指定してください**")
            
            for idx, file_info in enumerate(st.session_state.pdf_files):
                col1, col2 = st.columns([3, 1])
                
                with col1:
                    st.write(f"**{file_info['name']}**")
                    if file_info['detected_month']:
                        st.caption(f"自動検出: {file_info['detected_month']}月")
                
                with col2:
                    default_idx = (file_info['selected_month'] or 1) - 1
                    selected = st.selectbox(
                        "月",
                        options=list(range(1, 13)),
                        index=default_idx,
                        format_func=lambda m: f"{m}月",
                        key=f"month_select_{idx}_{file_info['name']}",
                    )
                    st.session_state.pdf_files[idx]['selected_month'] = selected

    # ② 真ん中：実行ボタン
    with mid:
        st.subheader("② 実行")

        has_files = len(st.session_state.pdf_files) > 0
        run_btn = st.button(
            "OCR → Excelテンプレートに反映",
            type="primary",
            use_container_width=True,
            disabled=not has_files,
        )

        if run_btn and has_files:
            _run_ocr_and_fill_excel(
                ocr_service,
                excel_service,
                corp_name=st.session_state.get("corp_name", "").strip(),
                mode=st.session_state.get("parse_mode", "single"),
                start_month=(
                    st.session_state.get("start_month")
                    if st.session_state.get("parse_mode") == "multi"
                    else None
                ),
            )

    # ③ 右：結果＆ダウンロード
    with right:
        st.subheader("③ 結果プレビュー・ダウンロード")
        _render_results_area()

    st.divider()
    st.caption("実行するとプロジェクト直下の `template_output.xlsx` を上書き保存します。")


# ------------------------------------------------------------
# OCR + Excel 書き込み
# ------------------------------------------------------------
def _run_ocr_and_fill_excel(
    ocr_service: OcrService,
    excel_service: ExcelService,
    corp_name: str = "",
    mode: str = "single",
    start_month: Optional[int] = None,
) -> None:
    st.session_state.output_file = ""

    invoices: List[Invoice] = []

    for idx, file_info in enumerate(st.session_state.pdf_files):
        st.session_state.pdf_files[idx]["status"] = "処理中"
        with st.spinner(f"🔄 {file_info['name']} をOCR実行中…"):
            try:
                # 単月モードの場合、ユーザーが選択した月を使用
                if mode == "single":
                    selected_month = file_info.get('selected_month')
                    # OCRでテキストを取得
                    invoice = ocr_service.analyze_invoice(
                        file_info["bytes"],
                        mode=mode,
                        start_month=None,
                    )
                    
                    # OCRテキストから直接kWh値を抽出
                    if selected_month and invoice.raw_text:
                        from ..services.ocr_service import OcrService
                        kwh_value = OcrService._extract_kwh_from_text(invoice.raw_text)
                        
                        if kwh_value:
                            # ユーザーが選択した月にkWh値を設定
                            invoice.fields = {f"{selected_month}月値": kwh_value}
                        else:
                            # kWh値が抽出できない
                            invoice.fields = {}
                            st.warning(f"⚠️ {file_info['name']} からkWh値を抽出できませんでした")
                else:
                    # 複数月モードの場合は従来通り
                    invoice = ocr_service.analyze_invoice(
                        file_info["bytes"],
                        mode=mode,
                        start_month=start_month,
                    )

                st.session_state.pdf_files[idx]["status"] = "完了"
                st.session_state.pdf_files[idx]["invoice"] = invoice
                st.session_state.pdf_files[idx]["text"] = invoice.raw_text or ""

                invoices.append(invoice)

                # デバッグ情報：抽出結果を表示
                month_info = f"（{file_info.get('selected_month')}月分）" if mode == "single" else ""
                fields_info = f" - フィールド: {invoice.fields}" if invoice.fields else " - フィールド: なし"
                st.success(f"✅ {file_info['name']} {month_info}の処理が完了しました{fields_info}")

            except Exception as e:
                st.session_state.pdf_files[idx]["status"] = "エラー"
                st.error(f"❌ エラー: {str(e)}")

    excel_path = excel_service.write_invoices(
        invoices,
        corp_name=corp_name,
    )
    st.session_state.output_file = excel_path


# ------------------------------------------------------------
# 結果表示
# ------------------------------------------------------------
def _render_results_area() -> None:
    output_path = st.session_state.get("output_file") or ""

    if output_path and Path(output_path).exists():
        with open(output_path, "rb") as f:
            st.download_button(
                label="テンプレExcel（上書き済み）をダウンロード",
                data=f.read(),
                file_name=_build_excel_filename(),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    if st.session_state.pdf_files:
        for file_info in st.session_state.pdf_files:
            st.write(f"**{file_info['name']}** - {file_info['status']}")
            if file_info["status"] == "完了":
                st.text_area(
                    "OCRテキスト",
                    file_info["text"],
                    height=150,
                    key=f"text_{file_info['name']}",
                )
            elif file_info["status"] == "エラー":
                st.write("エラーが発生しました。")
