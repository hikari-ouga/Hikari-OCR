from __future__ import annotations

from typing import Dict, Any, List
from pathlib import Path

import streamlit as st

from ..domain.invoice import Invoice
from ..services.ocr_service import OcrService
from ..services.excel_service import ExcelService


def _init_session_state() -> None:
    defaults = {
        "pdf_files": [],
        "output_file": "",
        "corp_name": "",   # 法人名
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def render_main_page(cfg: Dict[str, Any]) -> None:
    """
    メイン画面の描画（UIレイヤー）
    """
    _init_session_state()

    # ★ ページを画面いっぱいに広げる CSS
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

    # 3カラム構成
    left, mid, right = st.columns([4, 1.5, 4])

    # Service を生成
    ocr_service = OcrService(cfg)
    excel_service = ExcelService(cfg)

    # ------------------------------------------------------------
    # ① 法人名入力 & PDF アップロード
    # ------------------------------------------------------------
    with left:
        st.subheader("① 法人名入力 & PDF アップロード")

        # 法人名入力欄（Excel B1 に反映）
        st.session_state.corp_name = st.text_input(
            "法人名（テンプレ B1 セルに反映）",
            value=st.session_state.get("corp_name", ""),
            placeholder="例：〇〇株式会社",
        )

        # PDF アップロード
        pdf_files = st.file_uploader(
            "PDFをアップロード（複数選択可 / 一個ずつでもOK）",
            type=["pdf"],
            accept_multiple_files=True,
            key="pdf_uploader",
        )

        # ★ アップロード内容に応じて state を更新
        if pdf_files is not None and len(pdf_files) > 0:
            # 新しいファイルが来たので、前回の結果を完全リセット
            st.session_state.pdf_files = []
            st.session_state.output_file = ""

            for f in pdf_files:
                st.session_state.pdf_files.append(
                    {
                        "name": f.name,
                        "status": "未処理",
                        "invoice": None,   # Invoice オブジェクト
                        "text": "",
                        "bytes": f.read(),
                    }
                )
        else:
            # 何も選ばれていない状態なら、PDFリストと出力も空にしておく
            st.session_state.pdf_files = []
            st.session_state.output_file = ""

    # ------------------------------------------------------------
    # ② 実行ボタン
    # ------------------------------------------------------------
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
            )

    # ------------------------------------------------------------
    # ③ 結果プレビュー・ダウンロード
    # ------------------------------------------------------------
    with right:
        st.subheader("③ 結果プレビュー・ダウンロード")
        _render_results_area()

    st.divider()
    st.caption(
        "`template_output.xlsx` を直接上書き保存します。"
        " 新しいPDFをアップロードすると、前回の結果はリセットされます。"
    )


# ====================================================================
# OCR ＆ Excel 書き込み処理
# ====================================================================
def _run_ocr_and_fill_excel(
    ocr_service: OcrService,
    excel_service: ExcelService,
    corp_name: str = "",
) -> None:
    # ★ 実行のたびに前回の Excel パスをクリア
    st.session_state.output_file = ""

    invoices: List[Invoice] = []

    for idx, file_info in enumerate(st.session_state.pdf_files):
        st.session_state.pdf_files[idx]["status"] = "処理中"

        with st.spinner(f"🔄 {file_info['name']} をOCR実行中…"):
            try:
                invoice = ocr_service.analyze_invoice(file_info["bytes"])
                st.session_state.pdf_files[idx]["status"] = "完了"
                st.session_state.pdf_files[idx]["invoice"] = invoice
                st.session_state.pdf_files[idx]["text"] = invoice.raw_text or ""
                invoices.append(invoice)

                st.success(f"✅ {file_info['name']} の処理が完了しました")

            except Exception as e:
                st.session_state.pdf_files[idx]["status"] = "エラー"
                st.error(
                    f"❌ {file_info['name']} の処理中にエラー: {str(e)}"
                )

    # 法人名も渡して Excel 書き込み
    excel_path = excel_service.write_invoices(
        invoices,
        corp_name=corp_name,
    )

    st.session_state.output_file = excel_path


# ====================================================================
# 結果表示部分
# ====================================================================
def _render_results_area() -> None:
    output_path = st.session_state.get("output_file") or ""

    # Excel ダウンロードボタン
    if output_path and Path(output_path).exists():
        with open(output_path, "rb") as f:
            st.download_button(
                label="テンプレExcel（上書き済み）をダウンロード",
                data=f.read(),
                file_name="template_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # PDFごとの OCR テキスト
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
