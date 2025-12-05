from __future__ import annotations

from typing import Dict, Any, List
from pathlib import Path

import streamlit as st

from ..domain.invoice import Invoice
from ..services.ocr_service import OcrService
from ..services.excel_service import ExcelService


# ====================================================================
# セッション状態の初期化
# ====================================================================
def _init_session_state() -> None:
    defaults = {
        "pdf_files": [],
        "output_file": "",
        "corp_name": "",      # 法人名（Excel B1）
        "parse_mode": "single",  # 解析モード: "single" or "multi"
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def _inject_style() -> None:
    """UI用のCSSをまとめて注入"""
    st.markdown(
        """
        <style>
        /* ページ横幅・余白 */
        .block-container {
            max-width: 100% !important;
            padding-left: 2rem !important;
            padding-right: 2rem !important;
            padding-top: 2rem !important;
            padding-bottom: 3rem !important;
        }

        /* タイトル周り */
        .app-header-title {
            font-size: 2.1rem;
            font-weight: 800;
            letter-spacing: .04em;
            margin-bottom: 0.3rem;
        }
        .app-header-subtitle {
            color: #6b7280;
            font-size: 0.95rem;
        }

        /* ステップカード */
        .step-card {
            background-color: #ffffff;
            padding: 1.4rem 1.3rem;
            border-radius: 0.9rem;
            border: 1px solid #e5e7eb;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.03);
        }
        .step-title {
            font-weight: 700;
            font-size: 1.05rem;
            display: flex;
            align-items: center;
            gap: .4rem;
            margin-bottom: 0.3rem;
        }
        .step-pill {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 1.8rem;
            height: 1.8rem;
            border-radius: 999px;
            background: #eff6ff;
            color: #2563eb;
            font-size: 0.9rem;
            font-weight: 700;
        }
        .step-caption {
            color: #6b7280;
            font-size: 0.85rem;
            margin-bottom: 0.7rem;
        }

        /* PDFファイル一覧 */
        .file-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: .45rem .7rem;
            border-radius: .55rem;
            background: #f9fafb;
            border: 1px solid #e5e7eb;
            margin-bottom: .35rem;
            font-size: 0.9rem;
        }
        .file-name {
            font-weight: 500;
            overflow-wrap: anywhere;
        }
        .status-badge {
            padding: .18rem .6rem;
            border-radius: 999px;
            font-size: 0.75rem;
            font-weight: 600;
            color: white;
            white-space: nowrap;
        }

        /* ダウンロードボタンを少しだけ目立たせる */
        .stDownloadButton button {
            border-radius: 999px;
            font-weight: 600;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ====================================================================
# メイン画面
# ====================================================================
def render_main_page(cfg: Dict[str, Any]) -> None:
    """
    メイン画面の描画（UIレイヤー）
    """
    _init_session_state()
    _inject_style()

    # ヘッダー
    st.markdown(
        """
        <div class="app-header">
          <div class="app-header-title">
            見積プロトタイプ｜PDF 明細 → テンプレExcelへ自動反映
          </div>
          <div class="app-header-subtitle">
            明細PDFをアップロードして「実行」を押すだけで、
            あらかじめ用意した Excel テンプレートに月別の数値と法人名（B1）を自動反映します。
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # 3カラム構成
    left, mid, right = st.columns([4, 2, 4], gap="large")

    # Service を生成
    ocr_service = OcrService(cfg)
    excel_service = ExcelService(cfg)

    # ------------------------------------------------------------
    # ① 法人名入力 & PDF アップロード
    # ------------------------------------------------------------
    with left:
        st.markdown(
            """
            <div class="step-card">
              <div class="step-title">
                <div class="step-pill">1</div>
                <span>法人名入力 ＆ PDFアップロード</span>
              </div>
              <div class="step-caption">
                法人名はテンプレートの <b>B1セル</b> に反映されます。
                複数PDFをまとめてドラッグ＆ドロップできます。
              </div>
            """,
            unsafe_allow_html=True,
        )

        # 法人名入力欄
        st.session_state.corp_name = st.text_input(
            "法人名（テンプレ B1 セルに反映）",
            value=st.session_state.get("corp_name", ""),
            placeholder="例：〇〇株式会社",
        )

        # ★ 解析モード（1PDF=1ヶ月 or 1PDFに複数月）
        mode_label = st.radio(
            "PDF の構造",
            options=["1PDF = 1ヶ月分", "1PDFの中に複数月が含まれている"],
            horizontal=False,
        )
        if mode_label == "1PDF = 1ヶ月分":
            st.session_state.parse_mode = "single"
        else:
            st.session_state.parse_mode = "multi"

        # PDF アップロード
        pdf_files = st.file_uploader(
            "PDFをアップロード（複数選択可 / 一個ずつでもOK）",
            type=["pdf"],
            accept_multiple_files=True,
            key="pdf_uploader",
        )

        # アップロード内容に応じて state を更新
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

            st.success(f"{len(pdf_files)} 件のPDFを読み込みました。")

        else:
            # 何も選ばれていない状態なら、PDFリストと出力も空にしておく
            st.session_state.pdf_files = []
            st.session_state.output_file = ""

        st.markdown("</div>", unsafe_allow_html=True)  # step-card close

    # ------------------------------------------------------------
    # ② 実行ボタン
    # ------------------------------------------------------------
    with mid:
        st.markdown(
            """
            <div class="step-card">
              <div class="step-title">
                <div class="step-pill">2</div>
                <span>OCR実行 ＆ Excel反映</span>
              </div>
              <div class="step-caption">
                Azure Document Intelligence で PDF を解析し、
                プロジェクト直下の <code>template_output.xlsx</code> に月別値と法人名を書き込みます。
              </div>
            """,
            unsafe_allow_html=True,
        )

        has_files = len(st.session_state.pdf_files) > 0

        run_btn = st.button(
            "🔄 OCR → Excelテンプレートに反映",
            type="primary",
            use_container_width=True,
            disabled=not has_files,
        )

        if not has_files:
            st.info("左でPDFをアップロードすると実行できるようになります。")

        if run_btn and has_files:
            _run_ocr_and_fill_excel(
                ocr_service,
                excel_service,
                corp_name=st.session_state.get("corp_name", "").strip(),
                mode=st.session_state.get("parse_mode", "single"),
            )

        st.markdown("</div>", unsafe_allow_html=True)  # step-card close

    # ------------------------------------------------------------
    # ③ 結果プレビュー・ダウンロード
    # ------------------------------------------------------------
    with right:
        st.markdown(
            """
            <div class="step-card">
              <div class="step-title">
                <div class="step-pill">3</div>
                <span>結果プレビュー ＆ ダウンロード</span>
              </div>
              <div class="step-caption">
                上書き済みの Excel をダウンロードし、PDFごとの OCR テキストも確認できます。
              </div>
            """,
            unsafe_allow_html=True,
        )

        _render_results_area()

        st.markdown("</div>", unsafe_allow_html=True)  # step-card close

    st.markdown("---")
    st.caption(
        "`template_output.xlsx` を <b>直接上書き保存</b> します。"
        " 新しいPDFをアップロードすると、前回の結果はリセットされます。",
        unsafe_allow_html=True,
    )


# ====================================================================
# OCR ＆ Excel 書き込み処理
# ====================================================================
def _run_ocr_and_fill_excel(
    ocr_service: OcrService,
    excel_service: ExcelService,
    corp_name: str = "",
    mode: str = "single",
) -> None:
    # 実行のたびに前回の Excel パスをクリア
    st.session_state.output_file = ""

    invoices: List[Invoice] = []

    for idx, file_info in enumerate(st.session_state.pdf_files):
        st.session_state.pdf_files[idx]["status"] = "処理中"

        with st.spinner(f"🔄 {file_info['name']} をOCR実行中…"):
            try:
                # ★ 解析モードを渡す（single / multi）
                invoice = ocr_service.analyze_invoice(file_info["bytes"], mode=mode)
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

    # 法人名も渡して Excel 書き込み（ExcelService側でB1に反映）
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
                label="📥 テンプレExcel（上書き済み）をダウンロード",
                data=f.read(),
                file_name="template_output.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )

    files = st.session_state.get("pdf_files", [])
    if not files:
        st.info("まだ処理結果がありません。左側でPDFをアップロードして実行してください。")
        return

    # ステータス別カラー
    status_colors = {
        "未処理": "#9ca3af",
        "処理中": "#f59e0b",
        "完了": "#10b981",
        "エラー": "#ef4444",
    }

    st.markdown("##### 処理状況")

    # PDFごとのステータスを一覧表示
    for file_info in files:
        color = status_colors.get(file_info["status"], "#6b7280")
        st.markdown(
            f"""
            <div class="file-row">
                <span class="file-name">{file_info['name']}</span>
                <span class="status-badge" style="background:{color};">
                    {file_info['status']}
                </span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # OCRテキストは expander で必要なときだけ開く
    for file_info in files:
        if file_info["status"] == "完了":
            with st.expander(f"🔍 {file_info['name']} の OCR テキストを表示"):
                st.text_area(
                    "OCRテキスト",
                    file_info["text"],
                    height=180,
                    key=f"text_{file_info['name']}",
                )
        elif file_info["status"] == "エラー":
            st.warning(f"{file_info['name']}：エラーが発生しました。")
