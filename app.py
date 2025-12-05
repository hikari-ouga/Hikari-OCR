from __future__ import annotations

import streamlit as st

from app.config import load_app_config, init_env
from app.ui.main_page import render_main_page


def main() -> None:
    # .env 読み込みなど
    init_env()

    # config.json 読み込み
    cfg = load_app_config()

    # メイン画面描画
    render_main_page(cfg)


if __name__ == "__main__":
    # Streamlit から実行されるときはここは通らないけど、
    # python app.py でデバッグしたいとき用に残しておく
    main()
