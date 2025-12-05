from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict

from dotenv import load_dotenv

# このファイルは estimate-proto/estimate_proto/config.py にある想定
# → プロジェクトルートは parents[1]
PROJECT_ROOT: Path = Path(__file__).resolve().parents[1]


def get_project_root() -> Path:
    """プロジェクトルート (estimate-proto/) を返す。"""
    return PROJECT_ROOT


def init_env() -> None:
    """
    .env を読み込む。
    - プロジェクトルート直下の .env を対象とする
    """
    env_path = PROJECT_ROOT / ".env"
    if env_path.exists():
        load_dotenv(env_path)
    else:
        # .env が無くてもエラーにはしない（Render などでは環境変数で渡す想定）
        load_dotenv()


def load_app_config(path: str | None = None) -> Dict[str, Any]:
    """
    config.json を読み込む。
    引数 path が None の場合、プロジェクトルート直下の config.json を読む。
    """
    if path is None:
        path = str(PROJECT_ROOT / "config.json")

    with open(path, encoding="utf-8") as f:
        return json.load(f)
