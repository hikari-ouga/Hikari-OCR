from __future__ import annotations

import os
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import uvicorn

from app.config import load_app_config, init_env


def create_app() -> FastAPI:
    """FastAPIアプリケーションを作成"""
    # .env 読み込み
    init_env()
    
    # config.json 読み込み（後続の処理で使用）
    cfg = load_app_config()
    
    # FastAPIアプリ作成
    app = FastAPI(
        title="見積プロトタイプ - PDF→Excel自動反映",
        description="PDFファイルをOCR処理してExcelテンプレートに自動反映するアプリケーション",
        version="2.0.0"
    )
    
    # 静的ファイル配信
    app.mount("/styles", StaticFiles(directory="app/ui/styles"), name="styles")
    app.mount("/scripts", StaticFiles(directory="app/ui/scripts"), name="scripts")
    
    # APIルーター読み込み
    from app.ui.pages.estimate_page import router
    app.include_router(router, prefix="/api", tags=["estimate"])
    
    # トップページ
    @app.get("/", tags=["frontend"])
    async def root():
        """メインページを返す"""
        return FileResponse("app/ui/templates/index.html")
    
    return app


# FastAPIアプリインスタンス
app = create_app()


if __name__ == "__main__":
    # python app.py で直接起動する場合
    port = int(os.environ.get("PORT", 8000))  # Render対応: 環境変数PORTを使用
    print("🚀 サーバーを起動中...")
    print(f"📱 ブラウザで http://localhost:{port} にアクセスしてください")
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=port,
        log_level="info"
    )
