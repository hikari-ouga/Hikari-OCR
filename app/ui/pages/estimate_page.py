from __future__ import annotations

import json
import logging
from typing import List, Optional
from pathlib import Path

from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel

from ...domain.invoice import Invoice
from ...services.ocr_service import OcrService
from ...services.excel_service import ExcelService
from ...config import load_app_config

# ロガー設定
logger = logging.getLogger(__name__)

router = APIRouter()

# グローバル変数でExcelファイルパスを保持
_last_excel_path: Optional[str] = None

# リクエストボディのモデル定義
class ExcelGenerationRequest(BaseModel):
    corp_name: str
    invoices_data: List[dict]  # OCR結果のリスト

@router.post("/process")
async def process_pdfs(
    corp_name: str = Form(...),
    address: str = Form(""),
    corp_number: str = Form(""),
    mode: str = Form(...),
    start_month: Optional[int] = Form(None),
    month_order: Optional[str] = Form("ascending"),
    month_mappings: str = Form(...),
    files: List[UploadFile] = File(...)
):
    """
    PDFファイルをOCR処理してExcelに反映する
    
    Args:
        corp_name: 法人名
        mode: "single" or "multi"
        start_month: 複数月モードの開始月
        month_order: 月の並び順 "ascending" (昇順) or "descending" (降順)
        month_mappings: JSON文字列 [{"filename": "xxx.pdf", "selectedMonth": 1}, ...]
        files: PDFファイルリスト
    
    Returns:
        処理結果のJSON
    """
    global _last_excel_path
    
    try:
        # 設定読み込み
        cfg = load_app_config()
        ocr_service = OcrService(cfg)
        excel_service = ExcelService(cfg)
        
        # 月マッピング情報をパース
        try:
            mappings = json.loads(month_mappings)
            month_map = {m['filename']: m['selectedMonth'] for m in mappings}
        except Exception as e:
            logger.error(f"月マッピングのパースに失敗: {e}")
            month_map = {}
        
        invoices: List[Invoice] = []
        results = []
        
        for file in files:
            try:
                # ファイル内容を読み込み
                content = await file.read()
                
                logger.info(f"処理開始: {file.filename}, mode={mode}")
                
                if mode == "single":
                    # 単月モード
                    selected_month = month_map.get(file.filename, 1)
                    
                    # OCRでテキスト取得
                    invoice = ocr_service.analyze_invoice(
                        content,
                        mode="single",
                        start_month=None
                    )
                    
                    # OCRテキストから直接kWh値を抽出
                    kwh_value = ""
                    ocr_confidence = invoice.fields.get("ocr_confidence", 0) if invoice.fields else 0
                    
                    if selected_month and invoice.raw_text:
                        kwh_value = OcrService._extract_kwh_from_text(invoice.raw_text)
                        
                        if kwh_value:
                            invoice.fields = {f"{selected_month}月値": kwh_value, "ocr_confidence": ocr_confidence}
                            logger.info(f"{file.filename}: {selected_month}月値={kwh_value}, 信頼度={ocr_confidence:.2f}")
                        else:
                            invoice.fields = {"ocr_confidence": ocr_confidence}
                            logger.warning(f"{file.filename}: kWh値を抽出できませんでした")
                    
                    invoices.append(invoice)
                    results.append({
                        "filename": file.filename,
                        "status": "完了" if kwh_value else "kWh未検出",
                        "fields": invoice.fields,
                        "kwh": kwh_value,
                        "ocr_text": invoice.raw_text,
                        "ocr_confidence": ocr_confidence
                    })
                    
                else:
                    # 複数月モード
                    invoice = ocr_service.analyze_invoice(
                        content,
                        mode="multi",
                        start_month=start_month,
                        month_order=month_order
                    )
                    
                    invoices.append(invoice)
                    results.append({
                        "filename": file.filename,
                        "status": "完了",
                        "fields": invoice.fields
                    })
                
            except Exception as e:
                logger.error(f"{file.filename}の処理中にエラー: {str(e)}", exc_info=True)
                results.append({
                    "filename": file.filename,
                    "status": "エラー",
                    "error": str(e)
                })
        
        # Excelに書き込み
        try:
            excel_path = excel_service.write_invoices(
                invoices, 
                corp_name=corp_name,
                address=address,
                corp_number=corp_number
            )
            _last_excel_path = excel_path
            logger.info(f"Excel書き込み完了: {excel_path}")
        except Exception as e:
            logger.error(f"Excel書き込みエラー: {str(e)}", exc_info=True)
            raise HTTPException(status_code=500, detail=f"Excel書き込みエラー: {str(e)}")
        
        return JSONResponse({
            "success": True,
            "results": results,
            "excel_path": excel_path
        })
        
    except Exception as e:
        logger.error(f"処理エラー: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/download")
async def download_excel(
    corp_name: str = "",
    address: str = "",
    corp_number: str = ""
):
    """
    生成されたExcelファイルをダウンロード
    ダウンロード時に最新の住所・法人番号を更新
    """
    global _last_excel_path
    
    if not _last_excel_path or not Path(_last_excel_path).exists():
        raise HTTPException(status_code=404, detail="Excelファイルが見つかりません")
    
    # 住所または法人番号が指定されている場合、Excelファイルを更新
    if address or corp_number:
        try:
            from openpyxl import load_workbook
            wb = load_workbook(_last_excel_path)
            ws = wb.active
            
            # 住所をB2に更新
            if address:
                ws['B2'] = address
            
            # 法人番号をB4に更新
            if corp_number:
                ws['B4'] = corp_number
            
            wb.save(_last_excel_path)
            logger.info(f"Excelファイル更新: 住所={address}, 法人番号={corp_number}")
        except Exception as e:
            logger.warning(f"Excelファイルの更新に失敗: {e}")
    
    return FileResponse(
        _last_excel_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=Path(_last_excel_path).name
    )


@router.post("/ocr_single")
async def process_single_pdf(
    mode: str = Form(...),
    selected_month: Optional[int] = Form(None),
    start_month: Optional[int] = Form(None),
    file: UploadFile = File(...)
):
    """
    単一のPDFファイルをOCR処理して結果を返す
    """
    try:
        # 設定読み込み
        cfg = load_app_config()
        ocr_service = OcrService(cfg)
        
        content = await file.read()
        logger.info(f"単一処理開始: {file.filename}, mode={mode}, month={selected_month}")
        
        invoice = None
        kwh_value = None
        
        if mode == "single":
            # 単月モード
            invoice = ocr_service.analyze_invoice(
                content,
                mode="single",
                start_month=None
            )
            
            # OCRテキストから直接kWh値を抽出
            if selected_month and invoice.raw_text:
                kwh_value = OcrService._extract_kwh_from_text(invoice.raw_text)
                if kwh_value:
                    invoice.fields = {f"{selected_month}月値": kwh_value}
                else:
                    invoice.fields = {}
        else:
            # 複数月モード
            invoice = ocr_service.analyze_invoice(
                content,
                mode="multi",
                start_month=start_month
            )
            # 複数月モードの場合はfieldsに既にデータが入っているはず
        
        return JSONResponse({
            "success": True,
            "filename": file.filename,
            "fields": invoice.fields,
            "raw_text": invoice.raw_text # デバッグ用
        })

    except Exception as e:
        logger.error(f"単一処理エラー: {str(e)}", exc_info=True)
        return JSONResponse({
            "success": False,
            "filename": file.filename,
            "error": str(e)
        }, status_code=500)


@router.post("/generate_excel")
async def generate_excel(request: ExcelGenerationRequest):
    """
    OCR結果のリストを受け取りExcelを生成する
    """
    global _last_excel_path
    
    try:
        cfg = load_app_config()
        excel_service = ExcelService(cfg)
        
        invoices = []
        for data in request.invoices_data:
            # 辞書データからInvoiceオブジェクトを復元（簡易的）
            inv = Invoice(raw_text="")
            inv.fields = data.get("fields", {})
            invoices.append(inv)
            
        excel_path = excel_service.write_invoices(invoices, corp_name=request.corp_name)
        _last_excel_path = excel_path
        
        return JSONResponse({
            "success": True,
            "excel_path": excel_path
        })
        
    except Exception as e:
        logger.error(f"Excel生成エラー: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))
