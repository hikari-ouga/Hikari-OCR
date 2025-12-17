from __future__ import annotations

import os
import io
import logging
from typing import Dict, Any, List, Optional, Tuple

from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from azure.core.exceptions import HttpResponseError
from pdf2image import convert_from_bytes
from PIL import Image, ImageEnhance, ImageFilter
import numpy as np

from ..domain.invoice import Invoice

# ロガー設定
logger = logging.getLogger(__name__)


class OcrService:
    """
    Azure Document Intelligence を使って PDF を解析し、

    - 単月モード: 1PDF = 1ヶ月分
    - 複数月モード: 12 / 24 / 36 ページ → 開始月から割り当てて12ヶ月分抽出

    の Invoice を生成するサービス。
    
    1つのモデル（prebuilt-invoice）のみを使用し、
    元PDFで失敗した場合は前処理PDFで再試行する。
    """

    # 試行するモデルの優先順位
    MODELS_TO_TRY = [
        "prebuilt-invoice",      # 請求書特化
    ]

    def __init__(self, cfg: Dict[str, Any]) -> None:
        self.cfg = cfg

        # 環境変数固定
        endpoint = os.getenv("AZURE_FORMREC_ENDPOINT")
        key = os.getenv("AZURE_FORMREC_KEY")

        if not endpoint or not key:
            raise ValueError(
                "Form Recognizer の endpoint / key が見つかりません。\n"
                "環境変数 AZURE_FORMREC_ENDPOINT と AZURE_FORMREC_KEY を確認してください。"
            )

        self.client = DocumentAnalysisClient(
            endpoint=endpoint,
            credential=AzureKeyCredential(key),
        )

        # デフォルトモデル
        self.model_id: str = cfg.get("FORM_RECOGNIZER_MODEL_ID", "prebuilt-invoice")

    # --------------------------------------------------------
    # PDF前処理：色付きPDFの改善
    # --------------------------------------------------------
    def _preprocess_pdf(self, pdf_bytes: bytes) -> bytes:
        """
        色付きPDFを前処理して、OCR精度を向上させる。
        
        処理内容:
        1. PDFを画像に変換
        2. グレースケール化
        3. コントラスト強調
        4. シャープネス調整
        5. 二値化（必要に応じて）
        6. 画像をPDFに再変換
        
        Args:
            pdf_bytes: 元のPDFバイト列
            
        Returns:
            前処理済みPDFバイト列
        """
        try:
            logger.info("PDF前処理開始")
            
            # PDFを画像に変換（300dpi推奨）
            images = convert_from_bytes(
                pdf_bytes,
                dpi=300,
                fmt='PNG',
                grayscale=False  # カラーで読み込んでから処理
            )
            
            processed_images = []
            
            for i, img in enumerate(images):
                logger.info(f"ページ {i+1}/{len(images)} を前処理中...")
                
                # 1. グレースケール化
                if img.mode != 'L':
                    img = img.convert('L')
                
                # 2. コントラスト強調（1.5倍）
                enhancer = ImageEnhance.Contrast(img)
                img = enhancer.enhance(1.5)
                
                # 3. シャープネス調整
                enhancer = ImageEnhance.Sharpness(img)
                img = enhancer.enhance(1.3)
                
                # 4. ノイズ除去（メディアンフィルタ）
                img = img.filter(ImageFilter.MedianFilter(size=3))
                
                # 5. 適応的二値化（背景色の影響を軽減）
                img_array = np.array(img)
                
                # Otsuの二値化を適用
                from PIL import ImageOps
                img = ImageOps.autocontrast(img, cutoff=2)
                
                processed_images.append(img)
            
            # 前処理済み画像をPDFに変換
            pdf_buffer = io.BytesIO()
            
            if len(processed_images) == 1:
                processed_images[0].save(pdf_buffer, format='PDF', resolution=300.0)
            else:
                processed_images[0].save(
                    pdf_buffer,
                    format='PDF',
                    save_all=True,
                    append_images=processed_images[1:],
                    resolution=300.0
                )
            
            processed_pdf = pdf_buffer.getvalue()
            logger.info(f"PDF前処理完了: {len(pdf_bytes)} → {len(processed_pdf)} bytes")
            
            return processed_pdf
            
        except Exception as e:
            logger.warning(f"PDF前処理でエラー: {str(e)}, 元のPDFを使用します")
            return pdf_bytes

    # --------------------------------------------------------
    # 公開メソッド：単月 / 複数月モードの切り替え
    # --------------------------------------------------------
    def analyze_invoice(
        self,
        content: bytes,
        mode: str = "single",
        start_month: Optional[int] = None,
        month_order: str = "ascending",
    ) -> Invoice:

        if mode == "multi":
            if start_month is None:
                raise ValueError("複数月モードでは start_month が必須です。")
            return self._analyze_multi(content, start_month, month_order)

        # デフォルトは単月モード
        return self._analyze_single(content)

    # --------------------------------------------------------
    # Azure Document Intelligence API 呼び出し（複数モデルでフォールバック）
    # --------------------------------------------------------
    def _call_azure_ocr_with_fallback(self, content: bytes) -> Tuple[Any, str]:
        """
        複数のモデルを順番に試行し、最初に成功したものを返す。
        元のPDFで失敗した場合、前処理版でも試行する。
        
        Returns:
            (result, model_id): OCR結果とモデルID
        
        Raises:
            Exception: すべてのモデルで失敗した場合
        """
        errors = []
        
        # まず元のPDFで全モデルを試行
        for model_id in self.MODELS_TO_TRY:
            try:
                logger.info(f"OCR試行開始（元PDF）: モデル={model_id}, PDFサイズ={len(content)} bytes")
                
                # 日本語認識を明示的に指定
                analyze_kwargs = {
                    "model_id": model_id,
                    "document": content,
                }
                
                # prebuilt-readの場合のみlanguagesパラメータを追加
                if model_id == "prebuilt-read":
                    analyze_kwargs["languages"] = ["ja"]
                else:
                    analyze_kwargs["locale"] = "ja-JP"
                
                poller = self.client.begin_analyze_document(**analyze_kwargs)
                result = poller.result()
                
                # デバッグ：抽出テキストの内容を確認
                if result and result.content:
                    content_preview = result.content[:200].replace("\n", "\\n")
                    logger.info(f"抽出テキスト（先頭200文字）: {content_preview}")
                    
                    # 日本語文字の検出チェック
                    import re
                    japanese_chars = re.findall(r'[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF]', result.content)
                    logger.info(f"日本語文字数: {len(japanese_chars)}, 総文字数: {len(result.content)}")
                
                if result and result.content and len(result.content.strip()) > 50:
                    logger.info(f"OCR成功（元PDF）: モデル={model_id}, 抽出文字数={len(result.content)}")
                    return result, model_id
                else:
                    msg = f"{model_id}(元PDF): 結果が不十分（文字数={len(result.content) if result and result.content else 0}）"
                    logger.warning(msg)
                    errors.append(msg)
                    
            except HttpResponseError as e:
                error_msg = f"{model_id}(元PDF): HTTP {e.status_code} - {e.message}"
                logger.warning(f"OCR失敗: {error_msg}")
                errors.append(error_msg)
                
            except Exception as e:
                error_msg = f"{model_id}(元PDF): {type(e).__name__} - {str(e)}"
                logger.warning(f"OCR失敗: {error_msg}")
                errors.append(error_msg)
        
        # 元のPDFで全て失敗 → 前処理版で再試行
        logger.info("元PDFで失敗したため、前処理版PDFで再試行します...")
        preprocessed_content = self._preprocess_pdf(content)
        
        for model_id in self.MODELS_TO_TRY:
            try:
                logger.info(f"OCR試行開始（前処理PDF）: モデル={model_id}")
                
                # 日本語認識を明示的に指定
                analyze_kwargs = {
                    "model_id": model_id,
                    "document": preprocessed_content,
                }
                
                # prebuilt-readの場合のみlanguagesパラメータを追加
                if model_id == "prebuilt-read":
                    analyze_kwargs["languages"] = ["ja"]
                else:
                    analyze_kwargs["locale"] = "ja-JP"
                
                poller = self.client.begin_analyze_document(**analyze_kwargs)
                result = poller.result()
                
                # デバッグ：抽出テキストの内容を確認
                if result and result.content:
                    content_preview = result.content[:200].replace("\n", "\\n")
                    logger.info(f"抽出テキスト（先頭200文字）: {content_preview}")
                    
                    # 日本語文字の検出チェック
                    import re
                    japanese_chars = re.findall(r'[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF]', result.content)
                    logger.info(f"日本語文字数: {len(japanese_chars)}, 総文字数: {len(result.content)}")
                
                if result and result.content and len(result.content.strip()) > 50:
                    logger.info(f"OCR成功（前処理PDF）: モデル={model_id}, 抽出文字数={len(result.content)}")
                    return result, f"{model_id}(前処理)"
                else:
                    msg = f"{model_id}(前処理PDF): 結果が不十分（文字数={len(result.content) if result and result.content else 0}）"
                    logger.warning(msg)
                    errors.append(msg)
                    
            except HttpResponseError as e:
                error_msg = f"{model_id}(前処理PDF): HTTP {e.status_code} - {e.message}"
                logger.warning(f"OCR失敗: {error_msg}")
                errors.append(error_msg)
                
            except Exception as e:
                error_msg = f"{model_id}(前処理PDF): {type(e).__name__} - {str(e)}"
                logger.warning(f"OCR失敗: {error_msg}")
                errors.append(error_msg)
        
        # すべてのモデル（元PDF+前処理PDF）で失敗
        error_summary = "\n".join([f"  - {err}" for err in errors])
        raise Exception(
            f"すべてのOCRモデル（元PDF + 前処理PDF）で読み取りに失敗しました:\n{error_summary}\n\n"
            f"PDFファイルの形式や内容を確認してください。"
        )

    # --------------------------------------------------------
    # 単月モード：1PDF = 1ヶ月分
    # --------------------------------------------------------
    def _analyze_single(self, content: bytes) -> Invoice:
        """単月モードでPDFを解析（複数モデルでフォールバック）"""
        try:
            result, used_model = self._call_azure_ocr_with_fallback(content)
            
            # ✅ SDK v3 系: 全文テキストは result.content に入っている
            full_text = result.content or ""
            
            if not full_text.strip():
                logger.error("OCRでテキストを抽出できませんでした")
                raise ValueError("PDFからテキストを抽出できませんでした。画像のみのPDFの可能性があります。")
            
            logger.info(f"単月モード解析完了: モデル={used_model}, 文字数={len(full_text)}")
            
            # 信頼度を計算（全単語の平均confidence）
            total_confidence = 0
            word_count = 0
            for page in result.pages:
                if hasattr(page, 'words') and page.words:
                    for word in page.words:
                        if hasattr(word, 'confidence') and word.confidence is not None:
                            total_confidence += word.confidence
                            word_count += 1
            
            average_confidence = total_confidence / word_count if word_count > 0 else 0
            logger.info(f"OCR平均信頼度: {average_confidence:.2f}")
            
            # 月の割り当てはUI側で行うため、ここではraw_textと信頼度を返す
            invoice = Invoice(
                fields={"ocr_confidence": average_confidence},
                raw_text=full_text
            )
            return invoice
            
        except Exception as e:
            logger.error(f"単月モード解析エラー: {str(e)}", exc_info=True)
            raise

    # --------------------------------------------------------
    # 複数月モード：12 / 24 / 36ページを開始月から割り当てて12ヶ月分生成
    # --------------------------------------------------------
    def _analyze_multi(self, content: bytes, start_month: int, month_order: str = "ascending") -> Invoice:
        """複数月モードでPDFを解析（複数モデルでフォールバック）"""
        try:
            result, used_model = self._call_azure_ocr_with_fallback(content)
            
            # ✅ ページごとのテキストは result.content から spans で切り出す
            page_texts: List[str] = []
            full_content = result.content or ""
            
            if not full_content.strip():
                logger.error("OCRでテキストを抽出できませんでした")
                raise ValueError("PDFからテキストを抽出できませんでした。画像のみのPDFの可能性があります。")

            # 信頼度を計算（全単語の平均confidence）
            total_confidence = 0
            word_count = 0
            for page in result.pages:
                if hasattr(page, 'words') and page.words:
                    for word in page.words:
                        if hasattr(word, 'confidence') and word.confidence is not None:
                            total_confidence += word.confidence
                            word_count += 1

            average_confidence = total_confidence / word_count if word_count > 0 else 0
            logger.info(f"OCR平均信頼度: {average_confidence:.2f}")

            for page in result.pages:
                if page.spans:
                    # 通常1ページ1spanなので先頭だけ取ればOK
                    span = page.spans[0]
                    start = span.offset
                    end = span.offset + span.length
                    page_texts.append(full_content[start:end])
                else:
                    page_texts.append("")

            num_pages = len(page_texts)
            
            logger.info(f"複数月モード解析完了: モデル={used_model}, ページ数={num_pages}")

            if num_pages not in (12, 24, 36):
                raise ValueError(
                    f"ページ枚数が違います。複数月モードは 12枚 / 24枚 / 36枚 のみ対応しています（実際: {num_pages}枚）"
                )

            pages_per_month = num_pages // 12  # 12→1、24→2、36→3
            fields: Dict[str, str] = {}

            fields["ocr_confidence"] = average_confidence

            current_month = start_month

            for i in range(12):
                start_idx = i * pages_per_month
                end_idx = start_idx + pages_per_month
                month_text = "\n".join(page_texts[start_idx:end_idx])

                # kWh 抽出（単月と同じロジック）
                kwh_value = self._extract_kwh_from_text(month_text)
                if kwh_value:
                    # ★ -1 のオフセットを削除: start_month=10なら10月として扱う
                    fields[f"{current_month}月値"] = kwh_value

                # 月の進め方を month_order に応じて切り替え
                if month_order == "descending":
                    current_month = self._prev_month(current_month)
                else:
                    current_month = self._next_month(current_month)

            full_text = "\n".join(page_texts)
            return Invoice(fields=fields, raw_text=full_text)
            
        except Exception as e:
            logger.error(f"複数月モード解析エラー: {str(e)}", exc_info=True)
            raise

    # --------------------------------------------------------
    # kWh 抽出（4桁以上限定 + カンマ後スペース対応）
    # --------------------------------------------------------
    @staticmethod
    def _extract_kwh_from_text(text: str) -> str:
        import re

        logger.info(f"=== kWh抽出開始 ===")
        logger.info(f"テキスト全体（先頭500文字）:\n{text[:500]}")
        
        # 全角を半角に統一（数字、カンマ、スペース、kWh）
        text = text.translate(str.maketrans({
            '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
            '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
            '，': ',', '、': ',', '　': ' ',
            'ｋ': 'k', 'ｋ': 'k', 'Ｋ': 'K',
            'ｗ': 'w', 'ｗ': 'w', 'Ｗ': 'W',
            'ｈ': 'h', 'ｈ': 'h', 'Ｈ': 'H',
        }))
        
        # デバッグ：kWh周辺のテキストを可視化（全角対応前後）
        kwh_contexts = re.findall(r".{0,50}[kKｋＫ]\s*[wWｗＷ]\s*[hHｈＨ].{0,10}", text, flags=re.IGNORECASE)
        logger.info(f"=== kWh周辺テキスト（全{len(kwh_contexts)}箇所）===")
        for i, ctx in enumerate(kwh_contexts, 1):
            visible = ctx.replace(" ", "␣").replace("\n", "↵").replace("\r", "⏎").replace("\t", "⇥").replace(",", "⸴")
            logger.info(f"  [{i}] {visible}")
        
        # 【重要】複数の改行パターンに対応（\r\n, \n, \r）
        # まず統一的な改行に変換
        normalized_text = text.replace('\r\n', '\n').replace('\r', '\n')
        
        # 改行で分割してkWhを含む行だけを抽出（スペースで分割されている場合も対応）
        lines = normalized_text.split('\n')
        
        # kWhパターンを柔軟にマッチング: "kWh", "k Wh", "k  W h" など
        kwh_pattern = r'k\s*[wW]\s*[hH]'
        kwh_lines = [line for line in lines if re.search(kwh_pattern, line, re.IGNORECASE)]
        
        logger.info(f"=== kWhを含む行（全{len(kwh_lines)}行）===")
        
        all_nums = []
        
        for i, line in enumerate(kwh_lines, 1):
            line_visible = line.replace(" ", "␣").replace(",", "⸴")
            logger.info(f"  [{i}] 行: '{line_visible}'")
            
            # この行から "数値 k Wh" のパターンを抽出
            # 例: "207,624kWh" → "207,624"
            # 例: "284,077 k Wh" → "284,077"
            # 例: "2,915 (kWh)" → "2,915"
            # 例: "2,915（kWh）" → "2,915" (全角括弧)
            match = re.search(r'([\d\s,\.]+)\s*[\(\[（]?\s*k\s*[wW]\s*[hH]\s*[\)\]）]?', line, flags=re.IGNORECASE)
            
            if not match:
                logger.warning(f"  [{i}] スキップ（パターンなし）")
                continue
            
            raw = match.group(1).strip()
            
            # スペースを削除（例: "284 077" → "284077"）
            raw_no_space = raw.replace(' ', '')
            
            # カンマの後のスペースを削除（例: "14, 662" → "14,662"）
            raw_normalized = re.sub(r',\s+', ',', raw_no_space)
            
            # 数字とカンマ以外を削除
            cleaned = re.sub(r'[^\d,]', '', raw_normalized)
            
            # カンマを削除して純粋な数値に
            final_num_str = cleaned.replace(',', '')
            
            logger.info(f"  [{i}] 元: '{raw}' → スペース除去: '{raw_no_space}' → 正規化: '{raw_normalized}' → 数値: '{final_num_str}'")
            
            if not final_num_str:
                logger.warning(f"  [{i}] スキップ（空）")
                continue
            
            try:
                v = int(final_num_str)
                
                # kWhは最低4桁以上（999以下は無視）
                if v < 1000:
                    logger.warning(f"  [{i}] スキップ（3桁以下: {v}）")
                    continue
                
                all_nums.append(v)
                logger.info(f"  [{i}] ✓ 有効: {v}")
                
            except Exception as e:
                logger.warning(f"  [{i}] エラー: {e}")
        
        if not all_nums:
            logger.error("❌ kWh未検出（4桁以上の値がありません）")
            # より詳細なデバッグ情報
            logger.error(f"改行で分割した行数: {len(lines)}")
            logger.error(f"kWhを含む行数: {len(kwh_lines)}")
            if kwh_lines:
                logger.error(f"kWhを含む行の例: {kwh_lines[:3]}")
            return ""
        
        # 重複除去して最大値を採用
        unique_nums = sorted(list(set(all_nums)), reverse=True)
        max_val = unique_nums[0]
        
        logger.info(f"=== 最終結果 ===")
        logger.info(f"  全候補: {all_nums}")
        logger.info(f"  ユニーク（降順）: {unique_nums}")
        logger.info(f"  ✅ 採用: {max_val}")
        
        return str(max_val)

    @staticmethod
    def _next_month(month: int) -> int:
        return 1 if month == 12 else month + 1

    @staticmethod
    def _prev_month(month: int) -> int:
        return 12 if month == 1 else month - 1

