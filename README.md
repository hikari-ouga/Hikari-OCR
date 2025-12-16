# 見積もりアシスト（Hikari OCR）

PDF明細書から電力使用量（kWh）を自動抽出し、Excelテンプレートに転記するWebアプリケーションです。Azure Document Intelligenceを使用した高精度OCRにより、営業見積もり業務を効率化します。

## 主な機能

- 📄 **PDF自動読み取り**: 電力会社の請求書PDFから使用電力量を自動抽出
- 📊 **Excel自動転記**: テンプレートExcelファイルに自動で値を書き込み
- 🔍 **高精度OCR**: Azure Document Intelligence（prebuilt-invoice/read/documentモデル）を複数フォールバック使用
- 📅 **月自動検出**: ファイル名から対象月を自動判定（マイナス1ヶ月補正対応）
- 👁️ **PDFプレビュー**: アップロードしたPDFをその場で確認可能
- 📋 **OCR結果表示**: 信頼度80%以上の場合、OCR全文を折りたたみ表示（法人番号・住所のコピペに便利）
- 🎨 **直感的UI**: ベージュ・アイボリーを基調とした落ち着いたデザイン

## 使い方

### 1. 初期設定（初回のみ）

#### 環境変数の設定
`.env`ファイルをプロジェクトルートに作成し、Azure Document Intelligenceの認証情報を設定：

```env
AZURE_FORMREC_ENDPOINT=https://your-resource.cognitiveservices.azure.com/
AZURE_FORMREC_KEY=your-api-key-here
```

#### Excelテンプレートの配置
`template_output.xlsx`をプロジェクトルートに配置してください。以下のセルに自動転記されます：

- **B1**: クライアント名
- **B2**: 住所
- **B4**: 法人番号
- **B21～M21**: 1月～12月の使用電力量（kWh）

### 2. サーバー起動

```bash
python app.py
```

ブラウザで `http://localhost:8000` にアクセス

### 3. 操作手順

#### Step 1: 基本情報の入力
1. **クライアント名**を入力（必須）
2. **住所**と**法人番号**を入力（任意・後からOCR結果からコピペ可能）

#### Step 2: 読み取りモードの選択
- **単月モード**: 1つのPDFに1ヶ月分のデータが含まれている場合
- **一年間モード**: 1つのPDFに12ヶ月分のデータが含まれている場合

#### Step 3: PDFのアップロード
- ドラッグ&ドロップ または クリックしてファイル選択
- 複数ファイルのアップロード可能
- ファイル名から月を自動検出（例: `202412月分_東京電力.pdf` → 11月として認識）
  - ※ 請求書は翌月発行されるため、自動的に-1ヶ月補正されます

#### Step 4: PDFプレビュー（任意）
- ファイル名（下線付き）をクリックすると新しいタブでPDFを確認できます

#### Step 5: 読み取り開始
- 「読み取り開始」ボタンをクリック
- AI解析が完了すると、抽出結果が表示されます

#### Step 6: OCR結果の確認とコピペ
- 信頼度80%以上の場合、「📄 OCR結果全文を表示」をクリック
- OCR全文から**法人番号**と**住所**をコピーして、上部の入力欄に貼り付け

#### Step 7: Excelを保存
- 「Excelを保存」ボタンをクリックして完了

## 技術スタック

### バックエンド
- **Python 3.9+**
- **FastAPI**: 高速Webフレームワーク
- **Azure Document Intelligence**: OCRエンジン
- **openpyxl**: Excel操作
- **pdf2image**: PDF前処理
- **Pillow**: 画像処理

### フロントエンド
- **Vanilla JavaScript**: フレームワーク不使用のシンプル設計
- **HTML5/CSS3**: レスポンシブUI
- **Google Fonts**: M PLUS Rounded 1c

## プロジェクト構造

```
Hikari_OCR/
├── app.py                      # アプリケーションエントリーポイント
├── config.json                 # 設定ファイル
├── requirements.txt            # Python依存パッケージ
├── template_output.xlsx        # Excelテンプレート
├── .env                        # 環境変数（要作成）
├── app/
│   ├── config.py              # 設定読み込み
│   ├── domain/
│   │   └── invoice.py         # Invoiceデータモデル
│   ├── services/
│   │   ├── ocr_service.py     # OCR処理ロジック
│   │   └── excel_service.py   # Excel書き込みロジック
│   └── ui/
│       ├── pages/
│       │   └── estimate_page.py  # APIエンドポイント
│       ├── templates/
│       │   └── index.html     # メインページHTML
│       ├── scripts/
│       │   └── main.js        # フロントエンドロジック
│       └── styles/
│           └── style.css      # スタイルシート
└── README.md
```

## インストール

### 1. リポジトリのクローン
```bash
git clone https://github.com/hikari-ouga/Hikari-OCR.git
cd Hikari-OCR
```

### 2. 依存パッケージのインストール
```bash
pip install -r requirements.txt
```

### 3. 環境変数の設定
`.env`ファイルを作成し、Azure認証情報を設定（上記「初期設定」参照）

### 4. 起動
```bash
python app.py
```

## 動作環境

- **Python**: 3.9以上
- **OS**: Windows, macOS, Linux
- **ブラウザ**: Chrome, Edge, Firefox（最新版推奨）

## トラブルシューティング

### ポート8000が使用中の場合
```bash
# Windowsの場合
Get-Process python* | Stop-Process -Force

# macOS/Linuxの場合
lsof -ti:8000 | xargs kill -9
```

### Azure OCRエラーが発生する場合
- `.env`ファイルのエンドポイントとキーが正しいか確認
- Azureリソースのクォータと課金状態を確認
- ネットワーク接続を確認

### kWhが抽出できない場合
- PDFが画像のみで構成されていないか確認
- ファイル名から月が正しく検出されているか確認
- OCR結果全文を確認し、手動で値を確認

## ライセンス

このプロジェクトはプライベート使用のために作成されています。

## 作者

hikari-ouga

## 更新履歴

### v2.0.0 (2025-12-16)
- FastAPIベースのWebアプリに全面リニューアル
- PDFプレビュー機能追加
- OCR結果全文表示機能追加（折りたたみ式）
- 住所・法人番号入力欄追加
- 月自動検出の-1補正機能追加
- UI/UXを大幅改善（ベージュ・アイボリーテーマ）

### v1.0.0
- 初版リリース
- Streamlitベースのプロトタイプ
