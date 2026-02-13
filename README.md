# one building 技術レポート生成ツール

省エネ診断レポートPowerPoint自動生成Webアプリケーション

## 📋 概要

このアプリケーションは、省エネ計算結果のPDFまたはテキストファイルから、
エネルギー診断レポートのPowerPointファイルを自動生成します。

### 主な機能

- **ファイルアップロード**: PDF、テキスト（.txt, .md）形式に対応
- **自動データ抽出**: 建物情報、BEI値、エネルギー消費量などを自動抽出
- **計算方法自動判定**: 標準入力法とモデル建物法を自動判定
- **グラフ自動生成**: BEI比較、積み上げ棒グラフ、パイチャートを生成
- **PowerPoint出力**: one buildingデザインガイドライン準拠のレポートを生成

### 対応する計算方法

- **標準入力法**: 詳細なエネルギー消費量分析（7スライド）
- **モデル建物法**: BEIm/BPIm表記での簡易分析（5スライド）

## 🚀 セットアップ

### 必要な環境

- Python 3.8以上
- pip3

### インストール

```bash
# リポジトリをクローン
cd /path/to/project

# 依存パッケージをインストール
pip3 install -r requirements.txt

# OS レベルのパッケージ（Ubuntu/Debian）
sudo apt-get update
sudo apt-get install -y fonts-noto-cjk fonts-ipafont-gothic
```

## 💻 使い方

### ローカルで起動

```bash
streamlit run app.py
```

ブラウザで http://localhost:8501 にアクセスします。

### Streamlit Cloudにデプロイ

1. GitHubリポジトリにコードをプッシュ
2. [Streamlit Cloud](https://streamlit.io/cloud) にアクセス
3. リポジトリを接続してデプロイ

### アプリの使用方法

1. **ファイルアップロード**: 省エネ計算結果のファイルを選択
2. **レポート生成**: 「レポート生成」ボタンをクリック
3. **ダウンロード**: 生成されたPowerPointファイルをダウンロード

## 📁 ファイル構成

```
streamlit_app/
├── app.py                      # Streamlitメインアプリ
├── report_generator.py         # データ抽出・グラフ生成モジュール
├── slides.py                   # スライド作成モジュール
├── requirements.txt            # Python依存パッケージ
├── packages.txt                # OSレベルの依存パッケージ
├── NotoSansJP-Regular.ttf      # 日本語フォント
├── README.md                   # このファイル
└── test_output.pptx            # テスト出力サンプル
```

## 🎨 デザイン仕様

- **スライド形式**: 16:9 ワイドスクリーン
- **フォント**: Noto Sans JP
- **カラースキーム**: 
  - メインカラー: #397577
  - アクセントカラー: #D97D54
  - 警告色: #DC3545
  - 適合色: #28A745

## 🧪 テスト

```bash
# データ抽出テスト
python3 -c "
import io
from report_generator import extract_data_from_file
with open('test_sample.txt', 'rb') as f:
    file_obj = io.BytesIO(f.read())
data = extract_data_from_file(file_obj, 'test_sample.txt')
print(f'建物名: {data[\"building_name\"]}')
print(f'BEI: {data[\"bei_total\"]}')
"

# 統合テスト（PowerPoint生成）
python3 test_integration.py
```

## 📝 技術スタック

- **Webフレームワーク**: Streamlit 1.31.0
- **PowerPoint生成**: python-pptx 0.6.23
- **グラフ生成**: matplotlib 3.8.2
- **PDF処理**: pdfplumber 0.10.4
- **画像処理**: Pillow 10.2.0

## 🔧 トラブルシューティング

### フォントが表示されない

```bash
# フォントキャッシュをクリア
rm -rf ~/.cache/matplotlib
```

### PDFが読み込めない

pdfplumberが正しくインストールされているか確認:

```bash
pip3 install --upgrade pdfplumber
```

## 📄 ライセンス

© 2026 one building | BIM sustaina for Energy

## 🤝 サポート

技術的な質問やバグ報告は、プロジェクトの担当者にお問い合わせください。
