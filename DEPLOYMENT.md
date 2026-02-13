# デプロイガイド

## 🚀 Streamlit Cloudへのデプロイ

### 前提条件

- GitHubアカウント
- Streamlit Cloudアカウント（無料）

### 手順

#### 1. GitHubリポジトリの作成

```bash
# 新しいリポジトリを作成
cd /path/to/streamlit_app
git init
git add .
git commit -m "Initial commit: one building energy report generator"

# GitHubにプッシュ
git remote add origin https://github.com/your-username/energy-report-app.git
git branch -M main
git push -u origin main
```

#### 2. Streamlit Cloudでデプロイ

1. [Streamlit Cloud](https://streamlit.io/cloud) にアクセス
2. 「New app」をクリック
3. GitHubリポジトリを選択
4. 以下を設定:
   - **Main file path**: `app.py`
   - **Python version**: 3.11
5. 「Deploy」をクリック

#### 3. 環境変数の設定（必要に応じて）

Streamlit Cloudの「Advanced settings」で環境変数を設定できます。

### デプロイ後の確認

- アプリのURLが自動生成されます（例: `https://your-app.streamlit.app`）
- ログを確認してエラーがないことを確認

---

## 🐳 Dockerでのデプロイ

### Dockerfileの作成

```dockerfile
FROM python:3.11-slim

WORKDIR /app

# システムパッケージのインストール
RUN apt-get update && apt-get install -y \
    fonts-noto-cjk \
    fonts-ipafont-gothic \
    && rm -rf /var/lib/apt/lists/*

# Pythonパッケージのインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# アプリケーションファイルのコピー
COPY . .

# ポート公開
EXPOSE 8501

# Streamlitアプリの起動
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

### ビルドと実行

```bash
# イメージのビルド
docker build -t energy-report-app .

# コンテナの実行
docker run -p 8501:8501 energy-report-app
```

ブラウザで http://localhost:8501 にアクセス

---

## ☁️ AWSへのデプロイ

### AWS Elastic Beanstalkを使用

#### 1. 必要なファイルの準備

`.ebextensions/python.config`:

```yaml
option_settings:
  aws:elasticbeanstalk:container:python:
    WSGIPath: app:application
  aws:elasticbeanstalk:application:environment:
    STREAMLIT_SERVER_PORT: 8501
```

#### 2. デプロイ

```bash
# EB CLIのインストール
pip install awsebcli

# 初期化
eb init -p python-3.11 energy-report-app

# 環境の作成とデプロイ
eb create energy-report-env
eb deploy
```

---

## 🔧 トラブルシューティング

### フォントが表示されない

**原因**: システムに日本語フォントがインストールされていない

**解決策**:
```bash
# Ubuntu/Debian
sudo apt-get install fonts-noto-cjk fonts-ipafont-gothic

# Alpine Linux (Docker)
apk add font-noto-cjk
```

### メモリ不足エラー

**原因**: グラフ生成やPDF処理でメモリを大量に使用

**解決策**:
- Streamlit Cloudの場合: 有料プランにアップグレード
- Dockerの場合: メモリ制限を増やす
  ```bash
  docker run -m 2g -p 8501:8501 energy-report-app
  ```

### PDFの読み込みエラー

**原因**: pdfplumberの依存関係が不足

**解決策**:
```bash
pip install --upgrade pdfplumber pillow
```

---

## 📊 パフォーマンス最適化

### キャッシュの活用

Streamlitの`@st.cache_data`デコレータを使用してデータ抽出をキャッシュ:

```python
@st.cache_data
def extract_data_from_file(file_obj, file_name):
    # ... データ抽出処理
    return data
```

### 画像圧縮

グラフのDPIを調整してファイルサイズを削減:

```python
plt.savefig(output_path, dpi=150)  # 300から150に変更
```

---

## 🔐 セキュリティ

### ファイルアップロードの制限

`app.py`で最大ファイルサイズを設定:

```python
st.set_page_config(
    page_title="...",
    max_upload_size=10  # 10MB
)
```

### 入力検証

アップロードされたファイルの検証を強化:

```python
if uploaded_file.size > 10 * 1024 * 1024:  # 10MB
    st.error("ファイルサイズが大きすぎます")
    return
```

---

## 📝 ログとモニタリング

### Streamlit Cloudのログ

Streamlit Cloudのダッシュボードでアプリのログを確認できます。

### カスタムログの追加

```python
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

logger.info(f"File uploaded: {uploaded_file.name}")
```

---

## 🤝 サポート

デプロイに関する質問は、プロジェクトの担当者にお問い合わせください。
