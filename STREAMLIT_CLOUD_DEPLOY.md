# Streamlit Cloudへのデプロイ手順

このドキュメントでは、one building技術レポート生成ツールをStreamlit Cloudに永久的にデプロイする手順を説明します。

---

## 前提条件

1. **GitHubアカウント**: [github.com](https://github.com)でアカウントを作成
2. **Streamlit Cloudアカウント**: [share.streamlit.io](https://share.streamlit.io)でGitHubアカウントと連携

---

## ステップ1: GitHubリポジトリを作成

### 方法A: GitHub CLIを使用（推奨）

```bash
cd /home/ubuntu/streamlit_app
gh auth login
gh repo create one-building-energy-report --public --source=. --remote=origin --push
```

### 方法B: GitHub Webインターフェースを使用

1. [github.com/new](https://github.com/new)にアクセス
2. リポジトリ名: `one-building-energy-report`
3. 公開設定: **Public**
4. 「Create repository」をクリック
5. ローカルからプッシュ:

```bash
cd /home/ubuntu/streamlit_app
git remote add origin https://github.com/<YOUR_USERNAME>/one-building-energy-report.git
git branch -M main
git push -u origin main
```

---

## ステップ2: Streamlit Cloudにデプロイ

1. **Streamlit Cloudにログイン**: [share.streamlit.io](https://share.streamlit.io)
2. **「New app」をクリック**
3. **リポジトリ情報を入力**:
   - Repository: `<YOUR_USERNAME>/one-building-energy-report`
   - Branch: `main`
   - Main file path: `app.py`
4. **「Deploy!」をクリック**

デプロイには数分かかります。完了すると、永久的なURLが発行されます:

```
https://<app-name>.streamlit.app
```

---

## ステップ3: カスタムドメインの設定（オプション）

Streamlit Cloudの設定画面から、独自ドメインを設定できます。

1. アプリの設定画面を開く
2. 「Settings」→「General」
3. 「Custom subdomain」に希望のサブドメインを入力
4. 保存

例: `https://energy-report.streamlit.app`

---

## 必要なファイル

以下のファイルがリポジトリに含まれていることを確認してください:

- ✅ `app.py` - Streamlitメインアプリ
- ✅ `report_generator.py` - データ抽出モジュール
- ✅ `slides.py` - PowerPoint生成モジュール
- ✅ `requirements.txt` - Python依存パッケージ
- ✅ `packages.txt` - システム依存パッケージ
- ✅ `NotoSansJP-Regular.ttf` - 日本語フォント
- ✅ `README.md` - プロジェクト説明
- ✅ `.gitignore` - Git除外設定

---

## トラブルシューティング

### エラー: "ModuleNotFoundError"

`requirements.txt`に必要なパッケージが記載されているか確認してください。

### エラー: "Font not found"

`NotoSansJP-Regular.ttf`がリポジトリに含まれているか確認してください。

### エラー: "poppler-utils not found"

`packages.txt`に以下が記載されているか確認してください:

```
poppler-utils
```

---

## 更新方法

コードを更新した場合、以下のコマンドでデプロイを更新できます:

```bash
cd /home/ubuntu/streamlit_app
git add .
git commit -m "Update: <変更内容>"
git push origin main
```

Streamlit Cloudが自動的に変更を検知し、再デプロイします。

---

## サポート

- Streamlit Cloud公式ドキュメント: https://docs.streamlit.io/streamlit-community-cloud
- Streamlit Community Forum: https://discuss.streamlit.io

---

**© 2026 one building | BIM sustaina for Energy**
