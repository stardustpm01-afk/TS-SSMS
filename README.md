# スキルシート変換ツール

各社バラバラのExcel/PDFスキルシートを自社フォーマット（SAP職務経歴書）に自動変換するツールです。

## セットアップ手順

### 1. 仮想環境の作成（TS-ONEと分離）
```bash
python -m venv venv
venv\Scripts\activate      # Windows
source venv/bin/activate   # Mac/Linux
```

### 2. ライブラリのインストール
```bash
pip install -r requirements.txt
```

### 3. APIキーの設定
`.env` ファイルを開いて、Gemini APIキーを貼り付けてください：
```
GEMINI_API_KEY=あなたのAPIキー
```

### 4. アプリ起動
```bash
streamlit run app.py
```

ブラウザが自動で開きます（http://localhost:8501）

## 使い方
1. サイドバーにGemini APIキーを入力
2. Excel/PDFのスキルシートをアップロード（複数可）
3. 「AIで自動変換する」ボタンをクリック
4. 結果を確認・PDFダウンロード

## 注意事項
- APIキーは `.env` ファイルにのみ保存し、他人に共有しないでください
- `.env` ファイルはGitにコミットしないでください
