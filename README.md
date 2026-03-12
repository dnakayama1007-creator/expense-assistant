# 経費精算アシスタント

Discordで経費メモや領収書を送るだけで自動記録・集計できるWebアプリ＋Discord Botシステムです。freeeとの連携や、Google Cloud Vision APIによるOCR（領収書の自動読み取り）にも対応しています。

---

## 主な機能

- 📱 **Discord Bot** — チャンネルにテキスト or 領収書画像を送信するだけで自動登録
- 🖥️ **Webダッシュボード** — ブラウザで経費一覧・編集・集計を確認
- 📷 **OCR対応** — 領収書画像から金額・日付を自動読み取り（Google Cloud Vision）
- 🔗 **freee連携** — 登録した経費を freee に自動で取り込み
- 🌐 **外部公開** — Cloudflaredトンネルでスマホからもアクセス可能

---

## セットアップ

### 1. 必要なもの

- Node.js（v18以上推奨）
- Discord Developer Portal でBotアカウントを作成
- （任意）Google Cloud Platform サービスアカウント（OCR用）
- （任意）freee APIアプリ登録（freee連携用）

### 2. インストール

```bash
git clone https://github.com/dnakayama1007-creator/expense-assistant.git
cd expense-assistant
npm install
```

### 3. 環境変数の設定

`.env.example` をコピーして `.env` を作成し、必要な値を入力してください。

```bash
cp .env.example .env
```

### 4. 起動

```bash
# Webサーバー起動（http://localhost:3000）
npm start

# Discord Bot起動
npm run bot

# 両方同時起動
npm run dev
```

---

## Discord Bot の使い方

### 経費を登録する（テキスト）

改行区切りで以下の形式で送信：

```
購入店舗
内容
単価@数量
日付（省略可）
```

**例：**
```
Yahoo(ジョーシン)
HDD 4TB
16800@2
3/1
```
→ ¥33,600 として登録されます

※数量が1の場合は `@ 数量` を省略できます
※日付を省略すると当日が使用されます

### 領収書を登録する（画像）

領収書の写真をそのまま送信してください。テキストと一緒に送ることもできます。

### コマンド

| コマンド | 説明 |
|---------|------|
| `!ヘルプ` | 使い方を表示 |
| `!一覧` | 最近の経費10件を表示 |
| `!合計` | 今月の合計・カテゴリ別内訳を表示 |

---

## freee 連携

1. ブラウザで `http://localhost:3000` を開く
2. 設定画面から freee Client ID / Secret を入力
3. 「freee認証」ボタンをクリックしてOAuth認証を完了
4. 経費一覧からfreeeに登録したい項目を選択して登録

---

## 機密ファイルについて

以下のファイルはセキュリティ上の理由からGitで管理していません。
サーバー上に直接配置してください。

| ファイル | 内容 |
|---------|------|
| `.env` | Discord Token 等の環境変数 |
| `freee_config.json` | freee OAuthトークン（自動生成） |
| `google_service_account.json` | Google Cloud サービスアカウント秘密鍵 |
| `discord_expenses.json` | 経費データ（自動生成） |
| `processed_msg_ids.json` | 重複防止用IDリスト（自動生成） |
| `receipts/` | 領収書画像ディレクトリ（自動生成） |

---

## ライセンス

MIT License