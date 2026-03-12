# CLAUDE.md

このファイルはAIアシスタント（Claude / Antigravity）がプロジェクトを即座に理解するための設定ファイルです。
新しいセッションを開始する際に、必ずこのファイルを最初に参照してください。

---

## プロジェクト概要

**経費精算アシスタント** — Discordで家族・チームの経費メモや領収書を送ると自動で記録・集計し、freeeへの連携もできるWebアプリ＋Discord Botシステム。

### 主な機能
- Discordからテキスト or 画像（領収書）を送信→自動で経費登録
- Webブラウザ（localhost:3000）でダッシュボード表示・編集
- Google Cloud Vision APIによるOCR（領収書テキスト読み取り）
- freee APIと連携して経費データを自動登録
- Cloudflaredトンネルで外部からもアクセス可能

---

## 技術スタック

- **言語**: Node.js（CommonJS）/ JavaScript
- **フロントエンド**: Vanilla HTML + CSS + JavaScript（`index.html`, `index.css`）
- **バックエンド**: Node.js HTTPサーバー（`server.js`）
- **Discord Bot**: discord.js v14（`discord-bot.js`）
- **外部API**:
  - freee API（経費登録・OAuth認証）
  - Google Cloud Vision API（OCR）
  - Google Sheets API（スプレッドシート読み込み）
- **トンネル**: Cloudflared（`start-tunnel.js`）
- **IDE**: Antigravity（ローカル開発）
- **バージョン管理**: Git / GitHub

---

## ファイル構成

```
expense-assistant/
├── CLAUDE.md                    # このファイル（AI向け設定）★必読
├── README.md                    # プロジェクト説明（人間向け）
├── .gitignore                   # Git除外設定
├── .env                         # 環境変数（Gitにコミットしない！）
├── .env.example                 # 環境変数サンプル（コミットOK）
├── package.json                 # Node.js依存関係
├── package-lock.json            # ロックファイル
│
├── server.js                    # ★メインサーバー（HTTPサーバー + APIエンドポイント）
├── discord-bot.js               # ★Discord Bot本体
├── app.js                       # サブアプリ（フロントエンドロジック補助）
├── start-tunnel.js              # Cloudflaredトンネル起動スクリプト
│
├── index.html                   # フロントエンドUI（ダッシュボード）
├── index.css                    # スタイルシート
│
├── docs/                        # ドキュメント類
│   └── LINE_BOT_SETUP.md        # LINE Bot設定手順
│
├── receipts/                    # 領収書画像（Gitにコミットしない！）
│
# 以下はローカルのみ（.gitignoreで除外済み）
├── .env                         # 環境変数
├── freee_config.json            # freee OAuthトークン（機密情報）
├── google_service_account.json  # GoogleサービスアカウントJSON（機密情報）
├── discord_expenses.json        # 経費データ（JSONファイルDB）
└── processed_msg_ids.json       # 重複防止用処理済みメッセージIDブラックリスト
```

---

## 起動方法

```bash
# 依存関係インストール
npm install

# Webサーバーのみ起動（http://localhost:3000）
npm start

# Discord Botのみ起動
npm run bot

# 両方同時起動
npm run dev

# Cloudflaredトンネル起動（外部公開）
node start-tunnel.js
```

---

## 主要APIエンドポイント（server.js）

| メソッド | パス | 説明 |
|---------|------|------|
| GET | `/api/expenses` | 全経費データ取得 |
| POST | `/api/expenses` | 経費データ保存（削除済みをブラックリスト登録） |
| GET | `/api/discord-expenses` | Discord由来の経費取得（ブラックリスト除外済み） |
| GET/POST | `/api/processed-msg-ids` | 処理済みDiscordメッセージIDの管理 |
| POST | `/api/ocr` | Google Vision APIでOCR実行 |
| GET | `/api/freee/status` | freee認証状態確認 |
| GET | `/api/freee/auth` | freee OAuth認証開始 |
| GET | `/api/freee/callback` | freee OAuthコールバック |
| POST | `/api/freee/config` | freee設定保存（client_id, secret, company_id） |
| GET | `/api/freee/master` | freeeマスタデータ取得（勘定科目・部門等） |
| GET | `/api/sheets/read` | Googleスプレッドシート読み込み |

---

## Discord Bot 仕様（discord-bot.js）

### 登録メッセージ形式（改行区切り）
```
購入店舗
内容
単価@数量   ← 数量省略可（@なしで単価のみ）
日付        ← 省略可（省略時は当日）
```

### コマンド
| コマンド | 説明 |
|---------|------|
| `!ヘルプ` / `!help` | ヘルプ表示 |
| `!一覧` / `!list` | 最近の経費10件表示 |
| `!合計` / `!total` | 今月の経費合計表示 |

### 購入者マッピング（discord-bot.js `getBuyer()`）
Discord表示名 → 購入者名のマッピングが定義されています。
新しいメンバーを追加する場合は `getBuyer()` 関数を修正してください。

### オフライン時メッセージの取り込み
Bot起動時に `fetchMissedMessages()` が呼ばれ、オフライン中のメッセージを遡って処理します。
ただし「ボットの最終応答時刻より前のメッセージ」はスキップ（重複防止）。

---

## データ管理

### 経費データ（discord_expenses.json）
全経費はローカルJSONファイルで管理。スキーマ：

```json
{
  "id": "discord_メッセージID_ランダム文字列",
  "discordMsgId": "DiscordメッセージID",
  "date": "YYYY-MM-DD",
  "orderDate": "YYYY-MM-DD",
  "unitPrice": 1000,
  "quantity": 2,
  "amount": 2000,
  "category": "消耗品費",
  "payment": "購入店舗名",
  "description": "内容説明",
  "receipt": "/receipts/receipt_xxx.jpeg",
  "status": "未清算",
  "source": "discord",
  "buyer": "購入者名",
  "createdAt": "ISO8601形式"
}
```

### 削除防止ロジック
- Webから経費を削除 → そのDiscordメッセージIDが `processed_msg_ids.json` に追記（ブラックリスト）
- Bot再起動時のオフラインメッセージ取り込みでも、ブラックリストのIDはスキップ

---

## 環境変数（.env）

`.env.example` を参照してください。必要な変数：

```
DISCORD_TOKEN=          # Discord BotトークN
CLOUDFLARE_TUNNEL_TOKEN= # Cloudflared トンネルトークン（任意）
```

freee・Google系の設定はWebUI経由でJSONファイルに保存されます。

---

## セキュリティ注意事項

- `freee_config.json` はOAuthトークンを含むため**絶対にGitにコミットしない**
- `google_service_account.json` はGCPサービスの秘密鍵を含むため**絶対にGitにコミットしない**
- `discord_expenses.json` は個人の経費情報を含むため**Gitにコミットしない**
- `.env` は**絶対にGitにコミットしない**（`.env.example` のみコミットOK）
- これらはすべて `.gitignore` で除外済み

---

## よくある問題

| 問題 | 対処方法 |
|------|---------|
| `git` コマンドが使えない | `$env:PATH += ";C:\Program Files\Git\bin"` を実行してパスを通す |
| Bot起動しない | `.env` に `DISCORD_TOKEN` が設定されているか確認 |
| freee連携できない | `/api/freee/auth` にブラウザでアクセスし再認証 |
| OCRが動かない | `google_service_account.json` が存在するか確認 |
| 外部からアクセスできない | `node start-tunnel.js` でCloudflaredトンネルを起動 |

---

## Git運用ルール

- `main` ブランチは常に動作する状態を保つ
- コミットメッセージは日本語でわかりやすく書く
  - `feat: Discord Botに新コマンドを追加`
  - `fix: freee連携時のトークン更新バグを修正`
  - `docs: CLAUDE.mdを更新`
