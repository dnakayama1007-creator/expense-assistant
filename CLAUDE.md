# CLAUDE.md

## プロジェクト概要

このファイルはAIアシスタント（Claude）がプロジェクトを理解するための設定ファイルです。
新しいセッションを開始する際に、このファイルを参照してください。

## 技術スタック

- **言語**: Kotlin / Java
- **IDE**: Antigravity（ローカル開発）
- **バージョン管理**: Git / GitHub

## プロジェクト構成

```
project-root/
├── CLAUDE.md                  # このファイル（AI向け設定）
├── README.md                  # プロジェクト説明（人間向け）
├── .gitignore                 # Git除外設定
├── src/
│   ├── main/
│   │   ├── kotlin/            # Kotlinソースコード
│   │   └── resources/         # 設定ファイル・リソース
│   └── test/
│       └── kotlin/            # テストコード
├── build.gradle.kts           # ビルド設定（Gradle）
└── docs/                      # ドキュメント類
```

## コーディング規約

- Kotlinの公式コーディング規約に従う
- 関数名・変数名はキャメルケース（例: `getUserData`）
- クラス名はパスカルケース（例: `UserService`）
- コメントは日本語でOK

## レビュー重点項目

- 🐛 バグ・エラーの発見（NullPointerException、例外処理の漏れ）
- ⚡ パフォーマンス改善（不要なループ、メモリ効率）
- 🔒 セキュリティ（入力検証、認証・認可）

## よく使うコマンド

```bash
# ビルド
./gradlew build

# テスト実行
./gradlew test

# 実行
./gradlew run
```

## Git運用ルール

- `main` ブランチは常に動作する状態を保つ
- 機能追加は `feature/機能名` ブランチで作業
- コミットメッセージは日本語でわかりやすく書く
  - 例: `feat: ユーザー認証機能を追加`
  - 例: `fix: ログイン時のNullPointerExceptionを修正`

## 注意事項

- APIキーや秘密情報は `.env` ファイルに保存し、絶対にGitにコミットしない
- `.gitignore` に機密ファイルが含まれていることを確認すること
