# 工数管理システム

Google Apps Scriptを使用したスプレッドシートベースの出退勤・工数管理システムです。

## 機能

- **出退勤管理**: ワンクリックで出勤・退勤を記録
- **工数記録**: 案件別の作業時間を記録
- **レポート機能**: 期間指定での工数・出退勤レポート生成
- **リアルタイム更新**: スプレッドシートと連動したデータ管理

## セットアップ

### 1. 認証設定

```bash
# Google Apps Script APIを有効化後、認証
npm run login
```

### 2. プロジェクト作成とデプロイ

```bash
# 新しいGASプロジェクトを作成してデプロイ
npm run setup
```

### 3. 手動セットアップ（既存プロジェクトを使用する場合）

```bash
# プロジェクトをビルド
npm run build

# .clasprc.jsonをコピーして設定
cp .clasprc.example.json .clasprc.json
# .clasprc.jsonのscriptIdを編集

# プロジェクトをプッシュ
npm run push

# デプロイ
npm run deploy
```

## 開発コマンド

```bash
# ビルド
npm run build

# ファイル変更を監視してビルド
npm run watch

# GASにプッシュ
npm run push

# 開発用（ビルド→プッシュ→ブラウザで開く）
npm run dev

# ログ確認
npm run logs

# プロジェクトをブラウザで開く
npm run open
```

## プロジェクト構造

```
├── src/
│   └── main.ts          # メインのGASコード
├── html/
│   ├── index.html       # メインHTML
│   ├── styles.html      # CSS
│   └── scripts.html     # JavaScript
├── types/
│   └── gas.d.ts         # GAS型定義
├── dist/                # ビルド出力（claspがこのディレクトリを監視）
├── tsconfig.json        # TypeScript設定
├── package.json         # Node.js設定
└── .clasprc.json        # clasp設定（要作成）
```

## スプレッドシート構造

システムは自動的に以下のシートを作成します：

### 工数記録シート
- 記録日時、ユーザー、作業日、案件名、工数（時間）、説明

### 出退勤記録シート
- 日付、ユーザー、出勤時刻、退勤時刻

### 案件マスタシート
- 案件ID、案件名、開始日、終了日

## 使用方法

1. **出退勤記録**: 出勤・退勤ボタンをクリック
2. **工数入力**: 作業日、案件名、工数を入力して記録
3. **レポート確認**: 期間を指定してレポートを生成

## 技術仕様

- **フロントエンド**: HTML/CSS/JavaScript
- **バックエンド**: Google Apps Script（TypeScript）
- **データベース**: Google Spreadsheet
- **認証**: Google認証（自動）
- **デプロイ**: clasp CLI

## ライセンス

MIT License