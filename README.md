# 見積書・請求書作成システム (Google Apps Script + TypeScript)

## 📄 概要

Googleスプレッドシート上のボタン操作により、入力された顧客情報・明細データを元に「見積書」または「請求書」をPDFとして生成し、指定されたメールアドレスに送信する自動化処理をGoogle Apps Script（GAS）でTypeScriptを使用して実装したシステムです。

## 🚀 セットアップ手順

### 1. Google Apps Scriptプロジェクトの作成

1. [Google Apps Script](https://script.google.com/) にアクセス
2. 「新しいプロジェクト」をクリック
3. プロジェクト名を「見積書請求書システム」に変更

### 2. 開発からデプロイまでの完全なワークフロー

このプロジェクトでは、ローカルでTypeScriptを使用して開発し、claspを使用してGoogle Apps Scriptにデプロイします。

#### 2.1 環境セットアップ

```bash
# 1. リポジトリをクローン（または依存関係をインストール）
npm install

# 2. claspにログイン
npm run clasp:login
```

#### 2.2 新規プロジェクトの場合

```bash
# 新しいGoogle Apps Scriptプロジェクトを作成
npm run clasp:create

# TypeScriptをコンパイルしてデプロイ
npm run deploy
```

#### 2.3 既存プロジェクトの場合

```bash
# .clasp.jsonファイルを作成
cp .clasp.json.example .clasp.json

# .clasp.jsonのscriptIdを既存のプロジェクトIDに変更
# scriptIdはGoogle Apps ScriptのURLから取得: https://script.google.com/d/{SCRIPT_ID}/edit

# TypeScriptをコンパイルしてデプロイ
npm run deploy
```

#### 2.4 開発ワークフロー

```bash
# ファイル変更を監視（開発時）
npm run watch

# 変更をGoogle Apps Scriptにプッシュ
npm run clasp:push

# または、ビルドとプッシュを一度に実行
npm run deploy
```

### 3. TypeScript開発環境のセットアップ（手動デプロイの場合）

このプロジェクトはTypeScriptで開発されています。claspを使用しない手動デプロイの場合は以下の手順を実行してください：

```bash
# 依存関係のインストール
npm install

# TypeScriptコンパイル
npm run build

# 型チェック
npm run type-check
```

### 4. Google Apps Script CLIセットアップ（clasp）（推奨）

claspを使用することで、ローカルで開発したTypeScriptコードを直接Google Apps Scriptにデプロイできます：

#### 3.1 claspの初期設定

```bash
# claspにログイン
npm run clasp:login

# 新しいApps Scriptプロジェクトを作成
npm run clasp:create

# 既存のプロジェクトを使用する場合は.clasp.jsonを設定
cp .clasp.json.example .clasp.json
# .clasp.jsonのscriptIdを既存のプロジェクトIDに変更
```

#### 3.2 デプロイメント

```bash
# TypeScriptをコンパイルしてGoogle Apps Scriptにプッシュ
npm run clasp:push

# または、ビルドとプッシュを一度に実行
npm run deploy
```

### 5. スクリプトファイルの追加（手動の場合）

以下のファイルをGoogle Apps Scriptプロジェクトに追加してください：

**claspを使用する場合（推奨）：**
上記の手順3でセットアップ後、`npm run deploy`を実行するだけです。

**手動でファイルを追加する場合：**

**TypeScriptソースファイル（推奨）：**
- **src/Code.ts** - メイン処理
- **src/Config.ts** - 設定定数
- **src/Utils.ts** - ユーティリティ関数

**またはコンパイル済みJavaScriptファイル：**
- **dist/Code.js** - メイン処理
- **dist/Config.js** - 設定定数  
- **dist/Utils.js** - ユーティリティ関数

> **注意**: Google Apps Scriptエディタに直接TypeScriptファイルをアップロードする場合は、`.ts`ファイルの内容をコピー&ペーストし、ファイル名から`.ts`拡張子を削除してください（例: `Code.ts` → `Code`）。

### 6. Googleスプレッドシートの作成と連携

1. 新しいGoogleスプレッドシートを作成
2. Google Apps Scriptの「リソース」→「このスクリプトに関連付けられたスプレッドシート」で連携

### 7. 初期セットアップの実行

1. Google Apps Scriptエディタで `initialSetup` 関数を実行
2. 必要な権限を許可
3. 入力シートとテンプレートシートが自動作成されます

### 8. 送信ボタンの設定

1. 入力シートのB19セルを選択
2. 「挿入」→「図形描画」でボタンを作成
3. ボタンに「送信」と入力
4. ボタンを右クリック→「スクリプトを割り当て」→ `sendDocument` を入力

### 9. フォルダ構造の準備

スプレッドシートと同じフォルダに以下のフォルダを作成してください：
- **見積書** - 見積書PDFの保存先
- **請求書** - 請求書PDFの保存先
- **バックアップ** - 送信記録の保存先

## 📝 使用方法

### 1. 基本情報の入力

入力シートで以下の項目を入力してください：

| 項目 | セル | 必須 | 説明 |
|------|------|------|------|
| 書類種別 | B2 | ✅ | 「見積書」または「請求書」 |
| 発行日 | B3 | ✅ | 書類の発行日 |
| 書類番号 | B4 | ✅ | 3桁の数字（例：001, 123） |
| 宛先会社名 | B5 | ✅ | 送付先の会社名 |
| 担当者名 | B6 | - | 担当者名 |
| 住所 | B7 | - | 会社住所 |
| メールアドレス | B8 | ✅ | 送付先メールアドレス |
| 備考 | B9 | - | 追加事項 |

### 2. 商品明細の入力

A10:D14の範囲に商品情報を入力してください：

| 列 | 項目 | 説明 |
|----|------|------|
| A | 品目 | 商品・サービス名 |
| B | 数量 | 数量 |
| C | 単価 | 単価（円） |
| D | 小計 | 数量×単価 |

### 3. 合計金額の入力

| 項目 | セル | 説明 |
|------|------|------|
| 小計 | F15 | 商品明細の合計 |
| 消費税 | F16 | 消費税額 |
| 合計 | F17 | 税込み合計額 |

### 4. 送信実行

1. すべての必須項目を入力
2. 「送信」ボタンをクリック
3. 送信確認ダイアログで「はい」をクリック
4. メール送信確認ダイアログで選択：
   - 「はい」: PDFを作成・保存・メール送信
   - 「いいえ」: PDFを作成・保存のみ（メール送信なし）
5. 処理完了まで待機

📧 **メール送信オプション**: メールアドレスを入力した場合でも、送信するかどうかを選択できます。

## 🔧 設定のカスタマイズ

### 発行元情報の変更

`Config.gs`ファイルの以下の部分を編集してください：

```javascript
EMAIL: {
  SENDER_COMPANY: '株式会社サンプル',    // 会社名
  SENDER_DEPARTMENT: '営業部',          // 部署名
  SENDER_NAME: '山田太郎'              // 担当者名
}
```

### セル位置の変更

`CONFIG.CELLS`や`CONFIG.TEMPLATE_CELLS`で各項目のセル位置を変更できます。

### メールテンプレートの変更

`Code.gs`の`sendEmailWithPDF`関数内でメールの件名と本文を変更できます。

## 📋 機能一覧

- ✅ PDF生成（見積書・請求書）
- ✅ メール送信の選択式実行（メール送信する/しないを選択可能）
- ✅ PDFファイルの自動保存
- ✅ 書類番号による管理（3桁の数字で指定）
- ✅ ファイル名の統一形式（[書類種別]-[日付]-[書類番号]-[会社名].pdf）
- ✅ 送信履歴の記録（メール送信有無も記録）
- ✅ Google Docsでのバックアップ
- ✅ エラーハンドリング
- ✅ 入力データの検証
- ✅ TypeScriptでの型安全性

## 🛠️ 開発

### TypeScript開発について

このプロジェクトはTypeScriptで開発されており、以下の利点があります：

- **型安全性**: コンパイル時に型エラーを検出
- **IntelliSense**: IDEでの優れた補完機能
- **保守性**: 大規模なコードベースでの保守が容易
- **ドキュメント**: 型定義がドキュメントとして機能

### 開発からclasp pushまでの完全手順

```bash
# 1. 初回セットアップ
npm install                    # 依存関係のインストール
npm run clasp:login           # claspにログイン
npm run clasp:create          # 新規プロジェクト作成（または既存プロジェクトの.clasp.json設定）

# 2. TypeScript開発
npm run watch                 # ファイル変更監視開始（別ターミナル）
# src/フォルダでTypeScriptファイルを編集

# 3. ビルドとデプロイ
npm run build                 # TypeScriptコンパイル
npm run clasp:push           # Google Apps Scriptにプッシュ

# または、ワンコマンドでビルド+プッシュ
npm run deploy               # ビルドとプッシュを一度に実行

# 4. 型チェック（任意）
npm run type-check           # 型エラーの確認
```

### 利用可能なNPMスクリプト

```bash
# 依存関係のインストール
npm install

# TypeScriptコンパイル
npm run build

# ファイル監視モード（開発時）
npm run watch

# 型チェックのみ実行
npm run type-check

# claspコマンド
npm run clasp:login      # claspにログイン
npm run clasp:create     # 新しいApps Scriptプロジェクトを作成
npm run clasp:push       # ビルド後、Google Apps Scriptにプッシュ
npm run deploy           # ビルドとプッシュを一度に実行
```

### 開発からデプロイまでの完全なワークフロー

```bash
# 1. 初回セットアップ
npm install
npm run clasp:login
npm run clasp:create  # 新規プロジェクトの場合

# 2. 開発サイクル
npm run watch         # 別ターミナルで実行（ファイル変更を監視）

# 3. デプロイ
npm run deploy        # TypeScriptをコンパイルしてGoogle Apps Scriptにプッシュ
```

### 既存プロジェクトとの連携

既存のGoogle Apps Scriptプロジェクトがある場合：

```bash
# 1. .clasp.jsonファイルを作成
cp .clasp.json.example .clasp.json

# 2. scriptIdを編集
# .clasp.jsonファイルのscriptIdを既存のプロジェクトIDに変更

# 3. appscript.jsonをプロジェクトに追加（必要に応じて）
# 既存プロジェクトにappscript.jsonがない場合、自動的に追加されます

# 4. デプロイ
npm run deploy
```

### ファイル構成

```
src/
├── Config.ts    # 設定定数と型定義
├── Utils.ts     # ユーティリティ関数
└── Code.ts      # メイン処理

dist/            # コンパイル済みJavaScript
├── Config.js
├── Utils.js
└── Code.js
```

### 型定義

主要な型定義：
- `DocumentType`: 書類種別（'見積書' | '請求書'）
- `ItemData`: 商品明細の型
- `InputData`: 入力データ全体の型
- `AppConfig`: アプリケーション設定の型

## 🔍 トラブルシューティング

### claspに関する問題

1. **「User has not enabled the Apps Script API」エラー**
   - https://script.google.com/home/usersettings にアクセス
   - 「Google Apps Script API」を有効にしてください

2. **「Invalid authentication credentials」エラー**
   - `npm run clasp:login`を再実行してください
   - 認証情報をクリアしたい場合は`~/.clasprc.json`を削除してください

3. **「File not found in project」エラー**
   - `.clasp.json`の`scriptId`が正しいことを確認してください
   - プロジェクトURLの`/d/{SCRIPT_ID}/edit`部分がscriptIdです

4. **push後にファイルが見つからない**
   - `.clasp.json`の`rootDir`と`filePushOrder`を確認してください
   - `npm run build`でdistフォルダが作成されていることを確認してください

### よくある問題

1. **「権限が必要です」エラー**
   - Google Apps Scriptの実行権限を許可してください

2. **「メールアドレスが無効です」エラー**
   - メールアドレスの形式を確認してください

3. **「フォルダが見つかりません」エラー**
   - 必要なフォルダ（見積書、請求書、バックアップ）を作成してください

4. **PDF生成エラー**
   - テンプレートシートが存在することを確認してください
   - セル参照が正しいことを確認してください

### ログの確認

Google Apps Scriptエディタの「実行の記録」でエラーログを確認できます。

## 📞 サポート

問題が発生した場合は、以下の情報を含めてお問い合わせください：

- エラーメッセージ
- 入力データの内容
- 実行ログ（Google Apps Scriptの実行の記録）

## 📄 ライセンス

このプロジェクトはMITライセンスの下で公開されています。

## 🔄 バージョン履歴

- **v1.0.0** - 初回リリース
  - 基本的な見積書・請求書作成機能
  - PDF生成・メール送信機能
  - 送信履歴・バックアップ機能