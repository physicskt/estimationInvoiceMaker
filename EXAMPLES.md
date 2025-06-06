# 📋 使用例・デモ

このファイルでは、見積書・請求書作成システムの実際の使用例を示します。

## 🎯 デモシナリオ：Webサイト制作の見積書作成

### 入力例

#### 基本情報
- **書類種別**: 見積書
- **発行日**: 2024年06月15日
- **宛先会社名**: 株式会社サンプルコーポレーション
- **担当者名**: 田中太郎様
- **住所**: 〒150-0001 東京都渋谷区神宮前1-2-3 サンプルビル5F
- **メールアドレス**: tanaka@sample-corp.com
- **備考**: 何かご不明な点がございましたら、お気軽にお声がけください。

#### 商品明細

| 品目 | 数量 | 単価 | 小計 |
|------|------|------|------|
| コーポレートサイト制作 | 1 | 800,000 | 800,000 |
| レスポンシブ対応 | 1 | 200,000 | 200,000 |
| CMS導入 | 1 | 300,000 | 300,000 |
| SEO基本設定 | 1 | 100,000 | 100,000 |
| 保守サポート（6ヶ月） | 6 | 50,000 | 300,000 |

#### 合計金額
- **小計**: ¥1,700,000
- **消費税**: ¥170,000
- **合計**: ¥1,870,000

### 送信されるメール例

```
件名: 【見積書】株式会社サンプルコーポレーション様宛

本文:
株式会社サンプルコーポレーション 御中

平素よりお世話になっております。
以下の通り、見積書をお送りします。

ご確認のほど、よろしくお願いいたします。

────────────────────────

株式会社サンプル
営業部 山田太郎

添付ファイル: 見積書_株式会社サンプルコーポレーション_20240615.pdf
```

### 生成されるPDFの内容

```
                                見積書

                                           発行日：2024年06月15日

株式会社サンプルコーポレーション                    株式会社サンプル
田中太郎様                                        営業部 山田太郎
〒150-0001 東京都渋谷区神宮前1-2-3                〒000-0000 東京都●●区●●
サンプルビル5F                                   TEL: 03-0000-0000
                                               EMAIL: sample@example.com

─────────────────────────────────────────────

| 品目                    | 数量 | 単価      | 小計      |
|------------------------|------|-----------|-----------|
| コーポレートサイト制作      | 1    | ¥800,000  | ¥800,000  |
| レスポンシブ対応          | 1    | ¥200,000  | ¥200,000  |
| CMS導入                  | 1    | ¥300,000  | ¥300,000  |
| SEO基本設定              | 1    | ¥100,000  | ¥100,000  |
| 保守サポート（6ヶ月）      | 6    | ¥50,000   | ¥300,000  |

                                               小計    ¥1,700,000
                                               消費税    ¥170,000
                                               合計    ¥1,870,000

備考：
何かご不明な点がございましたら、お気軽にお声がけください。
```

## 🚀 ステップバイステップ操作例

### Step 1: データ入力
1. Googleスプレッドシートの「入力」シートを開く
2. B2セルで「見積書」を選択
3. B3セルに「2024/06/15」を入力
4. B4セルに「株式会社サンプルコーポレーション」を入力
5. B5セルに「田中太郎様」を入力
6. B6セルに住所を入力
7. B7セルに「tanaka@sample-corp.com」を入力
8. B8セルに備考を入力

### Step 2: 商品明細入力
1. A10セルに「コーポレートサイト制作」を入力
2. B10セルに「1」を入力
3. C10セルに「800000」を入力
4. 同様に他の商品も入力

### Step 3: 金額計算
1. 「計算」ボタンをクリック（calculateTotals関数）
2. 自動的に小計、消費税、合計が計算される

### Step 4: 送信実行
1. 「送信」ボタンをクリック（sendDocument関数）
2. 確認ダイアログで「はい」をクリック
3. 処理完了まで待機

### Step 5: 結果確認
1. 完了メッセージで送信結果を確認
2. 「送信履歴」シートで履歴を確認
3. Google Driveの指定フォルダでPDFを確認

## 📊 システムによる自動処理

### 1. PDFファイル生成
- ファイル名: `見積書_株式会社サンプルコーポレーション_20240615.pdf`
- 保存先: Google Drive > 見積書フォルダ

### 2. メール送信
- 宛先: tanaka@sample-corp.com
- 件名: 【見積書】株式会社サンプルコーポレーション様宛
- PDFファイルが添付される

### 3. 送信履歴記録
「送信履歴」シートに以下が記録される：
- 送信日時: 2024/06/15 14:30:15
- 書類種別: 見積書
- 宛先会社: 株式会社サンプルコーポレーション
- メールアドレス: tanaka@sample-corp.com
- ファイル名: 見積書_株式会社サンプルコーポレーション_20240615.pdf
- ファイルURL: https://drive.google.com/file/d/xxxxx

### 4. バックアップ作成
Google Docsに以下の内容でバックアップドキュメントが作成される：
- ドキュメント名: `見積書_バックアップ_株式会社サンプルコーポレーション_20240615_143015`
- 内容: 送信記録の詳細情報

## 🔄 様々な使用パターン

### パターン1: 請求書の作成
基本的な流れは同じですが、書類種別を「請求書」に変更します。
- PDFは「請求書」フォルダに保存される
- メール件名が「【請求書】〜」になる

### パターン2: 定期的なサービス請求書
```
品目: 月額保守サービス
数量: 1
単価: 50,000
小計: 50,000
```

### パターン3: 商品販売の請求書
```
品目: Webサイト制作ソフトウェア ライセンス
数量: 5
単価: 100,000
小計: 500,000
```

## ⚠️ 注意事項とベストプラクティス

### 入力時の注意点
1. **メールアドレス**: 正確な形式で入力する
2. **金額**: 数値のみ入力（カンマや円マークは不要）
3. **日付**: 正しい日付形式で入力
4. **会社名**: 正式な会社名を入力

### 効率的な使用方法
1. **テンプレート作成**: よく使用する顧客情報は別シートに保存
2. **計算確認**: 送信前に「計算」ボタンで金額を確認
3. **履歴活用**: 送信履歴シートで過去の取引を参照
4. **バックアップ**: 定期的にスプレッドシート全体をバックアップ

### トラブル回避
1. **事前確認**: 送信前に宛先メールアドレスを再確認
2. **テスト送信**: 新規設定時は自分のメールアドレスでテスト
3. **権限確認**: Google Apps Scriptの実行権限を定期的に確認