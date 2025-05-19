# Notion週次タスクレポート自動化 (GAS)

NotionのタスクデータベースからSlackへ週次レポートを自動送信し、同時にスプレッドシートへ履歴を記録するGoogle Apps Scriptプロジェクトです。

## 🌟 機能概要

- Notionタスクの状態（未着手・進行中）を自動抽出
- 期限順にソートし、担当者情報付きでSlackに通知
- 優先度の高いタスクを強調表示
- スプレッドシートに履歴として記録
- エラー発生時にSlackへ通知

## 📋 動作イメージ

### Slack表示例

```
📋 今週の残タスク（5件）
------------------------
• プレゼン資料の作成   〆 2025-05-20   👤 松本朋也  📌 未着手
• 🔴 代行業者のプロトタイプ作成   〆 2025-05-23   👤 松本朋也   📌 進行中
...
```

### スプレッドシートログ例

| タイムスタンプ | タスク名 | URL | 期限 | 担当者 | 進捗 | 優先度 |
|----------------|----------|-----|------|--------|------|--------|
| 2025-05-19 10:00:00 | プレゼン資料の作成 | https://notion.so/... | 2025-05-20 | 松本朋也 | 未着手 | 中 |
| 2025-05-19 10:00:00 | 代行業者のプロトタイプ作成 | https://notion.so/... | 2025-05-23 | 松本朋也 | 進行中 | 高 |

## 🚀 セットアップ手順

### 準備物

1. Notionアカウント（統合作成権限必須）
2. タスク管理用Notionデータベースへのアクセス権限
3. Slackワークスペース（Webhook設定権限必須）
4. Googleアカウント

### Notion設定

1. [my-integrations](https://notion.so/my-integrations)にアクセス
2. 「New integration」をクリック、名前を入力
3. 「Capabilities」タブで「Read content」のみON（他はOFF）
4. 「Save changes」をクリック
5. 「Internal Integration Secret」をコピー（`ntn_`で始まる文字列）
6. タスクデータベースを開き、右上の「共有」→「接続を追加」で先ほど作成した統合を選択
7. データベースURLの末尾32文字をコピー（これがDATABASE_ID）

### Slack設定

1. [Slack API Apps](https://api.slack.com/apps)ページにアクセス
2. 「Create New App」→「From scratch」を選択
3. 名前とワークスペースを設定
4. 「Incoming Webhooks」を有効化
5. 「Add New Webhook to Workspace」をクリック
6. 通知を送信するチャンネルを選択し、許可
7. 生成されたWebhook URLをコピー

### Googleスプレッドシート設定

1. 新しいスプレッドシートを作成
2. スプレッドシートURLの`/d/`と`/edit`の間の文字列をコピー（これがSHEET_ID）

### Google Apps Script設定

1. Google Driveで「新規」→「その他」→「Google Apps Script」を選択
2. 下記の3つのファイルを作成し、コードをコピペ
3. 「ファイル」→「保存」でプロジェクトを保存
4. `debugRun()`関数を実行して初期設定と動作確認
5. `createWeeklyTrigger()`を実行して週次実行を設定

## 📁 ファイル構成

- `Code.gs` - メイン処理とトリガー設定
- `Config.gs` - 設定と定数
- `Utils.gs` - ユーティリティ関数（Notion・Slack・Sheets連携）

## 🛠️ カスタマイズ方法

### 列名の変更

Notionデータベースの列名を変更した場合は、`Config.gs`の`NOTION.COLUMNS`を編集

```javascript
COLUMNS: {
  STATUS: '進捗',    // 変更する場合はここを編集
  DUE: '期限',       // 変更する場合はここを編集
  ASSIGNEE: '担当者', // 変更する場合はここを編集
  PRIORITY: '優先度'  // 変更する場合はここを編集
}
```

### ステータス値の変更

タスクステータスの値を変更した場合は、`Config.gs`の`NOTION.STATUS_VALUES`を編集

```javascript
STATUS_VALUES: {
  TODO: '未着手',  // 変更する場合はここを編集
  WIP: '進行中',   // 変更する場合はここを編集
  DONE: '完了'     // 変更する場合はここを編集
}
```

### Slack表示形式の変更

Slackメッセージの表示形式を変更する場合は、`Config.gs`の`SLACK.FORMAT`を編集

```javascript
FORMAT: {
  HEADER: '{emoji} 今週の残タスク（{count}件）', // ヘッダー形式
  TASK_ROW: '{bullet} <{url}|*{title}*>　{due_emoji} {due}　{assignee_emoji} {assignees}　{status_emoji} {status}' // タスク行
}
```

## 🔄 定期実行の設定

- デフォルトでは毎週月曜日の午前9時に実行
- 変更する場合は`createWeeklyTrigger()`内の設定を編集

```javascript
ScriptApp.newTrigger('weeklyNotionReport')
  .timeBased()
  .onWeekDay(ScriptApp.WeekDay.MONDAY) // 曜日を変更
  .atHour(9) // 時間を変更
  .create();
```

## 実際のスプレッドシートの例は以下になります。
https://docs.google.com/spreadsheets/d/16TBt6Zxbse7GqnUacOKHjCj9gjZfBjzXRf3IiCUdfGc/edit?gid=0#gid=0


## 📊 ログ記録の詳細

スプレッドシートには以下の情報が記録されます：

- タイムスタンプ - 実行日時（日本時間）
- タスク名 - Notionのタスク名
- URL - タスクへの直接リンク
- 期限 - タスクの期限
- 担当者 - 担当者名（複数の場合はカンマ区切り）
- 進捗 - タスクのステータス
- 優先度 - タスクの優先度

## ⚠️ トラブルシューティング

### 「Notion API エラー: 401 Unauthorized」
- Notionトークンが無効またはデータベースへのアクセス権限がない
- 解決策: 統合を再作成し、データベースに接続し直す

### 「Notion API エラー: 400 Bad Request」
- 列名や型の不一致
- 解決策: `debugDescribeDB()`を実行して列名・型を確認し、`Config.gs`の設定を修正

### 「Slack webhook エラー」
- Webhook URLが無効または期限切れ
- 解決策: Slackアプリ設定から新しいWebhook URLを取得

### 「スクリプトプロパティが未設定です」
- APIキーなどが設定されていない
- 解決策: `oneTimeSetup()`または`manualOneTimeSetup()`を実行して設定
