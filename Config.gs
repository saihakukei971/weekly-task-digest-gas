/**
 * Notion → Slack 週次タスクレポート
 * 設定と定数
 * 作成日: 2024-04-19
 */

// スクリプトプロパティ (安全な保存領域)
const PROPS = PropertiesService.getScriptProperties();

// Notion API設定
const NOTION = {
  API_VERSION: '2022-06-28',
  BASE_URL: 'https://api.notion.com/v1',
  // Notionデータベースの列名 (DBに合わせて変更)
  COLUMNS: {
    STATUS: '進捗',    // Status型プロパティ
    DUE: '期限',       // Date型プロパティ
    ASSIGNEE: '担当者', // People型プロパティ
    PRIORITY: '優先度'  // Select型プロパティ (オプション)
  },
  // ステータス値 (DBに合わせて変更)
  STATUS_VALUES: {
    TODO: '未着手',
    WIP: '進行中',
    DONE: '完了'
  },
  PAGE_SIZE: 100 // 1リクエストあたりの最大取得数
};

// Slackメッセージ形式
const SLACK = {
  // 絵文字
  EMOJI: {
    REPORT: '📋',  // レポートアイコン
    TASK: '•',    // タスク箇条書き
    DUE: '〆',     // 期限
    ASSIGNEE: '👤', // 担当者
    STATUS: '📌',  // 状態
    PRIORITY: '🔴', // 高優先度
    ERROR: '❌',   // エラー
    SUCCESS: '✅'  // 成功
  },
  // メッセージ形式
  FORMAT: {
    HEADER: '{emoji} 今週の残タスク（{count}件）',
    TASK_ROW: '{bullet} <{url}|*{title}*>　{due_emoji} {due}　{assignee_emoji} {assignees}　{status_emoji} {status}'
  }
};

// スプレッドシート設定
const SHEET = {
  HEADERS: ['タイムスタンプ', 'タスク名', 'URL', '期限', '担当者', '進捗', '優先度'],
  LOG_SHEET_NAME: 'タスクログ（週次）'
};

// エラーメッセージ
const ERROR = {
  MISSING_PROPS: '❌ スクリプトプロパティ（NOTION_TOKEN/DATABASE_ID/SLACK_WEBHOOK/SHEET_ID）が未設定です',
  NOTION_API: 'Notion API エラー: ',
  SLACK_WEBHOOK: 'Slack webhook エラー: ',
  SHEET_LOG: 'スプレッドシートログ記録エラー: '
};