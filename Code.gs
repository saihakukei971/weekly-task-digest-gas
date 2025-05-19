/**
 * Notion → Slack 週次タスクレポート
 * メイン処理
 * 作成日: 2024-03-25
 */

/**
 * 初回設定用関数（安全なプロパティ保存）
 * 一度実行後はコメントアウトまたは削除すること
 */
function oneTimeSetup() {
  const ui = SpreadsheetApp.getUi();
  
  // Notionトークン入力
  const tokenResponse = ui.prompt(
    'Notion APIトークン設定',
    'Notion統合シークレットを入力してください（ntn_から始まる文字列）:', 
    ui.ButtonSet.OK_CANCEL
  );
  if (tokenResponse.getSelectedButton() == ui.Button.CANCEL) return;
  
  // データベースID入力
  const dbIdResponse = ui.prompt(
    'NotionデータベースID設定',
    'データベースIDを入力してください（URLの末尾32文字）:', 
    ui.ButtonSet.OK_CANCEL
  );
  if (dbIdResponse.getSelectedButton() == ui.Button.CANCEL) return;
  
  // Slack Webhook入力
  const webhookResponse = ui.prompt(
    'Slack Webhook設定',
    'Slack Incoming WebhookのURLを入力してください:', 
    ui.ButtonSet.OK_CANCEL
  );
  if (webhookResponse.getSelectedButton() == ui.Button.CANCEL) return;
  
  // スプレッドシートID入力
  const sheetIdResponse = ui.prompt(
    'スプレッドシートID設定',
    'ログ記録用スプレッドシートのIDを入力してください:', 
    ui.ButtonSet.OK_CANCEL
  );
  if (sheetIdResponse.getSelectedButton() == ui.Button.CANCEL) return;
  
  // プロパティを設定
  PROPS.setProperties({
    NOTION_TOKEN: tokenResponse.getResponseText(),
    DATABASE_ID: dbIdResponse.getResponseText(),
    SLACK_WEBHOOK: webhookResponse.getResponseText(),
    SHEET_ID: sheetIdResponse.getResponseText()
  }, true);
  
  ui.alert('設定完了', 'すべての設定が完了しました！', ui.ButtonSet.OK);
}

/**
 * 手動設定用関数（GUIが使えない場合）
 * 必ず実行後は削除すること
 */
function manualOneTimeSetup() {
  PROPS.setProperties(
    {
      NOTION_TOKEN: 'ntn_XXXXXXXXXXXXXXXXXXXXX',   // Notionトークン
      DATABASE_ID: 'XXXXXXXXXXXXXXXXXXXXX',        // DBのID
      SLACK_WEBHOOK: 'https://hooks.slack.com/services/XXX/YYY/ZZZ', // Webhook URL
      SHEET_ID: '16TBt6Zxbse7GqnUacOKHjCj9gjZfBjzXRf3IiCUdfGc'            // シートID
    },
    true
  );
  console.log('✅ プロパティを設定しました。この関数は削除してください。');
}

/**
 * メイン処理（週次実行される関数）
 * NotionからタスクDBを取得し、Slackに通知、スプレッドシートに記録
 */
function weeklyNotionReport() {
  try {
    // プロパティから設定を取得
    const NOTION_TOKEN = PROPS.getProperty('NOTION_TOKEN');
    const DATABASE_ID = PROPS.getProperty('DATABASE_ID');
    const SLACK_WEBHOOK = PROPS.getProperty('SLACK_WEBHOOK');
    const SHEET_ID = PROPS.getProperty('SHEET_ID');
    
    // プロパティ検証
    if (!NOTION_TOKEN || !DATABASE_ID || !SLACK_WEBHOOK || !SHEET_ID) {
      throw new Error(ERROR.MISSING_PROPS);
    }
    
    // Notionからタスク取得
    console.log('Notionからタスク取得中...');
    const tasks = fetchTasksFromNotion(NOTION_TOKEN, DATABASE_ID);
    
    // タスクがない場合はスキップ
    if (!tasks.length) {
      console.log('📭 対象タスク 0 件 → 処理をスキップします');
      return;
    }
    
    // スプレッドシートにログ記録
    console.log(`スプレッドシートにログ記録中...（${tasks.length}件）`);
    logTasksToSheet(SHEET_ID, tasks);
    
    // Slackメッセージ作成
    console.log('Slackメッセージを作成中...');
    const blocks = buildSlackBlocks(tasks);
    
    // Slackに投稿
    console.log('Slackへ投稿中...');
    postToSlack(SLACK_WEBHOOK, blocks);
    
    console.log('週次レポート処理が完了しました');
  } catch (error) {
    console.error('週次レポート処理でエラーが発生しました:', error);
    
    // エラーをSlackに通知
    try {
      const webhook = PROPS.getProperty('SLACK_WEBHOOK');
      if (webhook) {
        sendErrorToSlack(webhook, error.toString());
      }
    } catch (slackError) {
      console.error('エラー通知の送信に失敗しました:', slackError);
    }
    
    throw error;
  }
}

/**
 * DBの構造を確認するデバッグ関数
 * 列名や型の確認に使用
 */
function debugDescribeDB() {
  const token = PROPS.getProperty('NOTION_TOKEN');
  const dbId = PROPS.getProperty('DATABASE_ID');
  
  try {
    // DBスキーマ取得
    const url = `${NOTION.BASE_URL}/databases/${dbId}`;
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Notion-Version': NOTION.API_VERSION
      }
    });
    
    if (resp.getResponseCode() !== 200) {
      throw new Error(ERROR.NOTION_API + resp.getContentText());
    }
    
    // 結果をログ出力
    const schema = JSON.parse(resp.getContentText());
    console.log('データベース名: ' + schema.title[0].plain_text);
    console.log('列の型一覧:');
    
    const properties = schema.properties;
    for (const [name, prop] of Object.entries(properties)) {
      console.log(`- ${name}: ${prop.type}`);
    }
    
    return schema;
  } catch (error) {
    console.error('DBスキーマの取得に失敗しました:', error);
    throw error;
  }
}

/**
 * 手動テスト用関数
 */
function debugRun() {
  console.log('デバッグ実行を開始します...');
  weeklyNotionReport();
  console.log('デバッグ実行が完了しました');
}

/**
 * 週次トリガーを作成する関数
 */
function createWeeklyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'weeklyNotionReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // 新しい週次トリガーを作成（月曜9:00AM）
  ScriptApp.newTrigger('weeklyNotionReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
  
  console.log('✅ 毎週月曜9:00にweeklyNotionReport()を実行するトリガーを設定しました');
}