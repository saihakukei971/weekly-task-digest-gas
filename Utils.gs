/**
 * Notion → Slack 週次タスクレポート
 * ユーティリティ関数
 * 作成日: 2024-04-19
 */

/**
 * Notionからタスクを取得する関数
 * @param {string} token - Notion APIトークン
 * @param {string} databaseId - NotionデータベースID
 * @return {Array} タスク情報の配列
 */
function fetchTasksFromNotion(token, databaseId) {
  // 未完了タスクのフィルター作成
  const filter = {
    or: [
      { property: NOTION.COLUMNS.STATUS, status: { equals: NOTION.STATUS_VALUES.TODO } },
      { property: NOTION.COLUMNS.STATUS, status: { equals: NOTION.STATUS_VALUES.WIP } }
    ]
  };
  
  // 期限でソート
  const sorts = [{ property: NOTION.COLUMNS.DUE, direction: 'ascending' }];
  
  // リクエストペイロード作成
  const payload = {
    filter: filter,
    sorts: sorts,
    page_size: NOTION.PAGE_SIZE
  };
  
  // APIエンドポイント
  const url = `${NOTION.BASE_URL}/databases/${databaseId}/query`;
  
  try {
    // Notion APIリクエスト実行
    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: {
        'Authorization': 'Bearer ' + token,
        'Notion-Version': NOTION.API_VERSION
      }
    });
    
    // レスポンスコード確認
    if (resp.getResponseCode() !== 200) {
      throw new Error(ERROR.NOTION_API + resp.getContentText());
    }
    
    // レスポンス解析
    const data = JSON.parse(resp.getContentText());
    
    // ページネーション対応（100件以上ある場合）
    let results = data.results || [];
    let hasMore = data.has_more;
    let nextCursor = data.next_cursor;
    
    // 100件以上ある場合は追加取得
    while (hasMore && nextCursor) {
      payload.start_cursor = nextCursor;
      
      const nextResp = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        headers: {
          'Authorization': 'Bearer ' + token,
          'Notion-Version': NOTION.API_VERSION
        }
      });
      
      if (nextResp.getResponseCode() !== 200) {
        throw new Error(ERROR.NOTION_API + nextResp.getContentText());
      }
      
      const nextData = JSON.parse(nextResp.getContentText());
      results = results.concat(nextData.results || []);
      hasMore = nextData.has_more;
      nextCursor = nextData.next_cursor;
    }
    
    // タスク情報を整形して返却
    return results.map(page => {
      const props = page.properties;
      
      // タイトルプロパティを探す（カスタム名の場合があるため）
      const titleKey = Object.keys(props).find(k => props[k].type === 'title');
      const title = titleKey ? props[titleKey].title.map(t => t.plain_text).join('') : 'タイトルなし';
      
      // 期限取得
      const due = props[NOTION.COLUMNS.DUE]?.date?.start ?? '未設定';
      
      // 進捗状態取得
      const progress = props[NOTION.COLUMNS.STATUS]?.status?.name ?? '—';
      
      // 担当者取得（複数可）
      const assignees = (props[NOTION.COLUMNS.ASSIGNEE]?.people || [])
        .map(p => p.name || p.person?.email || '未割当')
        .join(', ') || '未割当';
      
      // 優先度取得（存在する場合）
      const priority = props[NOTION.COLUMNS.PRIORITY]?.select?.name || '—';
      
      return { 
        title, 
        url: page.url, 
        due, 
        assignees, 
        progress,
        priority
      };
    });
  } catch (error) {
    console.error('Notionからの取得に失敗しました:', error);
    throw error;
  }
}

/**
 * Slack用のブロックメッセージを構築する関数
 * @param {Array} tasks - タスク情報の配列
 * @return {Array} Slackブロックの配列
 */
function buildSlackBlocks(tasks) {
  // ヘッダーブロック
  const header = {
    type: 'header',
    text: { 
      type: 'plain_text', 
      text: SLACK.FORMAT.HEADER
        .replace('{emoji}', SLACK.EMOJI.REPORT)
        .replace('{count}', tasks.length)
    }
  };
  
  // 区切り線
  const divider = { type: 'divider' };
  
  // タスクごとのブロック
  const list = tasks.map(t => {
    // 高優先度かチェック
    const isPriorityHigh = t.priority && t.priority.includes('高');
    const priorityPrefix = isPriorityHigh ? `${SLACK.EMOJI.PRIORITY} ` : '';
    
    // タスク行のテキスト整形
    const text = SLACK.FORMAT.TASK_ROW
      .replace('{bullet}', SLACK.EMOJI.TASK)
      .replace('{url}', t.url)
      .replace('{title}', t.title)
      .replace('{due_emoji}', SLACK.EMOJI.DUE)
      .replace('{due}', t.due)
      .replace('{assignee_emoji}', SLACK.EMOJI.ASSIGNEE)
      .replace('{assignees}', t.assignees)
      .replace('{status_emoji}', SLACK.EMOJI.STATUS)
      .replace('{status}', t.progress);
    
    return {
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: priorityPrefix + text
      }
    };
  });
  
  return [header, divider, ...list];
}

/**
 * Slackにメッセージを投稿する関数
 * @param {string} webhook - Slack Webhook URL
 * @param {Array} blocks - Slackブロックの配列
 */
function postToSlack(webhook, blocks) {
  try {
    const resp = UrlFetchApp.fetch(webhook, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ blocks })
    });
    
    // エラーチェック
    if (resp.getResponseCode() >= 300) {
      throw new Error(ERROR.SLACK_WEBHOOK + resp.getContentText());
    }
    
    console.log(`${SLACK.EMOJI.SUCCESS} Slackへ送信完了: ${blocks.length - 2} 件`);
  } catch (error) {
    console.error('Slackへの投稿に失敗しました:', error);
    throw error;
  }
}

/**
 * エラーをSlackに通知する関数
 * @param {string} webhook - Slack Webhook URL
 * @param {string} errorMessage - エラーメッセージ
 */
function sendErrorToSlack(webhook, errorMessage) {
  try {
    // エラー通知用ブロック
    const errorBlocks = [
      {
        type: 'header',
        text: { 
          type: 'plain_text', 
          text: `${SLACK.EMOJI.ERROR} Notion-Slack 週次レポートエラー` 
        }
      },
      { type: 'divider' },
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: `*エラー内容:* \n\`\`\`${errorMessage}\`\`\``
        }
      },
      {
        type: 'context',
        elements: [
          {
            type: 'mrkdwn',
            text: `発生時刻: ${new Date().toLocaleString('ja-JP')}`
          }
        ]
      }
    ];
    
    // エラー通知送信
    UrlFetchApp.fetch(webhook, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ blocks: errorBlocks })
    });
  } catch (e) {
    // Slack通知に失敗した場合はログに記録
    console.error('エラー通知の送信に失敗しました:', e);
  }
}

/**
 * タスクをスプレッドシートに記録する関数
 * @param {string} sheetId - スプレッドシートID
 * @param {Array} tasks - タスク情報の配列
 */
function logTasksToSheet(sheetId, tasks) {
  try {
    // スプレッドシートを開く
    const ss = SpreadsheetApp.openById(sheetId);
    let sheet = ss.getSheetByName(SHEET.LOG_SHEET_NAME);
    
    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = ss.insertSheet(SHEET.LOG_SHEET_NAME);
      sheet.appendRow(SHEET.HEADERS);
      
      // ヘッダー行の書式設定
      sheet.getRange(1, 1, 1, SHEET.HEADERS.length)
        .setBackground('#f3f3f3')
        .setFontWeight('bold');
      
      // 列幅を自動調整
      sheet.autoResizeColumns(1, SHEET.HEADERS.length);
    }
    
    // 現在のタイムスタンプ（日本時間）
    const timestamp = Utilities.formatDate(
      new Date(), 
      'Asia/Tokyo', 
      'yyyy-MM-dd HH:mm:ss'
    );
    
    // 各タスクを行として追加
    tasks.forEach(task => {
      sheet.appendRow([
        timestamp,
        task.title,
        task.url,
        task.due,
        task.assignees,
        task.progress,
        task.priority
      ]);
    });
    
    // 最新の行を反転色で強調（視認性向上）
    const lastRow = sheet.getLastRow();
    const newRows = tasks.length;
    if (lastRow > 1 && newRows > 0) {
      sheet.getRange(lastRow - newRows + 1, 1, newRows, SHEET.HEADERS.length)
        .setBackground('#e8f5e9');
    }
    
    console.log(`${SLACK.EMOJI.SUCCESS} スプレッドシートへログ記録: ${tasks.length} 件`);
  } catch (error) {
    console.error('スプレッドシートへの記録に失敗しました:', error);
    // ログ失敗してもメイン処理は続行するため、ここでは例外を再スローしない
  }
}