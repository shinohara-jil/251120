/**
 * メール自動転記スクリプト
 * 
 * 要件:
 * - shinohara@jilinc.co.jp 宛のメール
 * - 件名に「入稿が完了しました」を含む
 * - スプレッドシート(1XO-56rbZtiwbCYpnObHMWxjUw50kqxhCATaEanr-x5)に転記
 * - 転記内容: A列=No, B列=差出人, C列=内容サマリ(先頭100文字)
 * - 重複防止: 処理済みメールにはラベル 'Processed' を付与
 */

// 設定値
const CONFIG = {
  SEARCH_QUERY: 'to:shinohara@jilinc.co.jp subject:"入稿が完了しました" -label:Processed',
  SPREADSHEET_ID: '1XO-56rbZtiwbCYpnObHMWxjUw50kqxhCATaEanr-x5',
  SHEET_NAME: 'シート1',
  PROCESSED_LABEL_NAME: 'Processed',
  SUMMARY_LENGTH: 100
};

function main() {
  // 1. スプレッドシートとシートを取得
  const sheet = getSheet_();
  if (!sheet) return;

  // 2. メールを検索
  const threads = GmailApp.search(CONFIG.SEARCH_QUERY);
  if (threads.length === 0) {
    console.log('対象のメールは見つかりませんでした。');
    return;
  }

  // 3. 処理済みラベルを取得または作成
  const label = getOrCreateLabel_(CONFIG.PROCESSED_LABEL_NAME);

  // 4. メールを処理して転記データを作成
  const newRows = [];
  let currentNo = getLastNo_(sheet);

  // スレッドごとに処理 (古い順に処理したい場合は threads.reverse() を使用)
  // ここでは検索結果順（通常は新しい順）だが、Noを振るため、古いものから順に処理した方が自然な場合もある
  // 今回は検索結果順（新しい順）で取得されるが、リストに追加する際は
  // 取得したスレッドを逆順にして古いものから順に追記する形にする
  threads.reverse().forEach(thread => {
    const messages = thread.getMessages();
    
    messages.forEach(message => {
      // 既にラベルが付いているメッセージはスキップ（スレッド単位で検索しているため念のため）
      // ただし検索条件で -label:Processed しているので、基本的には未処理スレッドの全メッセージが対象
      // ここではシンプルにスレッド内の全メッセージを対象とするが、
      // 厳密にはメッセージ単位で判定が必要な場合もある。
      // 今回は「スレッドにラベルが付いていなければ未処理」とみなす。
      
      const from = message.getFrom();
      const body = message.getPlainBody();
      const summary = body.substring(0, CONFIG.SUMMARY_LENGTH).replace(/\r?\n/g, ' '); // 改行をスペースに置換

      currentNo++;
      newRows.push([
        currentNo,
        from,
        summary
      ]);
    });
    
    // スレッドに処理済みラベルを付与
    thread.addLabel(label);
  });

  // 5. スプレッドシートに書き込み
  if (newRows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newRows.length, 3).setValues(newRows);
    console.log(`${newRows.length} 件のメールを転記しました。`);
  } else {
    console.log('転記対象のメッセージはありませんでした。');
  }
}

/**
 * スプレッドシートとシートを取得するヘルパー関数
 */
function getSheet_() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      console.error(`シート "${CONFIG.SHEET_NAME}" が見つかりません。`);
      return null;
    }
    return sheet;
  } catch (e) {
    console.error(`スプレッドシートのオープンに失敗しました: ${e.message}`);
    return null;
  }
}

/**
 * ラベルを取得、なければ作成するヘルパー関数
 */
function getOrCreateLabel_(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

/**
 * 現在の最終Noを取得するヘルパー関数
 * A列の最終行の値を取得する。数値でなければ0を返す。
 */
function getLastNo_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0; // ヘッダーのみまたは空の場合

  const lastVal = sheet.getRange(lastRow, 1).getValue();
  return isNaN(lastVal) ? 0 : Number(lastVal);
}
