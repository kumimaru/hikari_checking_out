// メールから取得した入退室記録をカレンダーに追記
function checking_out() {
  // メールの送信者：奈良すこやか保育園 未読
  const query = 'from:cdm7019rqBk@codmon.com is:unread';
  // メールの検索範囲
  const min = 0;
  const max = 100;

  // メールを検索
  const threads = GmailApp.search(query, min, max);
  const messagesForThreads = GmailApp.getMessagesForThreads(threads);

  // 取得したメールから日時と件名を取得
  const values = [];
  for (const messages of messagesForThreads) {
    const message = messages[0];
    const record = [
      message.getDate()
      , message.getSubject()
    ];
    values.push(record);
  }
  if (values.length <= 0) {
    // メールが取得できなければ処理終了
    return;
  }

  // シートのフォーマットを指定の日時に設定
  let sheet = SpreadsheetApp.getActiveSheet();
  let lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1).setNumberFormat("yyyy/mm/dd HH:mm:ss")

  // シートにメールの取得日時と件名を記録
  sheet.getRange(lastRow, 1, values.length, values[0].length).setValues(values);

  // 取得したメールの件名が入室か退室のどちらであるかを判定
  if(sheet.getRange(lastRow, 2) === "【奈良すこやか保育園】入室のお知らせ") {
    Logger.log("入室処理の開始");
    // 入室記録を取得
    let checking_out = sheet.getRange(lastRow, 1).getValue();
    // 入室のカレンダー名を取得
    let calendar_nm = sheet.getRange(lastRow, 2).getValue();
    // カレンダーを作成
    createEvent(calendar_nm, checking_out, checking_out, null);
    Logger.log("入室処理の終了");
    return;
  }

  // 退室の場合、入室から退室までの時間をカレンダーに設定
  Logger.log("退室処理の開始");
  createTimeCheckingOut(lastRow - 1, lastRow);
  Logger.log("退室処理の終了");
  return;
}

// 入退室記録をカレンダーに作成する
function createTimeCheckingOut(x_first_row, x_last_row) {
  let sheet = SpreadsheetApp.getActiveSheet();
  let enter_time = sheet.getRange(x_first_row, 1).getValue();
  let checking_out = sheet.getRange(x_last_row, 1).getValue();
  let checking_out_time = getDiff(checking_out, enter_time);
  Logger.log("入退室:" + checking_out_time);

  // カレンダーを作成する
  let calendar_nm = sheet.getRange(x_last_row, 2).getValue();
  createEvent(calendar_nm, checking_out, enter_time, checking_out_time);
}

// カレンダーに日付をセットする
function createEvent(x_calendar_nm, x_last_time, x_enter_time, x_sleeping_time){
  Logger.log("カレンダー名:" + x_calendar_nm);

  // 光ちゃんカレンダー
  let hikari_calendar = PropertiesService.getScriptProperties().getProperty("HIKARI_CALENDAR");
  let calendar = CalendarApp.getCalendarById(hikari_calendar);
  calendar.createEvent(x_calendar_nm, new Date(x_last_time), new Date(x_enter_time) , {description: x_sleeping_time});

  return;
}

// 差分の時間を取得
function getDiff(x_last_time, x_enter_time) {

  let checking_out = Moment.moment(x_last_time);
  let enter_time = Moment.moment(x_enter_time);
  Logger.log("checking_out:" + checking_out);
  Logger.log("enter_time:" + enter_time);

  // 時間計算
  let hour = enter_time.diff(checking_out,"h");
  Logger.log("hour:" + hour);

  // 分計算
  let minute = enter_time.diff(checking_out,"m");
  Logger.log("minute:" + minute);

  let mm = minute - (hour * 60);
  Logger.log("分:" + mm);

  // 結果
  let result = hour + '時間' + mm + '分';
  Logger.log('結果：' + result);
  return result;
}