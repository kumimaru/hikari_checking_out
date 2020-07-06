// メールから取得した入退室記録をカレンダーに追記
function checking_out() {
  // TODO: メールを取得する処理を実装する
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
  if (values.length < 0) {
    // メールが取得できなければ処理終了
    return;
  }

  // シートのフォーマットを指定の日時に設定
  var lastRow = SpreadsheetApp.getActiveSheet().getLastRow();
  sheet.getRange(lastRow, 1).setNumberFormat("yyyy/mm/dd HH:mm:ss")
  // シートにメールの取得日時と件名を記録
  SpreadsheetApp.getActiveSheet().getRange(lastRow, 1, values.length, values[0].length).setValues(values);

  // 取得したメールの件名が入室か退室のどちらであるかを判定
  if(values[1] === "【奈良すこやか保育園】入室のお知らせ") {
    // 入室記録を取得
    var checking_out = sheet.getRange(lastRow, 1).getValue();
    // 入室のカレンダー名を取得
    var calendar_nm = sheet.getRange(lastRow, 2).getValue();
    // カレンダーを作成
    createEvent(calendar_nm, checking_out, checking_out, null);
    return;
  }

  // 退室の場合、入室から退室までの時間をカレンダーに設定
  createTimeCheckingOut(lastRow - 1, lastRow);
}

// 入退室記録をカレンダーに作成する
function createTimeCheckingOut(x_first_row, x_last_row) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var enter_time = sheet.getRange(x_first_row, 1).getValue();
  var checking_out = sheet.getRange(x_last_row, 1).getValue();
  var checking_out_time = getDiff(checking_out, enter_time);
  Logger.log("入退室:" + checking_out_time);

  // カレンダーを作成する
  var calendar_nm = sheet.getRange(x_last_row, 2).getValue();
  createEvent(calendar_nm, checking_out, enter_time, checking_out_time);
}

// カレンダーに日付をセットする
function createEvent(x_calendar_nm, x_last_time, x_enter_time, x_sleeping_time){
  Logger.log("カレンダー名:" + x_calendar_nm);

  // 光ちゃんカレンダー
  var hikari_calendar = PropertiesService.getScriptProperties().getProperty("HIKARI_CALENDAR");
  var calendar = CalendarApp.getCalendarById(hikari_calendar);
  calendar.createEvent(x_calendar_nm, new Date(x_last_time), new Date(x_enter_time) , {description: x_sleeping_time});
}

// 差分の時間を取得
function getDiff(x_last_time, x_enter_time) {

  var checking_out = Moment.moment(x_last_time);
  var enter_time = Moment.moment(x_enter_time);
  Logger.log("checking_out:" + checking_out);
  Logger.log("enter_time:" + enter_time);

  // 時間計算
  var hour = enter_time.diff(checking_out,"h");
  Logger.log("hour:" + hour);

  // 分計算
  var minute = enter_time.diff(checking_out,"m");
  Logger.log("minute:" + minute);

  var mm = minute - (hour * 60);
  Logger.log("分:" + mm);

  // 結果
  var result = hour + '時間' + mm + '分';
  Logger.log('結果：' + result);
  return result;
}