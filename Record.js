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
  if(sheet.getRange(lastRow, 2).getValue() === "【奈良すこやか保育園】入室のお知らせ") {
    Logger.log("入室処理の開始");
    // 入室記録を取得
    let enter_time = sheet.getRange(lastRow, 1).getValue();
    // カレンダーを作成
    createEvent("登園", enter_time, enter_time, null);
    Logger.log("入室処理の終了");
    // メールを既読にする
    for (let i = 0; i < threads.length; i++) {
      threads[i].markRead();
    }

    return;
  }

  // 退室の場合、入室から退室までの時間をカレンダーに設定
  Logger.log("退室処理の開始");
  createTimeCheckingOut();
  Logger.log("退室処理の終了");
  // メールを既読にする
  for (let i = 0; i < threads.length; i++) {
    threads[i].markRead();
  }

  return;
}

// 入退室記録をカレンダーに作成する
function createTimeCheckingOut() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let enter_time = sheet.getRange(sheet.getLastRow()-1, 1).getValue();
  let checking_out = sheet.getRange(sheet.getLastRow(), 1).getValue();
  let checking_out_time = getDiff(enter_time, checking_out);
  Logger.log("入退室:" + checking_out_time);

  // カレンダーを作成する
  createEvent("保育園", enter_time, checking_out, checking_out_time);
}

// カレンダーに日付をセットする
function createEvent(x_calendar_nm, x_from, x_to, x_description){
  Logger.log("カレンダー名:" + x_calendar_nm);

  // 光ちゃんカレンダー
  let hikari_calendar = PropertiesService.getScriptProperties().getProperty("HIKARI_CALENDAR");
  let calendar = CalendarApp.getCalendarById(hikari_calendar);
  calendar.createEvent(x_calendar_nm, new Date(x_from), new Date(x_to), {description: x_description});

  return;
}

// 差分の時間を取得
function getDiff(x_from, x_to) {

  let from = Moment.moment(x_from);
  Logger.log("from:" + from);
  let to = Moment.moment(x_to);
  Logger.log("to:" + to);

  // 時間計算
  let hour = to.diff(from,"h");
  Logger.log("hour:" + hour);

  // 分計算
  let minute = to.diff(from,"m");
  Logger.log("minute:" + minute);

  let mm = minute - (hour * 60);
  Logger.log("分:" + mm);

  // 結果
  let result = hour + '時間' + mm + '分';
  Logger.log('結果：' + result);
  return result;
}