// メールから取得した入退室記録をカレンダーに追記
function checking_out() {
  var p_sheet = SpreadsheetApp.getActiveSheet();  
  
  // 未読の入退室メールを取得
  var p_lastRow = p_sheet.getLastRow();
  // 受信日時を取得 
  p_sheet.getRange(p_lastRow, 1).setValue(new Date());
  p_sheet.getRange(p_lastRow, 1).setNumberFormat("yyyy/mm/dd HH:mm:ss")
  
  // 取得したメールの件名か入室か退室のどちらであるかを判定
  if(isWakeup(p_lastRow)) {
    // 退室の場合、入室から退室までの時間をカレンダーに設定
    createSleepDiary(p_lastRow, p_lastRow - 1);
    return;
  }
  
  // 入室のカレンダーを作成
  // カレンダー名を取得
  var p_calendar_nm = p_sheet.getRange(p_lastRow, 2).getValue();  // すやあを取得
  // 睡眠時間を取得
  var p_bed_time = p_sheet.getRange(p_lastRow, 1).getValue();
  Logger.log("すやあの時間:" + p_bed_time);
  createEvent(p_calendar_nm, p_bed_time, p_bed_time, null);
}

// 行の値が光ちゃんおはであるか判定
function isWakeup(x_row) {
  var p_sheet = SpreadsheetApp.getActiveSheet();  
  Logger.log("行の値:" + p_sheet.getRange(x_row, 2).getValue());
  return p_sheet.getRange(x_row, 2).getValue() == "睡眠" 
}

// 睡眠履歴をカレンダーに作成する
function createSleepDiary(x_wakeup_row, x_sleep_row) {
  var p_sheet = SpreadsheetApp.getActiveSheet();  

  var p_bed_time = p_sheet.getRange(x_sleep_row, 1).getValue();
  var p_wakeup_time = p_sheet.getRange(x_wakeup_row, 1).getValue();
  var p_sleeping_time = getDiff(p_bed_time, p_wakeup_time);
  Logger.log("睡眠時間:" + p_sleeping_time);

  // カレンダーを作成する
  var p_calendar_nm = p_sheet.getRange(x_wakeup_row, 2).getValue();  // 睡眠を取得
  createEvent(p_calendar_nm, p_bed_time, p_wakeup_time, p_sleeping_time);
}

// カレンダーに日付をセットする
function createEvent(x_calendar_nm, x_bed_time, x_wakeup_time, x_sleeping_time){
  Logger.log("カレンダー名:" + x_calendar_nm);
  
  // 光ちゃんカレンダー
  var p_calendar = CalendarApp.getCalendarById("bgq6sq6oh7l7ptkig2lmboulq8@group.calendar.google.com");
  p_calendar.createEvent(x_calendar_nm, new Date(x_bed_time), new Date(x_wakeup_time) , {description: x_sleeping_time}); 
}

// 睡眠時間を取得
function getDiff(x_bed_time, x_wakeup_time) {
  // テスト用
//  x_bed_time = '2018/12/16 8:19:58'
//  x_wakeup_time = '2018/12/16 11:07:20'
  
  var p_bed_time = Moment.moment(x_bed_time);
  var p_wakeup_time = Moment.moment(x_wakeup_time);
  Logger.log("bed_time:" + p_bed_time);
  Logger.log("wakeup_time:" + p_wakeup_time);
  
  // 時間計算
  var p_hour = p_wakeup_time.diff(p_bed_time,"h");
  Logger.log("p_hour:" + p_hour);
  
  // 分計算
  var p_minute = p_wakeup_time.diff(p_bed_time,"m");
  Logger.log("p_minute:" + p_minute);
  
  var p_mm = p_minute - (p_hour * 60);
  Logger.log("分:" + p_mm);
  
  // 結果
  var p_result = p_hour + '時間' + p_mm + '分';
  Logger.log('結果：' + p_result);
  return p_result;
}