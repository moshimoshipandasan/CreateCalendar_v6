/*
TimetableToCalenar:年間行事予定をカレンダーに１年間分で流し込む
Author:noboru ando
Date:2020/02/23 ver1作成
Date:2021/03/31 1年分書込み対応
Date:2021/04/01 ver3 ,区切りで複数日程対応
Date:2021/08/18 ver4 <>で囲むと時間指定できるようにしました。時間指定は<12:00-14:00>のようにハイフンを入れてください
Date:2021/09/05 ver5 全角のハイフン「ー」でカレンダーに書き込まないバグを修正しました
Date:2022/02/02 ver5「−」でカレンダーに書き込まないバグを修正しました
*/
function myFunction() {
  var sheet = SpreadsheetApp.getActiveSheet();
//var start_day = new Date(sheet.getRange(4,1).getValue());
  var result = Browser.msgBox("この時間割をGoolgeカレンダーに作成して良いですか？\\n 【注意】 この操作は取り消せません！",Browser.Buttons.OK_CANCEL);
  var CALENDAR_ID = sheet.getRange(1,5).getValue(); //カレンダーIDの取得
  if (CALENDAR_ID == '') {
    var result = Browser.msgBox("カレンダーIDが指定されていません。\\n カレンダーIDを入力して再度[作成]を実行してください。\\n 操作を終了します");
    /* プログラムの終了 */
  } else {
    var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if(result == "ok"){
    try {
      var schedule_table = sheet.getRange(3,1,31,24).getValues();
      var date = Utilities.formatDate(schedule_table[0][0], 'Asia/Tokyo', 'yyyy/MM/dd');
      var recurrence = CalendarApp.newRecurrence()   
      for (var j = 0; j < 24; j = j + 2){
        for (var i = 0; i < 31; i++){
          var tmp_date = schedule_table[i][j];
          if (tmp_date !== ''){
            var date = Utilities.formatDate(tmp_date, 'Asia/Tokyo', 'yyyy/MM/dd');
            var schedule = schedule_table[i][j + 1];
            var scheduleAry = schedule.split(','); //2021/04/01 noboru ando
            var sn = scheduleAry.length; //2021/04/01 noboru ando
            for (var n = 0; n < sn; n++){
              if (scheduleAry[n] !== ''){
                var str = zen_han(scheduleAry[n]);
                var reg1 = /.*?(?=[<])/;
                var str1 = str.match(reg1);
                if (str1 === null){
                  calendar
                    .createAllDayEvent(
                    str
                    , new Date(date.toString()) 
                  )
                } else {
                  var reg23= /(?<=[<]).*?(?=[>])/;
                  var seTime = zen_han(str.match(reg23));
//                  Logger.log(seTime);
//                  var reg2 = /(?<=[<＜]).*?(?=[-ー])/;
                  var reg2 = /.*?(?=[-ー−])/;
                  var startTime = zen_han(seTime.match(reg2));
//                  Logger.log("開始時刻" + startTime);
//                  var reg3 = /(?<=[-ー]).*?(?=[>＞])/;
                  var reg3 = /(?<=[-ー−]).*/;
                  var endTime = zen_han(seTime.match(reg3));
//                  Logger.log("終了時刻" + endTime);
                  var startDate = new Date(date.toString()+' '+ startTime.replace(/[：;；]/, ":"));
                  var endDate = new Date(date.toString()+' '+ endTime.replace(/[：;；]/, ":"));
                  calendar.createEvent(str1,startDate,endDate);
                }
              Utilities.sleep(200);
              }
            }
          }
        }
      }
      Browser.msgBox('年間行事予定のカレンダーへの流し込みが終了しました');
    } catch(e) {
      Browser.msgBox('エラーが発生しました:' + e.message);
    }
  }
  }      
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('年間行事予定');
  menu.addItem('カレンダーへ書き込み実行', 'myFunction');
  menu.addItem('祝日を行事予定に追加', 'addHolidaysToSchedule');
  menu.addItem('祝日を行事予定から削除', 'removeHolidaysFromSchedule');
  menu.addItem('毎週の予定を追加', 'addWeeklySchedule');
  menu.addItem('毎週の予定を削除', 'removeWeeklySchedule');
  menu.addToUi();
}

/**
 * 毎週の予定を行事予定から削除する関数
 * ユーザーが入力した予定内容をカレンダーの行事予定欄から削除します
 */
function removeWeeklySchedule() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var year = sheet.getRange("A1").getValue(); // 西暦を取得
  
  if (!year || isNaN(year)) {
    Browser.msgBox("A1セルに有効な西暦を入力してください。");
    return;
  }
  
  // 予定内容の入力ダイアログを表示
  var ui = SpreadsheetApp.getUi();
  var eventResponse = ui.prompt(
    '削除する予定内容を入力',
    '行事予定から削除したい予定内容を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  // ダイアログの結果を取得
  var eventButton = eventResponse.getSelectedButton();
  var eventText = eventResponse.getResponseText().trim();
  
  // キャンセルボタンが押された場合は終了
  if (eventButton === ui.Button.CANCEL || eventText === "") {
    return;
  }
  
  // カレンダーデータの範囲（4月から3月まで）
  var calendarRange = [
    {col: 1, eventCol: 2},    // 4月: A列（日付）、B列（行事予定）
    {col: 3, eventCol: 4},    // 5月: C列（日付）、D列（行事予定）
    {col: 5, eventCol: 6},    // 6月
    {col: 7, eventCol: 8},    // 7月
    {col: 9, eventCol: 10},   // 8月
    {col: 11, eventCol: 12},  // 9月
    {col: 13, eventCol: 14},  // 10月
    {col: 15, eventCol: 16},  // 11月
    {col: 17, eventCol: 18},  // 12月
    {col: 19, eventCol: 20},  // 1月
    {col: 21, eventCol: 22},  // 2月
    {col: 23, eventCol: 24}   // 3月
  ];
  
  var calendarData = sheet.getRange(3, 1, 31, 24).getValues(); // カレンダーデータを取得
  var updatedCells = 0;
  
  // カレンダーの各月をループ
  for (var m = 0; m < calendarRange.length; m++) {
    var monthCol = calendarRange[m].col - 1;     // 0ベースのインデックスに変換
    var eventCol = calendarRange[m].eventCol - 1; // 0ベースのインデックスに変換
    
    // 各月の日付をループ
    for (var d = 0; d < 31; d++) {
      var calDate = calendarData[d][monthCol];
      
      // 日付が空でない場合
      if (calDate !== "") {
        var currentEvent = calendarData[d][eventCol] || "";
        
        // 行事予定が空でない場合
        if (currentEvent !== "") {
          // 指定された予定が含まれているか確認
          if (currentEvent.indexOf(eventText) !== -1) {
            var eventItems = currentEvent.split(","); // カンマで区切られた行事予定を配列に
            var newEventItems = [];
            
            // 各行事予定項目をループ
            for (var e = 0; e < eventItems.length; e++) {
              var eventItem = eventItems[e].trim();
              
              // 削除対象の予定でない場合のみ新しい配列に追加
              if (eventItem !== eventText) {
                newEventItems.push(eventItem);
              }
            }
            
            // 更新された行事予定をセットし、更新カウントを増やす
            calendarData[d][eventCol] = newEventItems.join(",");
            updatedCells++;
          }
        }
      }
    }
  }
  
  // 更新されたデータをシートに書き込む
  if (updatedCells > 0) {
    sheet.getRange(3, 1, 31, 24).setValues(calendarData);
    Browser.msgBox(updatedCells + "件の予定を行事予定から削除しました。");
  } else {
    Browser.msgBox("削除する予定はありませんでした。");
  }
}

/**
 * 特定の曜日に毎週同じ予定を行事予定に追加する関数
 * ユーザーが選択した曜日に一致するカレンダーの日付の行事予定欄に予定を追加します
 */
function addWeeklySchedule() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var year = sheet.getRange("A1").getValue(); // 西暦を取得
  
  if (!year || isNaN(year)) {
    Browser.msgBox("A1セルに有効な西暦を入力してください。");
    return;
  }
  
  // 曜日の選択肢
  var daysOfWeek = ["日", "月", "火", "水", "木", "金", "土"];
  var dayIndex = -1;
  
  // 曜日の選択ダイアログを表示
  var ui = SpreadsheetApp.getUi();
  var dayResponse = ui.prompt(
    '曜日を選択',
    '追加したい曜日を入力してください（日、月、火、水、木、金、土）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  // ダイアログの結果を取得
  var dayButton = dayResponse.getSelectedButton();
  var dayText = dayResponse.getResponseText().trim();
  
  // キャンセルボタンが押された場合は終了
  if (dayButton === ui.Button.CANCEL) {
    return;
  }
  
  // 入力された曜日が有効かチェック
  dayIndex = daysOfWeek.indexOf(dayText);
  if (dayIndex === -1) {
    Browser.msgBox("有効な曜日を入力してください（日、月、火、水、木、金、土）。");
    return;
  }
  
  // 予定内容の入力ダイアログを表示
  var eventResponse = ui.prompt(
    '予定内容を入力',
    '毎週' + dayText + '曜日に追加する予定内容を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  // ダイアログの結果を取得
  var eventButton = eventResponse.getSelectedButton();
  var eventText = eventResponse.getResponseText().trim();
  
  // キャンセルボタンが押された場合は終了
  if (eventButton === ui.Button.CANCEL || eventText === "") {
    return;
  }
  
  // カレンダーデータの範囲（4月から3月まで）
  var calendarRange = [
    {col: 1, eventCol: 2},    // 4月: A列（日付）、B列（行事予定）
    {col: 3, eventCol: 4},    // 5月: C列（日付）、D列（行事予定）
    {col: 5, eventCol: 6},    // 6月
    {col: 7, eventCol: 8},    // 7月
    {col: 9, eventCol: 10},   // 8月
    {col: 11, eventCol: 12},  // 9月
    {col: 13, eventCol: 14},  // 10月
    {col: 15, eventCol: 16},  // 11月
    {col: 17, eventCol: 18},  // 12月
    {col: 19, eventCol: 20},  // 1月
    {col: 21, eventCol: 22},  // 2月
    {col: 23, eventCol: 24}   // 3月
  ];
  
  var calendarData = sheet.getRange(3, 1, 31, 24).getValues(); // カレンダーデータを取得
  var updatedCells = 0;
  
  // カレンダーの各月をループ
  for (var m = 0; m < calendarRange.length; m++) {
    var monthCol = calendarRange[m].col - 1;     // 0ベースのインデックスに変換
    var eventCol = calendarRange[m].eventCol - 1; // 0ベースのインデックスに変換
    
    // 各月の日付をループ
    for (var d = 0; d < 31; d++) {
      var calDate = calendarData[d][monthCol];
      
      // 日付が空でない場合
      if (calDate !== "") {
        // 日付の曜日を取得
        var calDayOfWeek = calDate.getDay(); // 0:日曜日, 1:月曜日, ..., 6:土曜日
        
        // 選択した曜日と一致する場合
        if (calDayOfWeek === dayIndex) {
          var currentEvent = calendarData[d][eventCol] || "";
          
          // 既に同じ予定が入力されていない場合のみ追加
          if (currentEvent.indexOf(eventText) === -1) {
            if (currentEvent === "") {
              calendarData[d][eventCol] = eventText;
            } else {
              // 既存の予定の末尾にカンマがあるかチェック
              if (currentEvent.trim().endsWith(",")) {
                calendarData[d][eventCol] = currentEvent + eventText;
              } else {
                calendarData[d][eventCol] = currentEvent + "," + eventText;
              }
            }
            updatedCells++;
          }
        }
      }
    }
  }
  
  // 更新されたデータをシートに書き込む
  if (updatedCells > 0) {
    sheet.getRange(3, 1, 31, 24).setValues(calendarData);
    Browser.msgBox(updatedCells + "件の" + dayText + "曜日に予定を追加しました。");
  } else {
    Browser.msgBox("追加する" + dayText + "曜日の日付はありませんでした。");
  }
}

/**
 * 祝日データを行事予定から削除する関数
 * シート1のAA列にある祝日データを取得し、カレンダーの行事予定欄からこれらの祝日名を削除します
 */
function removeHolidaysFromSchedule() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var year = sheet.getRange("A1").getValue(); // 西暦を取得
  
  if (!year || isNaN(year)) {
    Browser.msgBox("A1セルに有効な西暦を入力してください。");
    return;
  }
  
  // 祝日データを取得（AA列、AB列、AC列）
  var holidayData = sheet.getRange("AA2:AC1000").getValues();
  var holidays = [];
  
  // 有効な祝日データを配列に格納
  for (var i = 0; i < holidayData.length; i++) {
    if (holidayData[i][0] !== "" && holidayData[i][2] !== "") {
      holidays.push({
        date: holidayData[i][0],  // 日付
        day: holidayData[i][1],   // 曜日
        name: holidayData[i][2]   // 祝日名
      });
    }
  }
  
  // カレンダーデータの範囲（4月から3月まで）
  var calendarRange = [
    {col: 1, eventCol: 2},    // 4月: A列（日付）、B列（行事予定）
    {col: 3, eventCol: 4},    // 5月: C列（日付）、D列（行事予定）
    {col: 5, eventCol: 6},    // 6月
    {col: 7, eventCol: 8},    // 7月
    {col: 9, eventCol: 10},   // 8月
    {col: 11, eventCol: 12},  // 9月
    {col: 13, eventCol: 14},  // 10月
    {col: 15, eventCol: 16},  // 11月
    {col: 17, eventCol: 18},  // 12月
    {col: 19, eventCol: 20},  // 1月
    {col: 21, eventCol: 22},  // 2月
    {col: 23, eventCol: 24}   // 3月
  ];
  
  var calendarData = sheet.getRange(3, 1, 31, 24).getValues(); // カレンダーデータを取得
  var updatedCells = 0;
  var holidayNames = holidays.map(function(h) { return h.name; }); // 祝日名の配列
  
  // カレンダーの各月をループ
  for (var m = 0; m < calendarRange.length; m++) {
    var monthCol = calendarRange[m].col - 1;     // 0ベースのインデックスに変換
    var eventCol = calendarRange[m].eventCol - 1; // 0ベースのインデックスに変換
    
    // 各月の日付をループ
    for (var d = 0; d < 31; d++) {
      var calDate = calendarData[d][monthCol];
      
      // 日付が空でない場合
      if (calDate !== "") {
        var currentEvent = calendarData[d][eventCol] || "";
        
        // 行事予定が空でない場合
        if (currentEvent !== "") {
          var eventItems = currentEvent.split(","); // カンマで区切られた行事予定を配列に
          var newEventItems = [];
          var changed = false;
          
          // 各行事予定項目をループ
          for (var e = 0; e < eventItems.length; e++) {
            var eventItem = eventItems[e].trim();
            var isHoliday = false;
            
            // 祝日名と一致するか確認
            for (var h = 0; h < holidayNames.length; h++) {
              if (eventItem === holidayNames[h]) {
                isHoliday = true;
                changed = true;
                break;
              }
            }
            
            // 祝日でない場合のみ新しい配列に追加
            if (!isHoliday) {
              newEventItems.push(eventItem);
            }
          }
          
          // 変更があった場合のみ更新
          if (changed) {
            calendarData[d][eventCol] = newEventItems.join(",");
            updatedCells++;
          }
        }
      }
    }
  }
  
  // 更新されたデータをシートに書き込む
  if (updatedCells > 0) {
    sheet.getRange(3, 1, 31, 24).setValues(calendarData);
    Browser.msgBox(updatedCells + "件の祝日を行事予定から削除しました。");
  } else {
    Browser.msgBox("削除する祝日はありませんでした。");
  }
}

/**
 * 祝日データを行事予定に追加する関数
 * シート1のAA列にある祝日データを取得し、カレンダーの日付と一致する場合に行事予定欄に祝日名を追加します
 */
function addHolidaysToSchedule() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var year = sheet.getRange("A1").getValue(); // 西暦を取得
  
  if (!year || isNaN(year)) {
    Browser.msgBox("A1セルに有効な西暦を入力してください。");
    return;
  }
  
  // 祝日データを取得（AA列、AB列、AC列）
  var holidayData = sheet.getRange("AA2:AC1000").getValues();
  var holidays = [];
  
  // 有効な祝日データを配列に格納
  for (var i = 0; i < holidayData.length; i++) {
    if (holidayData[i][0] !== "") {
      holidays.push({
        date: holidayData[i][0],  // 日付（例：4/29(水)）
        day: holidayData[i][1],   // 曜日
        name: holidayData[i][2]   // 祝日名
      });
    }
  }
  
  // カレンダーデータの範囲（4月から3月まで）
  var calendarRange = [
    {col: 1, eventCol: 2},    // 4月: A列（日付）、B列（行事予定）
    {col: 3, eventCol: 4},    // 5月: C列（日付）、D列（行事予定）
    {col: 5, eventCol: 6},    // 6月
    {col: 7, eventCol: 8},    // 7月
    {col: 9, eventCol: 10},   // 8月
    {col: 11, eventCol: 12},  // 9月
    {col: 13, eventCol: 14},  // 10月
    {col: 15, eventCol: 16},  // 11月
    {col: 17, eventCol: 18},  // 12月
    {col: 19, eventCol: 20},  // 1月
    {col: 21, eventCol: 22},  // 2月
    {col: 23, eventCol: 24}   // 3月
  ];
  
  var calendarData = sheet.getRange(3, 1, 31, 24).getValues(); // カレンダーデータを取得
  var updatedCells = 0;
  
  // カレンダーの各月をループ
  for (var m = 0; m < calendarRange.length; m++) {
    var monthCol = calendarRange[m].col - 1;     // 0ベースのインデックスに変換
    var eventCol = calendarRange[m].eventCol - 1; // 0ベースのインデックスに変換
    
    // 各月の日付をループ
    for (var d = 0; d < 31; d++) {
      var calDate = calendarData[d][monthCol];
      
      // 日付が空でない場合
      if (calDate !== "") {
        var formattedDate = Utilities.formatDate(calDate, 'Asia/Tokyo', 'M/d');
        var formattedDateWithDay = Utilities.formatDate(calDate, 'Asia/Tokyo', 'M/d') + "(" + getDayOfWeekJP(calDate) + ")";
        
        // 祝日データと比較
        for (var h = 0; h < holidays.length; h++) {
          // 日付オブジェクトを文字列に変換して処理
          var holidayDate = holidays[h].date;
          var holidayDateStr = "";
          
          if (typeof holidayDate === "string") {
            // 既に文字列の場合
            holidayDateStr = holidayDate.split("(")[0]; // 括弧を除いた日付部分を取得
          } else if (holidayDate instanceof Date) {
            // Dateオブジェクトの場合
            holidayDateStr = Utilities.formatDate(holidayDate, 'Asia/Tokyo', 'M/d');
          }
          
          // 日付が一致する場合
          if (holidayDateStr === formattedDate) {
            var currentEvent = calendarData[d][eventCol] || "";
            
            // 既に祝日が入力されていない場合のみ追加
            if (currentEvent.indexOf(holidays[h].name) === -1) {
              if (currentEvent === "") {
                calendarData[d][eventCol] = holidays[h].name;
              } else {
                calendarData[d][eventCol] = currentEvent + "," + holidays[h].name;
              }
              updatedCells++;
            }
          }
        }
      }
    }
  }
  
  // 更新されたデータをシートに書き込む
  if (updatedCells > 0) {
    sheet.getRange(3, 1, 31, 24).setValues(calendarData);
    Browser.msgBox(updatedCells + "件の祝日を行事予定に追加しました。");
  } else {
    Browser.msgBox("追加する祝日はありませんでした。");
  }
}

/**
 * 日付から日本語の曜日を取得する関数
 */
function getDayOfWeekJP(date) {
  var dayOfWeek = date.getDay();
  var days = ["日", "月", "火", "水", "木", "金", "土"];
  return days[dayOfWeek];
}

function zen_han_lower() {
  var zen = "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ：；＜＞";
  var han = zen_han(zen);
  var lower = get_lower(han);
  Logger.log([zen, han, lower]);
}

function zen_han(zen) {
  var han = "";
  var pattern = /[Ａ-Ｚａ-ｚ０-９：；＜＞]/;
  for (var i = 0; i < zen.length; i++) {
    if(pattern.test(zen[i])){
      var letter = String.fromCharCode(zen[i].charCodeAt(0) - 65248);
      han += letter;
    }else{
      han += zen[i];
    }
  }
  return han;
}

function get_lower(han){
  var lower = han.toLowerCase();
  return lower;
}
