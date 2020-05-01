// 定数(座標)
const RECORD_DATE_COL = 1;
const RECORD_ACTION_COL = 2;
const RECORD_WORKTIME_COL = 4;
const RECORD_MEMO_COL = 5;
const STATUS_STATUS_ROW = 3;
const STATUS_STATUS_COL = 1;
const STATUS_WORKTIME_ROW = 7;
const STATUS_WORKTIME_COL = 3;
const STATUS_PASTTIME_ROW = 4;
const STATUS_PASTTIME_COL = 2;
const SUMMARY_DATE_COL = 1;
// 定数(月)
const MONTHS = {
  "January"    : 0,
  "February"   : 1,
  "March"      : 2,
  "April"      : 3,
  "May"        : 4,
  "June"       : 5,
  "July"       : 6,
  "August"     : 7,
  "September"  : 8,
  "October"    : 9,
  "November"   : 10,
  "December"   : 11,
};

function main() {
  console.log('start');
  
  // シートを取得する
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let recordSheet = spreadsheet.getSheetByName("作業記録");
  let statusSheet = spreadsheet.getSheetByName("ステータス");
  let summarySheet = spreadsheet.getSheetByName("日付別集計");
  
  try{
    // 一番下の作業記録をとる
    var lastCell = recordSheet.getRange(recordSheet.getMaxRows(), RECORD_DATE_COL).getNextDataCell(SpreadsheetApp.Direction.UP);
    let lastRow = lastCell.getRow();
    let todayStr = lastCell.getValue();
    
    // 1件もデータがない、もしくは日付情報がとれなかったら終わり
    let today = iftttDateStr2Date(todayStr);
    if(!today){
      return;
    }
    
    // 上がりながら、日付の切れ目を探す
    let checkDates = recordSheet.getRange(1,RECORD_DATE_COL,lastRow,1).getValues();
    let dates = [iftttDateStr2Date(todayStr)];    
    for(j=checkDates.length-1;j>=1;j--){
      console.log(j + ":" + checkDates[j][0]);
      let prevDate = iftttDateStr2Date(checkDates[j-1][0]);
      if(!prevDate || !isSameDate(today, prevDate)){
        break;
      }
      dates.unshift(prevDate);
    }
    let firstRow = j + 1;
    
    // 下がりながら集計する    
    // 走査する範囲を取得
    let actions = recordSheet.getRange(firstRow,RECORD_ACTION_COL,lastRow-firstRow+1,1).getValues();
    // 不要範囲をクリア
    recordSheet.getRange(firstRow,RECORD_WORKTIME_COL,lastRow-firstRow+1,1).clearContent();
    recordSheet.getRange(firstRow,RECORD_MEMO_COL,lastRow-firstRow+1,1).clearContent();

    let startRow=firstRow,startTime,startAction,stopRow,stopTime,stopAction;
    let workTime,workTimeSum=0,restTimeSum;
    let dayWorkStartTime=null,dayWorkStopTime=null,currentWorkStartTime=null;
    let findStart=false,findStop=false;
    let lastAction = '';
    
    while(1){      
      // 仕事を始めるアクションを探す
      let findStart = false;
      while(startRow <= lastRow){
        startTime = dates[startRow-firstRow].getTime();
        startAction = actions[startRow-firstRow][0];
        
        if(startAction == '開始' || startAction =='再開'){
          findStart = true;
          // 一番最初に見つかった開始時刻を、その日の開始時刻にする。
          if(!dayWorkStartTime){
            dayWorkStartTime = dates[startRow-firstRow];
          }
          currentWorkStartTime = dates[startRow-firstRow];
          lastAction = startAction;
          break;
        }
        recordSheet.getRange(startRow,RECORD_MEMO_COL).setValue("(skip)");
        startRow++;
      }      
      // 仕事を始めるアクションが見つからなければ終わり
      if(!findStart){
        break;
      }
      
      // 仕事を止めるアクションを探す
      stopRow = startRow + 1;
      let findStop = false;
      while(stopRow <= lastRow){
        stopTime = dates[stopRow-firstRow].getTime();
        stopAction = actions[stopRow-firstRow][0];
        
        if(stopAction == '中断' || stopAction =='終了'){
          findStop = true;
          dayWorkStopTime = dates[stopRow-firstRow];
          lastAction = stopAction;
          break;
        }
        recordSheet.getRange(stopRow,RECORD_MEMO_COL).setValue("(skip)");
        stopRow++;  
      }      
      // 仕事を止めるアクションが見つからなければ終わり
      if(!findStop){
        break;
      }
      // 仕事時間を計算
      workTime = stopTime - startTime;
      recordSheet.getRange(stopRow,RECORD_WORKTIME_COL).setValue(msToTime(workTime));
      workTimeSum += workTime;
      // 次の開始行をセット
      startRow = stopRow + 1;
    }
    
    
    // ------------ スタータスシート整備 ------------
    // ステータ作成
    let status = 'オフ';
    if(lastAction == '開始' || lastAction == '再開'){
      status = '仕事中';
    }else if(lastAction == '中断'){
      status = '休憩中';
    }
    if(recordSheet.getRange(lastRow,RECORD_ACTION_COL).getValue()=='終了'){
      status = 'オフ';
    }
    // シートに反映
    statusSheet.getRange(STATUS_STATUS_ROW,STATUS_STATUS_COL).setValue(status);
    statusSheet.getRange(STATUS_WORKTIME_ROW,STATUS_WORKTIME_COL).setValue(msToTime(workTimeSum));
    
    let pastTime = '';
    if(status == '仕事中'){
      pastTime = '(' + currentWorkStartTime.toString().match(/(\d\d:\d\d):\d\d/)[1] + '～)';
    }else if(status == '休憩中'){
      pastTime = '(' + dayWorkStopTime.toString().match(/(\d\d:\d\d):\d\d/)[1] + '～)';
    }
    statusSheet.getRange(STATUS_PASTTIME_ROW,STATUS_PASTTIME_COL).setValue(pastTime);
    
    
    // ------------ 日付別集計シート整備 ------------
    let summaryLastCell = summarySheet.getRange(summarySheet.getMaxRows(), SUMMARY_DATE_COL).getNextDataCell(SpreadsheetApp.Direction.UP);
    let summaryRow = summaryLastCell.getRow()+1; 
    let summaryLastCellValue = summaryLastCell.getValue();
    let summaryLastCellRow = summaryLastCell.getRow();
    if(summaryLastCellValue=='日付'){
      summaryRow = summaryLastCellRow+1;
    }else if(isSameDate(summaryLastCellValue,today)){
      summaryRow = summaryLastCellRow;      
    }
    // 各種情報を表示
    // 日付け、曜日
    summarySheet.getRange(summaryRow,SUMMARY_DATE_COL).setValue(today);
    summarySheet.getRange(summaryRow,SUMMARY_DATE_COL+1).setFormulaR1C1('=TEXT(R[0]C[-1],"ddd")');
    // 開始、終了時刻、作業時間
    summarySheet.getRange(summaryRow,SUMMARY_DATE_COL+2,1,3).setValues([[dayWorkStartTime,dayWorkStopTime,msToTime(workTimeSum)]]);
    // 休憩時間
    summarySheet.getRange(summaryRow,SUMMARY_DATE_COL+5).setFormulaR1C1('=R[0]C[-2]-R[0]C[-3]-R[0]C[-1]');    

    

  }catch(error){
    recordSheet.getRange(10,8).setValue(printError(error));
    console.error(printError(error));
  }finally{    
  }
}

function printError(error){
  return "[メッセージ]" + error.message + "\n" + "[StackTrace]\n" + error.stack;
}

// ミリ秒を時間に変換
function msToTime(duration) {
  return (new Date(duration)).toUTCString().match(/(\d\d:\d\d):\d\d/)[1];
}

// IFTTTの日付時刻形式から、日付時刻を取り出す
// April 28, 2020 at 02:41PM → 4/28 14:41
function iftttDateStr2Date(str){
  if(!str){
    return null;
  }
  let match = str.match(/([a-zA-Z]+) (\d+), ([\d]{4}) at (\d+):(\d+)(AM|PM)/);
  if(match){
    let year = match[3];
    let month = MONTHS[match[1]];
    let day = match[2];
    let ampm = match[6];
    let hour = (ampm == 'PM' && match[4] <= 11) ? Number(match[4]) + 12 : match[4];
    let minutes = match[5];
    return new Date(year,month,day,hour,minutes);
  } else {
    return null;
  }
}

// 日付けオブジェクト同士で、日付部分が一致するかを確認する
function isSameDate(date1,date2){
  if(date1.getYear()==date2.getYear() &&
     date1.getMonth()==date2.getMonth() &&
     date1.getDay()==date2.getDay()){
    return true;
  } else {
    return false;
  }
}