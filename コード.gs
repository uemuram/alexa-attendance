// 定数(座標)
const RECORD_DATE_COL = 1;
const RECORD_TIME_COL = 2;
const RECORD_ACTION_COL = 3;
const RECORD_WORKTIME_COL = 5;
const RECORD_MEMO_COL = 6;
const STATUS_STATUS_ROW = 3;
const STATUS_STATUS_COL = 1;
const STATUS_WORKTIME_ROW = 5;
const STATUS_WORKTIME_COL = 3;
const STATUS_RESTTIME_ROW = 6;
const STATUS_RESTTIME_COL = 3;

function main() {
  console.log('start');
  
  // シートを取得する
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let recordSheet = spreadsheet.getSheetByName("作業記録");
  let statusSheet = spreadsheet.getSheetByName("ステータス");
  let summarySheet = spreadsheet.getSheetByName("日付別集計");
  
  // ログ用のセル
  let logCell1 = recordSheet.getRange(20, 7);
  let logCell2 = recordSheet.getRange(21, 7);
  let logCell3 = recordSheet.getRange(22, 7);
  let now = new Date(); //現在日時を取得
  let time = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  
  let logStr = "";

  try{
    // 一番下の作業記録をとる
    var lastCell = recordSheet.getRange(recordSheet.getMaxRows(), RECORD_DATE_COL).getNextDataCell(SpreadsheetApp.Direction.UP);
    let lastRow = lastCell.getRow();
    logStr += (lastRow + ":" + lastCell.getColumn() + "\n");
    // 1件もデータがなければ終わり
    if(recordSheet.getRange(lastRow,RECORD_DATE_COL).getValue()=='日付'){
      return;
    }
    
    // 上がりながら、日付の切れ目を探す
    let i;
    let prevCellValue;
    let currentCellValue = lastCell.getValue();
    
    for (i=lastRow; i>0; i--){
      prevCellValue = recordSheet.getRange(i-1,RECORD_DATE_COL).getValue();
      if(prevCellValue=='日付' || currentCellValue.getTime() != prevCellValue.getTime()){
        break;
      }
      currentCellValue = prevCellValue;
    }
    let firstRow = i;
    logStr += ("-" + firstRow + "\n");
    
    // 下がりながら集計する
    let startRow=firstRow,startTime,startAction,stopRow,stopTime,stopAction;
    let workTime,workTimeSum=0,restTimeSum;
    let dayWorkStartTime=null,dayWorkStopTime=null;
    let findStart=false,findStop=false;
    let lastAction = '';
    while(1){      
      // 仕事を始めるアクションを探す
      let findStart = false;
      while(startRow <= lastRow){
        startTime = recordSheet.getRange(startRow,RECORD_TIME_COL).getValue().getTime();
        startAction = recordSheet.getRange(startRow,RECORD_ACTION_COL).getValue();
        recordSheet.getRange(startRow,RECORD_WORKTIME_COL).setValue("");
        recordSheet.getRange(startRow,RECORD_MEMO_COL).setValue("");
        if(startAction == '開始' || startAction =='再開'){
          findStart = true;
          // 一番最初に見つかった開始時刻を、その日の開始時刻にする。
          if(!dayWorkStartTime){
            dayWorkStartTime = recordSheet.getRange(startRow,RECORD_TIME_COL).getDisplayValue();
          }
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
        stopTime = recordSheet.getRange(stopRow,RECORD_TIME_COL).getValue().getTime();
        stopAction = recordSheet.getRange(stopRow,RECORD_ACTION_COL).getValue();
        recordSheet.getRange(stopRow,RECORD_WORKTIME_COL).setValue("");
        recordSheet.getRange(stopRow,RECORD_MEMO_COL).setValue("");
        if(stopAction == '中断' || stopAction =='終了'){
          findStop = true;
          dayWorkStopTime = recordSheet.getRange(stopRow,RECORD_TIME_COL).getDisplayValue();
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
    
    
    
    
    logStr += dayWorkStartTime + "\n";
    logStr += dayWorkStopTime + "\n";
    
    logStr += msToTime(workTimeSum);

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

    
    // ------------ 日付別集計シート整備 ------------
    
    
    
    
    
    
    
    
  }catch(error){
    logCell3.setValue(printError(error));
  }finally{
    logCell1.setValue(time);
    logCell2.setValue(logStr);
    
  }   
}

function printError(error){
  return "[名前] " + error.name + "\n" +
    "[場所] " + error.fileName + "(" + error.lineNumber + "行目)\n" +
      "[メッセージ]" + error.message + "\n" +      
        "[StackTrace]\n" + error.stack;
}

function msToTime(duration) {
  return (new Date(duration)).toUTCString().match(/(\d\d:\d\d):\d\d/)[1];
}
