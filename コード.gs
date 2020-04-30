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
const SUMMARY_DATE_COL = 1;

function main() {
  console.time('total');
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
    let today = lastCell.getValue();
    
    // 1件もデータがなければ終わり
    if(today=='日付'){
      return;
    }
    
    // 上がりながら、日付の切れ目を探す    
    let checkDates = recordSheet.getRange(1,RECORD_DATE_COL,lastRow,1).getValues();
    let todayTime = today.getTime(),j;
    for(j=checkDates.length-1;j>=0;j--){
      if(checkDates[j][0]=='日付' || checkDates[j][0].getTime() != todayTime){
        break;
      }
    }
    let firstRow = j + 2;
    
    // 下がりながら集計する    
    // 走査する範囲を取得する
    let times = recordSheet.getRange(firstRow,RECORD_TIME_COL,lastRow-firstRow+1,1).getValues();
    let actions = recordSheet.getRange(firstRow,RECORD_ACTION_COL,lastRow-firstRow+1,1).getValues();
    // 不要範囲をクリアする
    recordSheet.getRange(firstRow,RECORD_WORKTIME_COL,lastRow-firstRow+1,1).clearContent();
    recordSheet.getRange(firstRow,RECORD_MEMO_COL,lastRow-firstRow+1,1).clearContent();
    
    let startRow=firstRow,startTime,startAction,stopRow,stopTime,stopAction;
    let workTime,workTimeSum=0,restTimeSum;
    let dayWorkStartTime=null,dayWorkStopTime=null;
    let findStart=false,findStop=false;
    let lastAction = '';
    
    while(1){      
      // 仕事を始めるアクションを探す
      let findStart = false;
      while(startRow <= lastRow){
        startTime = times[startRow-firstRow][0].getTime();
        startAction = actions[startRow-firstRow][0];
        
        if(startAction == '開始' || startAction =='再開'){
          findStart = true;
          // 一番最初に見つかった開始時刻を、その日の開始時刻にする。
          if(!dayWorkStartTime){
            dayWorkStartTime = times[startRow-firstRow][0];
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
        stopTime = times[stopRow-firstRow][0].getTime();
        stopAction = actions[stopRow-firstRow][0];
        
        if(stopAction == '中断' || stopAction =='終了'){
          findStop = true;
          dayWorkStopTime = times[stopRow-firstRow][0];
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
    
    
    // ------------ 日付別集計シート整備 ------------
    let summaryLastCell = summarySheet.getRange(summarySheet.getMaxRows(), SUMMARY_DATE_COL).getNextDataCell(SpreadsheetApp.Direction.UP);
    let summaryRow = summaryLastCell.getRow()+1; 
    let summaryLastCellValue = summaryLastCell.getValue();
    let summaryLastCellRow = summaryLastCell.getRow();
    if(summaryLastCellValue=='日付'){
      summaryRow = summaryLastCellRow+1;
    }else if(summaryLastCellValue.getTime()==today.getTime()){
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
    console.error(printError(error));
  }finally{    
  }
  console.timeEnd('total');
}

function printError(error){
  return "[メッセージ]" + error.message + "\n" + "[StackTrace]\n" + error.stack;
}

function msToTime(duration) {
  return (new Date(duration)).toUTCString().match(/(\d\d:\d\d):\d\d/)[1];
}

