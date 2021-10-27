function updateTutorsDataMain() {
  //最終更新日時を取得(ミリ秒で)
  const lastDate = getLastDate();
  //全シートを取得
  const sheets = getAllsheets();
  //answerシートの列数を取得
  const answerLen = sheets[0].getLastColumn();
  //シートのデータを配列に取得
  //1列しか取得しない場合はflatメソッドで一次元配列にする
  const timeStamps = getData(sheets[0], 1, 2);  //タイムスタンプと講師番号のみ
  const answer = getData(sheets[0], 2, answerLen - 1); //タイムスタンプを含まない
  const tutorNumList = getData(sheets[1], 1, 1).flat(); //講師リスト
  const toExcelList = getData(sheets[2], 1, 1).flat(); //toExcelの講師番号リスト
  //update対象の配列を取得
  let updateData = getUpdatelist(timeStamps, lastDate);
  //アップデート対象の新人/在籍それぞれの配列を取得
  let [newTutors, oldTutors] = checkTutor(tutorNumList, updateData);
  //対象が新人講師の場合の処理
  addNewTutors(newTutors, answer, sheets);
  //在籍講師を更新する場合の処理
  updateOldTutorData(oldTutors, answer, toExcelList, sheets[2])
  //最終更新日時をアップデート
  updateLastDate();
}

//最終更新日時を取得する(戻り値はUNIX時間)
function getLastDate() {
  //PropertyServiceから最終更新日時を取得(日付)
  const propertyDate = PropertiesService.getScriptProperties().getProperty("date");
  let lastDate;
  if (propertyDate == null) {
    //初回起動時などはpropertyDateはnullになるため、その場合は0に設定する
    lastDate = 0;
  } else {
    //それ以外は日付をUNIX時間に変換する
    lastDate = Date.parse(propertyDate);
  }
  return lastDate;
}

//最終更新日時をアップデートする
function updateLastDate() {
  let date = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  PropertiesService.getScriptProperties().setProperty("date", date);
}

//スプレッドシート上のすべてのシートを取得する
function getAllsheets() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = book.getSheets();
  return sheets;
}

//シートから指定の範囲を抜き取り配列に格納する
function getData(sheet, startCol, col) {
  const rowLen = sheet.getLastRow();
  if (rowLen <= 1) {
    return [0];
  } else {
    return sheet.getRange(2, startCol, rowLen - 1, col).getValues();
  }
}

//アップデート対象になる行と講師番号を取得
function getUpdatelist(timeStamps, lastDate) {
  //最終更新日時より後のデータのみ抽出
  //格納するのはanswerシートの行番号,講師番号
  let updateData = [];
  timeStamps.forEach((element, index) => {
    //タイムスタンプをミリ秒に変換
    let timeStamp = Date.parse(element[0]);
    //最終更新日時より新しいならばupdateDataに追加
    if (timeStamp >= lastDate) {
      updateData.push([index, element[1]]);
    }
  });
  return updateData;
}

//在職リストと照らし合わせる
function checkTutor(tutorNumList, updateData) {
  //二分探索するためにsortしておく
  //tutorListに存在するかのみ判定するのでindexは関係ない
  tutorNumList.sort();
  //新人講師と在籍講師それぞれの配列を作成
  let newTutors = []; //新人講師
  let oldTutors = []; //在籍講師
  //新人講師か在籍講師か判別し、処理を分ける
  updateData.forEach(element => {
    let index = binarySarch(element[1], tutorNumList);
    if (index == -1) {
      //新人講師の場合
      newTutors.push(element);
    } else {
      //在籍講師の場合
      oldTutors.push(element);
    }
  });
  //配列で2つの配列を戻り値にする
  return [newTutors, oldTutors];
}

//新人講師をtutorListとtoExcelに追加
function addNewTutors(newTutors, answer, sheets) {
  //新人がいないなら処理しない
  if (newTutors.length == 0) return;
  //tutorListに追加
  addTotorList(newTutors, sheets[1]);
  //toExcelに追加
  addNewtoExcel(newTutors, answer, sheets[2]);
}

//講師をtotorListに追加する
function addTotorList(tutors, sheet) {
  //挿入するデータを作る
  let flagData = [];
  tutors.forEach(element => {
    //講師番号,在籍フラグ(1に設定)を末尾に追加
    flagData.push([element[1], 1]);
  });
  //挿入範囲を決める
  const rowLen = sheet.getLastRow();
  const rangeLen = flagData.length;
  //データを挿入する
  sheet.getRange(rowLen + 1, 1, rangeLen, 2).setValues(flagData);
}

//新人講師をtoExcelに追加
function addNewtoExcel(tutors, answer, sheet) {
  //挿入するデータを作る
  let tutorData = [];
  tutors.forEach(element => {
    tutorData.push(answer[element[0]]);
  });
  //挿入範囲を決める
  const rowLen = sheet.getLastRow();
  const colLen = sheet.getLastColumn();
  const rangeLen = tutorData.length;
  //データを挿入する
  sheet.getRange(rowLen + 1, 1, rangeLen, colLen).setValues(tutorData);
}

//在籍講師のデータを更新(toExcel)
function updateOldTutorData(tutors, answer, toExcelList, sheet) {
  //更新する講師がいないなら処理しない
  if (tutors.length == 0) return;
  //最終列を取得
  tutors.forEach(tutor => {
    //toExcelListから該当の講師番号を検索
    const index = toExcelList.indexOf(tutor[1]);
    //置き換え後のデータ
    const replaceData = [answer[tutor[0]]];
    //indexに相当する行を書き換え
    sheet.getRange(index + 2, 1, 1, replaceData[0].length).setValues(replaceData);
  });
}

//二分探索
function binarySarch(tutuorNum, tutorNumList) {
  //戻り値の初期値(-1はエラー)
  let index = -1;
  let left = 0, right = tutorNumList.length - 1;
  while (left <= right) {
    let mid = Math.floor((left + right) / 2);
    if (tutorNumList[mid] == tutuorNum) {
      index = mid;
      break;
    } else if (tutorNumList[mid] < tutuorNum) {
      left = mid + 1;
    } else {
      right = mid - 1;
    }
  }
  return index;
}