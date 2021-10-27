//メイン
function taisyokuMain() {
	//開始前にポップアップを出す
	let okbutton = startAlart();
	if (okbutton == 0) return; //キャンセルを押されたら終了する

	//bookをactiveにする
	let book = SpreadsheetApp.getActiveSpreadsheet();
	/*book内すべてのシートを取得
	0.anser(フォーム回答を貯める場所)
	1.enrolledFlag(在籍/退職名簿)
	2.toExcel(Excelに流すデータ)*/
	let sheets = book.getSheets();
	//toExcelの講師番号の列だけ取得(A2からデータ取得)
	let exceldata = getData(sheets[2], 1);
	//enrolledFlag全体を取得(検索対象)
	let retiredTutors = getData(sheets[1], 3);
	//flagdataから退職済みの講師番号だけを配列として取得(検索対象)
	let retiredTutor = [];
	for (let i = 0; i < retiredTutors.length; i++) {
		//退職済みなら在籍フラグは0に設定している
		if (retiredTutors[i][2] == 0) {
			//退職済みなら1次元配列の末尾に追加
			retiredTutor.push(retiredTutors[i][0]);
		}
	}
	//検索と削除
	eraseTutor(sheets[2], sarchIndex(exceldata, retiredTutor));
	
	//処理後に終了のダイアログを表示する
	endAlart();
}

//シートから必要な列だけ抜き出す関数
function getData(sheet, row) {
	const len = sheet.getLastRow();
	return sheet.getRange(2, 1, len - 1, row).getValues();
}

//toExcel中の退職済み講師のindexを検索する関数
function sarchIndex(exceldata, retiredTutor) {
	//削除対象を格納する一次元配列
	let deleteTutor = [];
	//retiredTutorを参照して、退職ならばtoExcelから該当講師の行を削除する
	for (let i = 0; i < exceldata.length; i++) {
		let index = retiredTutor.indexOf(exceldata[i][0]);
		if (index != -1) {
			deleteTutor.push(i + 2);  //indexと行数は2ずれるため補正
		}
	}
	return deleteTutor;
}

//sarchIndexで得られたindexを削除する
function eraseTutor(sheet, deleteTutor) {
	//削除対象のindexの配列を降順にソートする
	deleteTutor.sort((a, b) => { return b - a; });
	//indexが大きいものから削除していく(deleteRowは1行削除すると上詰めされるため下から削除していく)
	for (let i = 0; i < deleteTutor.length; i++) {
		sheet.deleteRows(deleteTutor[i]);
	}
}

//開始前のポップアップを出す
function startAlart() {
  let ui = SpreadsheetApp.getUi();
  let result = ui.alert("リストから退職済みの講師を削除します", ui.ButtonSet.OK_CANCEL);
	if (result == ui.Button.OK) { return 1; }
	else { return 0; }
}

//処理後のポップアップを出す
function endAlart(){
	Browser.msgBox("リストから退職者を削除しました")
}