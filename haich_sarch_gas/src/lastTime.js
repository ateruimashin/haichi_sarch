function getLastTime() {
  //現在の日時を取得
  var date = Utilities.formatDate(new Date(),"Asia/Tokyo","yyyy/MM/dd HH:mm:ss");
  PropertiesService.getScriptProperties().setProperty("date",date);
  var value = PropertiesService.getScriptProperties().getProperty("date");
  return value;
}
