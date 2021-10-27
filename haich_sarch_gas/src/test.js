function myFunction() {
  var value = PropertiesService.getScriptProperties().getProperty("date");
  Logger.log(value);
}
