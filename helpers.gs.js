// Shows an alert normally, but collects messages instead when running "all at once"
let RUN_LOG = null;
function notify(message) {
  if (RUN_LOG) RUN_LOG.push(message);
  else SpreadsheetApp.getUi().alert(message);
}