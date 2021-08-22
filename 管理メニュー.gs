//管理シートを開いたらメニュー表示
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('管理ツール');
  menu.addItem('フォーム１受付開始', 'startReception1');
  menu.addItem('フォーム２受付開始', 'startReception2');
  menu.addItem('フォーム３受付開始', 'startReception3');
  menu.addToUi();
}

