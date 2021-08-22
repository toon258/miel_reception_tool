//*** 申込開始時の処理 ***//
// Description フォームを受付中にし、過去の受付データをリセットします
function startReception3() {
  //フォーム３
  var mielForm = FormApp.openById('********************************************');
  var mielApplication = SpreadsheetApp.openById('********************************************');

    console.log("Debug:startReception3を開始しました");

  //確認メッセージ表示
  if(confirmMessage3(mielApplication)){
    console.log("申込開始処理をはじめます（フォーム３）");
  } else {
    console.log("申込開始処理はキャンセルされました（フォーム３）");
    return
  }

  //管理シートをリセットする
  updateInformation3(mielApplication);

  //フォームの注文内容を更新する
  updateReceptionForm3(mielForm,mielApplication);

  //フォームを受付中にする
  mielForm.setAcceptingResponses(true);
  console.log("受付開始します（フォーム３）");

}

//*** 申込み開始：確認メッセージ ***//
function confirmMessage3(mielApplication){
  //表示メッセージ作成
  var formTitle = mielApplication.getSheetByName('【管理】').getRange("A6").getValue();
  var date = mielApplication.getSheetByName('【管理】').getRange("B6").getValue();
  var message = date + '受取分 「'+ formTitle +'」の注文受付を始めます。前回の受付情報は削除されますがよろしいですか？';

  var confirm = Browser.msgBox("確認してください",message,Browser.Buttons.OK_CANCEL);
  if (confirm == 'ok') {
    return true;
  } else if (confirm == 'cancel') {
    return false;
  }
}

//*** 申込み開始：管理シート更新 ***//
function updateInformation3(mielApplication){
  //管理情報を空欄にする
  mielApplication.getSheetByName('【管理】').getRange("D6").setValue("");
  mielApplication.getSheetByName('【管理】').getRange("E6").setValue("");
  mielApplication.getSheetByName('【管理】').getRange("F6").setValue("");

  //申込み情報をリセットする
  var lastRow = mielApplication.getSheetByName('フォーム(3)').getLastRow();
  mielApplication.getSheetByName('フォーム(3)').deleteRows(2,lastRow);
}

//*** 申込み開始：フォームの注文内容を更新 ***//
function updateReceptionForm3(mielForm,mielApplication){
  //受付メニュー名を取得
  var menu = mielApplication.getSheetByName('【管理】').getRange("A6").getValue();
  var date = mielApplication.getSheetByName('【管理】').getRange("B6").getValue();
  var newMenu = [date + " " + menu];

  //受付メニュー名をフォームにセット
  var formItems = mielForm.getItems();
  formItems[0].asMultipleChoiceItem().setChoiceValues(newMenu).setRequired(true);
}
