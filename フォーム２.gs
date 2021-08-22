//*** 申込開始時の処理 ***//
// Description フォームを受付中にし、過去の受付データをリセットします
function startReception2() {
  //フォーム２
  var mielForm = FormApp.openById('********************************************');
  var mielApplication = SpreadsheetApp.openById('********************************************');

    console.log("Debug:startReception2を開始しました");

  //確認メッセージ表示
  if(confirmMessage2(mielApplication)){
    console.log("申込開始処理をはじめます（フォーム２）");
  } else {
    console.log("申込開始処理はキャンセルされました（フォーム２）");
    return
  }

  //管理シートをリセットする
  updateInformation2(mielApplication);

  //フォームの注文内容を更新する
  updateReceptionForm2(mielForm,mielApplication);

  //フォームを受付中にする
  mielForm.setAcceptingResponses(true);
  console.log("受付開始します（フォーム２）");

}

//*** 申込み開始：確認メッセージ ***//
function confirmMessage2(mielApplication){
  //表示メッセージ作成
  var formTitle = mielApplication.getSheetByName('【管理】').getRange("A5").getValue();
  var date = mielApplication.getSheetByName('【管理】').getRange("B5").getValue();
  var message = date + '受取分 「'+ formTitle +'」の注文受付を始めます。前回の受付情報は削除されますがよろしいですか？';

  var confirm = Browser.msgBox("確認してください",message,Browser.Buttons.OK_CANCEL);
  if (confirm == 'ok') {
    return true;
  } else if (confirm == 'cancel') {
    return false;
  }
}

//*** 申込み開始：管理シート更新 ***//
function updateInformation2(mielApplication){
  //管理情報を空欄にする
  mielApplication.getSheetByName('【管理】').getRange("D5").setValue("");
  mielApplication.getSheetByName('【管理】').getRange("E5").setValue("");
  mielApplication.getSheetByName('【管理】').getRange("F5").setValue("");

  //申込み情報をリセットする
  var lastRow = mielApplication.getSheetByName('フォーム(2)').getLastRow();
  mielApplication.getSheetByName('フォーム(2)').deleteRows(2,lastRow);
}

//*** 申込み開始：フォームの注文内容を更新 ***//
function updateReceptionForm2(mielForm,mielApplication){
  //受付メニュー名を取得
  var menu = mielApplication.getSheetByName('【管理】').getRange("A5").getValue();
  var date = mielApplication.getSheetByName('【管理】').getRange("B5").getValue();
  var newMenu = [date + " " + menu];

  //受付メニュー名をフォームにセット
  var formItems = mielForm.getItems();
  formItems[0].asMultipleChoiceItem().setChoiceValues(newMenu).setRequired(true);
}
