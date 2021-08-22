//*** メイン処理 ***//
function submitReception(e) {
  //フォーム１
  var mielForm = FormApp.openById('********************************************');
  var mielApplication = SpreadsheetApp.openById('********************************************');

  //回答内容
  var lastRow = mielApplication.getSheetByName('フォーム(1)').getLastRow();
  var formResponses = mielApplication.getSheetByName('フォーム(1)').getRange(2,1,lastRow-1,8).getValues();
  var newResponse = e.response;
  console.log("formResponses："+formResponses);

  //ミエルのメールアドレス
  var mielEMail ='**************@**************'

  //品数チェック
  if (checkCount(mielForm,mielApplication,formResponses)){
    //売り切れなら受付終了
    mielForm.setAcceptingResponses(false);
    console.log("フォームを受付終了にします");
  }

  //申込受付後の通知
  sendMailtoMiel(mielForm,mielApplication,newResponse,mielEMail);
}


//*** 品数チェック ***//
// Description 現在の申込数をカウントし、最大申込数に達した場合は受付を終了します。
function checkCount(mielForm,mielApplication,formResponses){
  //クローズフラグ
  var isClosed = false;
  //最大申込数を確認
  var max = mielApplication.getSheetByName('【管理】').getRange("C4").getValue();
  console.log('Spreadsheet.注文可能数量=' + max);

  //現在の申込数を確認
  var count = 0;
  formResponses.forEach(function(response){
    count += Number(response[4]);

    if(count >=max) {
      // 売切なら
      isClosed = true;
    }
  });

  //管理シートの申込数を更新
  updateInformation(mielForm, mielApplication, count, isClosed);

  return isClosed;
}

//*** 管理シート更新処理 ***//
function updateInformation(mielForm, mielApplication, count, isClosed){
  //管理シートを更新する
  mielApplication.getSheetByName('【管理】').getRange("D4").setValue(count);
  mielApplication.getSheetByName('【管理】').getRange("E4").setValue(Utilities.formatDate(new Date(),'Asia/Tokyo', 'yyyy年MM月dd日 hh:mm'));
  if (isClosed) {
    mielApplication.getSheetByName('【管理】').getRange("F4").setValue("売切");
  }
}

//*** 申込受付後の通知 ***//
function sendMailtoMiel(mielForm,mielApplication,newResponse,mielEMail){
  //メール送信情報
  var formTitle = mielApplication.getSheetByName('【管理】').getRange("A4").getValue();
  var date = mielApplication.getSheetByName('【管理】').getRange("B4").getValue();
  var max = mielApplication.getSheetByName('【管理】').getRange("C4").getValue();
  var current = mielApplication.getSheetByName('【管理】').getRange("D4").getValue();
  var nowDate = mielApplication.getSheetByName('【管理】').getRange("E4").getValue();
  var isClosed = mielApplication.getSheetByName('【管理】').getRange("F4").getValue();
  var mailAddress = mielApplication.getSheetByName('【管理】').getRange("H4").getValue();
  var zan = Number(max)-Number(current);

  var reception = "受付時刻："+Utilities.formatDate(newResponse.getTimestamp(),'Asia/Tokyo', 'yyyy年MM月dd日 hh:mm')+"\n";
  var name = "お名前："+newResponse.getItemResponses()[1].getResponse()+"さん\n";
  var count = "個数："+newResponse.getItemResponses()[2].getResponse()+"個\n";
  var pay = "支払い方法："+newResponse.getItemResponses()[3].getResponse()+"\n";
  var hour ="受け取り時間："+newResponse.getItemResponses()[4].getResponse()+"\n";
  var remarks ="備考："+newResponse.getItemResponses()[5].getResponse()+"\n";

  //メール送信
  var mailTitle = date + ' '+ formTitle;
  var mailMessage = date + '受取分 「'+ formTitle +'」の注文を受け付けました。\n'
                    +'現在: '+ current +'食受付済み \n残り' + zan +'食\n\n------------\n';
  var mailMessage2 = reception + name + count + pay + hour + remarks;
  
  //売切れだったら
  if (isClosed == '') {
    mailTitle = '【受付完了】'+ mailTitle;
  } else {
    mailTitle = '【完売です！】'+ mailTitle;
  }

  //残数がマイナスになっちゃったら
  if (zan < 0) {
    mailMessage = mailMessage + '（足りなくなっちゃいました！調整お願いします。）'
  }

  console.log('Mail=' + mailTitle);
  console.log(mailMessage + mailMessage2);

  //管理者へメール通知
  GmailApp.sendEmail(mailAddress, mailTitle, mailMessage + mailMessage2);

  //ミエルへメール通知
  GmailApp.sendEmail(mielEMail, mailTitle, mailMessage + mailMessage2);
}


