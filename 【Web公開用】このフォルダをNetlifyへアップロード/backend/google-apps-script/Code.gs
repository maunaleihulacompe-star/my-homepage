const SPREADSHEET_ID = 'ここにスプレッドシートIDを入力';
const NOTIFY_EMAIL = 'maunaleihulacompe@gmail.com';

const FORM_CONFIG = {
  lei: {sheet:'レイ申込',subject:'レイ・コンテスト申込',fields:[['applicantName','申込者名'],['furigana','フリガナ'],['postalCode','郵便番号'],['address','住所'],['phone','電話番号'],['fax','FAX番号'],['email','メールアドレス'],['workTitle','作品タイトル'],['leiStyle','レイのスタイル'],['deliveryMethod','提出方法'],['materials','主な花材'],['comment','作品コメント'],['agreement','規程への同意']]},
  photo: {sheet:'フォト申込',subject:'フォト・コンテスト申込',fields:[['applicantName','申込者名'],['furigana','フリガナ'],['postalCode','郵便番号'],['address','住所'],['phone','電話番号'],['fax','FAX番号'],['email','メールアドレス'],['workTitle1','作品1タイトル'],['workTitleKana1','作品1タイトルふりがな'],['size1','作品1サイズ'],['deliveryMethod1','作品1提出方法'],['comment1','作品1コメント'],['workTitle2','作品2タイトル'],['workTitleKana2','作品2タイトルふりがな'],['size2','作品2サイズ'],['deliveryMethod2','作品2提出方法'],['comment2','作品2コメント'],['agreement','規程への同意']]},
  vendor: {sheet:'出店申込',subject:'会場出店申込',fields:[['businessName','出店登録名'],['businessKana','フリガナ'],['representative','代表者名'],['representativeKana','代表者フリガナ'],['postalCode','郵便番号'],['address','住所'],['phone','電話番号'],['fax','FAX番号'],['email','メールアドレス'],['vendorType','出店形態'],['products','出店内容・取扱商品'],['notes','備考'],['agreement','案内への同意']]},
  ad: {sheet:'広告申込',subject:'プログラム広告掲載申込',fields:[['companyName','社名'],['companyKana','フリガナ'],['contactName','担当者名'],['contactKana','担当者フリガナ'],['postalCode','郵便番号'],['address','住所'],['phone','電話番号'],['fax','FAX番号'],['email','メールアドレス'],['adSize','広告サイズ'],['notes','備考'],['agreement','案内への同意']]},
  competition: {sheet:'コンペ申込',subject:'コンペ・エキシビション申込',fields:[['harauName','ハラウ名'],['repLastName','代表者氏名（姓）'],['repFirstName','代表者氏名（名）'],['postalCode','郵便番号'],['address','住所'],['phone','電話番号'],['email','メールアドレス'],['soloCategories','Soloエントリー','multi'],['groupCategories','Groupエントリー','multi'],['exhibition','エキシビションエントリー'],['notes','備考'],['agreement','参加規程への同意']]}
};

function getFieldValue(e,item){
  if(item[2]==='multi'){const arr=(e.parameters&&e.parameters[item[0]])||[];return arr.join('、')}
  return e.parameter[item[0]]||'';
}

function doPost(e){
  try{
    const type=e.parameter.formType;
    const config=FORM_CONFIG[type];
    if(!config) throw new Error('不明なフォームです');
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet=ss.getSheetByName(config.sheet)||ss.insertSheet(config.sheet);
    const headers=['受付日時'].concat(config.fields.map(item=>item[1]));
    if(sheet.getLastRow()===0) sheet.appendRow(headers);
    const values=[new Date()].concat(config.fields.map(item=>getFieldValue(e,item)));
    sheet.appendRow(values);
    const lines=config.fields.map(item=>item[1]+'：'+getFieldValue(e,item)).join('\n');
    MailApp.sendEmail({to:NOTIFY_EMAIL,subject:'【Web申込】'+config.subject,body:'Webサイトから申込が届きました。\n\n'+lines});
    const reply=e.parameter.email;
    let noteText='';
    if(type==='lei') noteText='\n\n【ご注意】\n作品と【作品の説明文】は、2026年11月27日（金）午前中までに事務局へお届けください（生花のため、この日より前はお預かりできません）。\n【作品の説明文】は、大会公式サイトの参加規程PDFを印刷し、必要事項をご記入のうえご用意ください。';
    if(reply) MailApp.sendEmail({to:reply,subject:'【Maunalei】'+config.subject+'を受け付けました',body:'お申し込みありがとうございます。\n以下の内容で受け付けました。\n\n'+lines+noteText+'\n\nマウナレイフラコンペティション事務局'});
    return ContentService.createTextOutput('OK');
  }catch(error){
    MailApp.sendEmail(NOTIFY_EMAIL,'【Maunalei】フォーム送信エラー',String(error));
    return ContentService.createTextOutput('ERROR');
  }
}
