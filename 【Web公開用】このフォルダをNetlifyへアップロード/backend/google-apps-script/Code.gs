const SPREADSHEET_ID = '1DJ2qpSGcCKZ-EZ94STMpYAJPFQg2gR1j0v0TSKbPHts';
const NOTIFY_EMAIL = 'maunaleihulacompe@gmail.com';
const INVOICE_FOLDER_NAME = 'Maunalei2026_請求書';
const BANK_INFO = 'ゆうちょ銀行\n記号：15480　番号：38175201\n（普通：五四八店　3817520）\n口座名義：マウナレイフラコンペティション事務局（マウナレイフラコンペティションジムキョク）\n※恐れ入りますが、振込手数料はご負担くださいますようお願いいたします。';
const OFFICE_INFO = 'マウナレイフラコンペティション事務局\n〒700-0822 岡山県岡山市北区表町3-18-52 吉本ビル4F ハウオリーズ マサコ アケタ フラスタジオ内\nTel：086-231-2314　Fax：086-231-2324\nE-mail：maunaleihulacompe@gmail.com';
const TICKET_PRICE = 7000;
const TICKET_SHIPPING_FEE = 500;
const TICKET_ALERT_THRESHOLD = 150;

const FORM_CONFIG = {
  lei: {sheet:'レイ申込',subject:'レイ・コンテスト申込',fields:[['applicantName','申込者名'],['furigana','フリガナ'],['postalCode','郵便番号'],['address','住所'],['phone','電話番号'],['fax','FAX番号'],['email','メールアドレス'],['workTitle','作品タイトル'],['leiStyle','レイのスタイル'],['deliveryMethod','提出方法'],['materials','主な花材'],['comment','作品コメント'],['agreement','規程への同意']]},
  photo: {sheet:'フォト申込',subject:'フォト・コンテスト申込',fields:[['applicantName','申込者名'],['furigana','フリガナ'],['postalCode','郵便番号'],['address','住所'],['phone','電話番号'],['fax','FAX番号'],['email','メールアドレス'],['workTitle1','作品1タイトル'],['workTitleKana1','作品1タイトルふりがな'],['size1','作品1サイズ'],['deliveryMethod1','作品1提出方法'],['comment1','作品1コメント'],['workTitle2','作品2タイトル'],['workTitleKana2','作品2タイトルふりがな'],['size2','作品2サイズ'],['deliveryMethod2','作品2提出方法'],['comment2','作品2コメント'],['agreement','規程への同意']]},
  vendor: {sheet:'出店申込',subject:'会場出店申込',fields:[['businessName','出店登録名'],['businessKana','フリガナ'],['representative','代表者名'],['representativeKana','代表者フリガナ'],['postalCode','郵便番号'],['address','住所'],['phone','電話番号'],['fax','FAX番号'],['email','メールアドレス'],['vendorType','出店形態'],['products','出店内容・取扱商品'],['notes','備考'],['agreement','案内への同意'],['invoiceFileId','請求書ファイルID']]},
  ad: {sheet:'広告申込',subject:'プログラム広告掲載申込',fields:[['companyName','社名'],['companyKana','フリガナ'],['contactName','担当者名'],['contactKana','担当者フリガナ'],['postalCode','郵便番号'],['address','住所'],['phone','電話番号'],['fax','FAX番号'],['email','メールアドレス'],['adSize','広告サイズ'],['adAmount','広告掲載料'],['notes','備考'],['agreement','案内への同意'],['invoiceFileId','請求書ファイルID']]},
  competition: {sheet:'コンペ申込',subject:'コンペ・エキシビション申込',fields:[['harauName','ハラウ名'],['repLastName','代表者氏名（姓）'],['repFirstName','代表者氏名（名）'],['postalCode','郵便番号'],['address','住所'],['phone','電話番号'],['email','メールアドレス'],['soloCategories','Soloエントリー','multi'],['groupCategories','Groupエントリー','multi'],['exhibition','エキシビションエントリー'],['notes','備考'],['agreement','参加規程への同意']]},
  ticket: {sheet:'チケット申込',subject:'チケット申込',fields:[['applicantName','申込者名'],['furigana','フリガナ'],['postalCode','郵便番号'],['address','住所'],['phone','電話番号'],['email','メールアドレス'],['quantity','枚数'],['notes','備考'],['agreement','同意'],['amount','合計金額'],['dueDate','お振込期限']]}
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
    if(sheet.getLastRow()===0){
      sheet.appendRow(headers);
    }else{
      const existing=sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),headers.length)).getValues()[0];
      const isSame=headers.every((h,i)=>existing[i]===h);
      if(!isSame) sheet.getRange(1,1,1,headers.length).setValues([headers]);
    }
    const values=[new Date()].concat(config.fields.map(item=>getFieldValue(e,item)));
    sheet.appendRow(values);
    const rowNum=sheet.getLastRow();
    if(type==='vendor'||type==='ad') createInvoicePdf(type,e,sheet,rowNum);
    let adSizeLabel='';
    let adAmountNum=0;
    if(type==='ad'){
      const parsedAd=parseAdItem(e.parameter.adSize);
      adSizeLabel=(e.parameter.adSize||'').replace(/\s*[\d,]+円\s*$/,'');
      adAmountNum=parsedAd.amount;
      setCellByHeader(sheet,rowNum,'広告サイズ',adSizeLabel);
      setCellByHeader(sheet,rowNum,'広告掲載料',adAmountNum);
    }
    let ticketAmountText='';
    let ticketDueDate='';
    let ticketQty=0;
    let ticketTotalSold=0;
    if(type==='ticket'){
      ticketQty=parseInt(e.parameter.quantity,10);
      if(!ticketQty||ticketQty<1) ticketQty=1;
      const amount=ticketQty*TICKET_PRICE+TICKET_SHIPPING_FEE;
      const due=new Date(Date.now()+14*24*60*60*1000);
      ticketAmountText=amount.toLocaleString('ja-JP')+'円';
      ticketDueDate=Utilities.formatDate(due,'Asia/Tokyo','yyyy年MM月dd日');
      setCellByHeader(sheet,rowNum,'合計金額',amount);
      setCellByHeader(sheet,rowNum,'お振込期限',ticketDueDate);
      ticketTotalSold=sumColumnByHeader(sheet,'枚数');
      const previousSold=ticketTotalSold-ticketQty;
      if(previousSold<TICKET_ALERT_THRESHOLD&&ticketTotalSold>=TICKET_ALERT_THRESHOLD){
        MailApp.sendEmail({to:NOTIFY_EMAIL,subject:'【重要】一般チケット（Web販売分）の申込が'+TICKET_ALERT_THRESHOLD+'枚に到達しました',body:'一般チケット（Web販売分）の累計申込枚数が'+TICKET_ALERT_THRESHOLD+'枚に到達しました。\n\n現在の一般チケット累計申込枚数：'+ticketTotalSold+'枚\n※先行チケット分は含みません。\n\n受付を継続するか停止するか、ご確認・ご判断をお願いいたします。\n\n※このメールは一般チケットの累計枚数が'+TICKET_ALERT_THRESHOLD+'枚を超えた際に自動送信されています。'});
      }
    }
    const lines=config.fields.map(function(item){
      if(type==='ticket'&&item[0]==='amount') return item[1]+'：'+ticketAmountText;
      if(type==='ticket'&&item[0]==='dueDate') return item[1]+'：'+ticketDueDate;
      if(type==='ad'&&item[0]==='adSize') return item[1]+'：'+adSizeLabel;
      if(type==='ad'&&item[0]==='adAmount') return item[1]+'：'+adAmountNum.toLocaleString('ja-JP')+'円';
      return item[1]+'：'+getFieldValue(e,item);
    }).join('\n');
    const officeExtra=type==='ticket'?'\n\n一般チケット累計申込枚数：'+ticketTotalSold+'枚（先行チケット分は含みません）':'';
    MailApp.sendEmail({to:NOTIFY_EMAIL,subject:'【Web申込】'+config.subject,body:'Webサイトから申込が届きました。\n\n'+lines+officeExtra});
    const reply=e.parameter.email;
    let noteText='';
    if(type==='lei') noteText='\n\n【ご注意】\n作品と【作品の説明文】は、2026年11月27日（金）午前中までに事務局へお届けください（生花のため、この日より前はお預かりできません）。\n【作品の説明文】は、大会公式サイトの参加規程PDFを印刷し、必要事項をご記入のうえご用意ください。';
    if(type==='ticket') noteText='\n\n【お支払いについて】\nチケット代：'+(ticketQty*TICKET_PRICE).toLocaleString('ja-JP')+'円（'+ticketQty+'枚）\n送料：'+TICKET_SHIPPING_FEE.toLocaleString('ja-JP')+'円\n【合計金額】'+ticketAmountText+'\nお振込期限：'+ticketDueDate+'\n\n【お振込先】\n'+BANK_INFO;
    if(reply) MailApp.sendEmail({to:reply,subject:'【Maunalei】'+config.subject+'を受け付けました',body:'お申し込みありがとうございます。\n以下の内容で受け付けました。\n\n'+lines+noteText+'\n\nマウナレイフラコンペティション事務局'});
    return ContentService.createTextOutput('OK');
  }catch(error){
    MailApp.sendEmail(NOTIFY_EMAIL,'【Maunalei】フォーム送信エラー',String(error));
    return ContentService.createTextOutput('ERROR');
  }
}

function escapeHtml(text){
  return String(text).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function parseAdItem(adSize){
  adSize=adSize||'';
  const sizeLabel=adSize.replace(/\s*[\d,]+円\s*$/,'');
  const m=adSize.match(/([\d,]+)\s*円/);
  const amount=m?parseInt(m[1].replace(/,/g,''),10):0;
  return {desc:'広告掲載料（'+sizeLabel+'）',amount:amount};
}

function findMatchingSubmission(sheetName,email){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName(sheetName);
  if(!sheet||sheet.getLastRow()<2) return null;
  const data=sheet.getDataRange().getValues();
  const headers=data[0];
  const emailCol=headers.indexOf('メールアドレス');
  const invoiceCol=headers.indexOf('請求書ファイルID');
  if(emailCol===-1) return null;
  for(let r=data.length-1;r>=1;r--){
    if(data[r][emailCol]&&data[r][emailCol]===email){
      const rowData=data[r];
      return{
        sheet:sheet,
        rowNum:r+1,
        get:function(headerName){const c=headers.indexOf(headerName);return c===-1?'':rowData[c]},
        invoiceFileId:invoiceCol===-1?'':rowData[invoiceCol]
      };
    }
  }
  return null;
}

function setCellByHeader(sheet,rowNum,headerName,value){
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const col=headers.indexOf(headerName);
  if(col!==-1) sheet.getRange(rowNum,col+1).setValue(value);
}

function setInvoiceFileId(sheet,rowNum,fileId){
  setCellByHeader(sheet,rowNum,'請求書ファイルID',fileId);
}

function sumColumnByHeader(sheet,headerName){
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const col=headers.indexOf(headerName);
  if(col===-1) return 0;
  const lastRow=sheet.getLastRow();
  if(lastRow<2) return 0;
  const values=sheet.getRange(2,col+1,lastRow-1,1).getValues();
  return values.reduce(function(sum,row){return sum+(parseInt(row[0],10)||0)},0);
}

function createInvoicePdf(type,e,sheet,rowNum){
  try{
    const email=e.parameter.email||'';
    let orgName,repName;
    const items=[];
    if(type==='vendor'){
      orgName=e.parameter.businessName||'';
      repName=e.parameter.representative||'';
      items.push({desc:'出店料（1店舗）',amount:20000});
    }else{
      orgName=e.parameter.companyName||'';
      repName=e.parameter.contactName||'';
      items.push(parseAdItem(e.parameter.adSize));
    }

    let matched=null;
    let combinedNote='';
    if(email){
      const otherSheetName=type==='vendor'?'広告申込':'出店申込';
      matched=findMatchingSubmission(otherSheetName,email);
      if(matched){
        if(type==='vendor'){
          items.push(parseAdItem(matched.get('広告サイズ')));
        }else{
          items.push({desc:'出店料（1店舗）',amount:20000});
        }
        combinedNote='<div style="margin:10px 0 0;font-size:12px;color:#8a4a25">※出店・広告のお申込みをまとめてご請求しています。</div>';
        if(matched.invoiceFileId){
          try{DriveApp.getFileById(matched.invoiceFileId).setTrashed(true);}catch(ignoreErr){}
        }
      }
    }

    const amount=items.reduce(function(sum,it){return sum+it.amount},0);
    const now=new Date();
    const due=new Date(now.getTime()+14*24*60*60*1000);
    const invoiceNo=Utilities.formatDate(now,'Asia/Tokyo','yyyyMMdd')+'-'+Utilities.formatDate(now,'Asia/Tokyo','HHmmss');
    const issueDate=Utilities.formatDate(now,'Asia/Tokyo','yyyy年MM月dd日');
    const dueDate=Utilities.formatDate(due,'Asia/Tokyo','yyyy年MM月dd日');
    const amountText=amount.toLocaleString('ja-JP');
    const nl2br=s=>escapeHtml(s).replace(/\n/g,'<br>');
    const itemRows=items.map(function(it,i){
      const bg=i%2===0?'#ffffff':'#F7F9FC';
      return '<tr><td style="border:1px solid #C8D5E5;padding:8px 10px;background:'+bg+'">'+escapeHtml(it.desc)+'</td><td style="border:1px solid #C8D5E5;padding:8px 10px;background:'+bg+';text-align:right">'+it.amount.toLocaleString('ja-JP')+'円</td></tr>';
    }).join('');

    const html='<html><head><meta charset="UTF-8"><style>@page{size:A4 portrait;margin:0}body{margin:0;padding:0}table{page-break-inside:avoid}tr{page-break-inside:avoid}td,th{overflow-wrap:break-word;word-wrap:break-word}th{font-weight:normal}</style></head><body><div style="font-family:\'Noto Sans JP\',sans-serif;padding:48px 55px 42px;max-width:684px;margin:auto;color:#303B4D;font-size:13px;line-height:1.65">'
      +'<h1 style="text-align:center;font-size:30px;font-weight:normal;letter-spacing:.3em;margin:0 0 18px;color:#145C74">請求書</h1>'
      +'<div style="text-align:right;font-size:11px;color:#657084;margin-bottom:30px;line-height:1.8">発行日：'+escapeHtml(issueDate)+'<br>請求書番号：'+escapeHtml(invoiceNo)+'</div>'
      +'<div style="margin-bottom:3px;font-size:17px;font-weight:normal">出展名：'+escapeHtml(orgName)+'</div>'
      +'<div style="margin-bottom:32px;font-size:15px">代表者名：'+escapeHtml(repName)+'　様</div>'
      +'<table style="width:100%;border-collapse:collapse;margin-bottom:24px">'
      +'<tr><td style="background:#E0F2FC;padding:14px 12px;width:41%;font-weight:normal;font-size:14px;color:#394C5D;border:1px solid #0D84AD;border-right:none">ご請求金額合計</td><td style="background:#ECFCF5;padding:14px 12px;text-align:center;font-size:26px;font-weight:normal;letter-spacing:.08em;color:#00867B;border:1px solid #0D84AD;border-left:none">'+amountText+'円</td></tr>'
      +'</table>'
      +'<table style="width:100%;border-collapse:collapse;margin-bottom:8px;font-size:13px">'
      +'<tr><th bgcolor="#E0F2FC" style="background-color:#E0F2FC !important;color:#303B4D !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;padding:8px 10px;text-align:center;border:1px solid #C8D5E5">内訳</th><th bgcolor="#E0F2FC" style="background-color:#E0F2FC !important;color:#303B4D !important;-webkit-print-color-adjust:exact;print-color-adjust:exact;padding:8px 10px;width:140px;text-align:center;border:1px solid #C8D5E5">金額</th></tr>'
      +itemRows
      +'<tr><td style="border:none;border-top:1px solid #0D84AD;padding:8px 10px;text-align:right;font-weight:normal;background:#E0F2FC;color:#303B4D">合計</td><td style="border:none;border-top:1px solid #0D84AD;padding:8px 10px;text-align:right;font-weight:normal;background:#E0F2FC;color:#303B4D">'+amountText+'円</td></tr>'
      +'</table>'
      +combinedNote
      +'<table style="width:100%;border-collapse:collapse;margin:24px 0">'
      +'<tr><td style="background:#FFF4C8;padding:10px 12px;width:26%;font-weight:normal;font-size:14px;color:#4F4B38;border:1px solid #EE9A30;border-right:none">お振込期限</td><td style="background:#FFFFFF;padding:10px 12px;font-size:14px;font-weight:normal;color:#303B4D;border:1px solid #EE9A30;border-left:none">'+escapeHtml(dueDate)+'</td></tr>'
      +'</table>'
      +'<table style="width:100%;border-collapse:collapse;margin-bottom:0"><tr><td style="background:#FFFFFF;padding:0 8px 8px;font-weight:normal;font-size:14px;color:#303B4D">お振込先</td></tr></table>'
      +'<table style="width:100%;border-collapse:collapse;margin-bottom:24px">'
      +'<tr><td style="background:#F3F7FB;padding:12px 14px;border:1px solid #C8D5E5">'+nl2br(BANK_INFO)+'</td></tr>'
      +'</table>'
      +'<table style="width:100%;border-collapse:collapse;margin-bottom:0"><tr><td style="background:#FFFFFF;padding:0 8px 8px;font-weight:normal;font-size:12px;color:#657084;text-align:right">発行元</td></tr></table>'
      +'<table style="width:100%;border-collapse:collapse;margin-bottom:10px">'
      +'<tr><td style="background:#FFFFFF;padding:0 8px;border:none;font-size:11px;color:#657084;text-align:right;line-height:1.8">'+nl2br(OFFICE_INFO)+'</td></tr>'
      +'</table>'
      +'</div></body></html>';

    const blob=Utilities.newBlob(html,'text/html','invoice.html').getAs('application/pdf');
    const typeLabel=matched?'出店広告':(type==='vendor'?'出店':'広告');
    const fileName='請求書_'+typeLabel+'_'+orgName+'_'+Utilities.formatDate(now,'Asia/Tokyo','yyyyMMdd')+'.pdf';
    blob.setName(fileName);
    const folders=DriveApp.getFoldersByName(INVOICE_FOLDER_NAME);
    const folder=folders.hasNext()?folders.next():DriveApp.createFolder(INVOICE_FOLDER_NAME);
    const file=folder.createFile(blob);
    setInvoiceFileId(sheet,rowNum,file.getId());
    if(matched) setInvoiceFileId(matched.sheet,matched.rowNum,file.getId());
  }catch(error){
    MailApp.sendEmail(NOTIFY_EMAIL,'【Maunalei】請求書PDF作成エラー',String(error));
  }
}
