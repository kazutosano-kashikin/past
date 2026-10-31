// 学習ログ受け取り用GAS(PWA版)。デプロイ:「ウェブアプリ」/ 実行ユーザー=自分 / アクセス=全員
const SHEET='学習ログ', HEAD=['ts','日時','id','年度','問','分野','論点','選択','正答','正誤(1=正解)','秒'];
function sheet_(){
  const p=PropertiesService.getScriptProperties();let ss=null;const id=p.getProperty('SSID');
  if(id){try{ss=SpreadsheetApp.openById(id)}catch(e){}}
  if(!ss){ss=SpreadsheetApp.create('貸金業 学習ログ');p.setProperty('SSID',ss.getId())}
  let s=ss.getSheetByName(SHEET);if(!s){s=ss.insertSheet(SHEET);s.appendRow(HEAD)}return s;
}
function doPost(e){
  const s=sheet_(),arr=JSON.parse(e.postData.contents),last=s.getLastRow();
  const have=new Set(last>1?s.getRange(2,1,last-1,1).getValues().map(r=>String(r[0])):[]);
  const rows=arr.filter(x=>!have.has(String(x.ts))).map(x=>[x.ts,new Date(x.ts),x.id,x.y,x.n,x.cat,x.topic||'',x.pick,x.ans,x.ok?1:0,x.sec||0]);
  if(rows.length)s.getRange(last+1,1,rows.length,HEAD.length).setValues(rows);
  return ContentService.createTextOutput('ok');
}
function doGet(e){
  const v=sheet_().getDataRange().getValues();v.shift();
  const out=v.map(r=>({ts:r[0],id:r[2],y:r[3],n:r[4],cat:r[5],topic:r[6],pick:r[7],ans:r[8],ok:r[9]==1,sec:r[10]})).reverse();
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
