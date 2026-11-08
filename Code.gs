const SHEETS = {
  Programs:['programId','programName','faculty','contact','createdAt'],
  Assessments:['id','programId','cycle','dimension','score','note','updatedAt'],
  Evidence:['id','programId','cycle','criterion','title','url','owner','status','updatedAt'],
  Actions:['id','programId','cycle','issue','action','owner','dueDate','status','result','updatedAt']
};
const DIMENSIONS = ['P1 นโยบายและภาวะผู้นำ','P2 บูรณาการกระบวนการ','P3 บริหารหลักฐาน','P4 ติดตามผล','P5 ปรับปรุงต่อเนื่อง','P6 บุคลากรและการเรียนรู้'];
function setup() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SPREADSHEET_ID');
  if (!id) { const ss = SpreadsheetApp.create('FQA-6P Research Data'); id=ss.getId(); props.setProperty('SPREADSHEET_ID',id); }
  const ss=SpreadsheetApp.openById(id);
  Object.entries(SHEETS).forEach(([name,cols])=>{
    let sh=ss.getSheetByName(name);
    if(!sh) sh=ss.insertSheet(name);
    if(sh.getLastRow()===0){sh.appendRow(cols);sh.setFrozenRows(1);sh.getRange(1,1,1,cols.length).setFontWeight('bold').setBackground('#dceaf5');}
  });
  Logger.log('Spreadsheet: '+ss.getUrl());
  return ss.getUrl();
}
function doGet() { return HtmlService.createHtmlOutputFromFile('Index').setTitle('FQA-6P | Faculty Quality Assurance').addMetaTag('viewport','width=device-width, initial-scale=1'); }
function sheet_(name){const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');if(!id)throw new Error('กรุณาเรียกใช้ setup() ก่อน');return SpreadsheetApp.openById(id).getSheetByName(name);}
function rows_(name){const sh=sheet_(name),v=sh.getDataRange().getValues();return v.slice(1).filter(r=>r[0]!=='' ).map(r=>Object.fromEntries(SHEETS[name].map((k,i)=>[k,r[i] instanceof Date?Utilities.formatDate(r[i],Session.getScriptTimeZone(),'yyyy-MM-dd'):r[i]])));}
function clean_(v){return String(v==null?'':v).trim().slice(0,1000);}
function append_(name,record){const fields=SHEETS[name];sheet_(name).appendRow(fields.map(k=>record[k]??''));}
function now_(){return new Date().toISOString();}
function auth_(programId,pin){
 const p=rows_('Programs').find(x=>x.programId===clean_(programId));
 if(!p)throw new Error('ไม่พบรหัสหลักสูตร');
 // PIN hashes live in script properties, not the spreadsheet.
 const stored=PropertiesService.getScriptProperties().getProperty('PIN_'+p.programId);
 if(!stored||stored!==hash_(clean_(pin)))throw new Error('รหัสผ่านหลักสูตรไม่ถูกต้อง');
 return p;
}
function hash_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s).map(x=>('0'+(x&255).toString(16)).slice(-2)).join('');}
function registerProgram(data){
 const name=clean_(data.programName),pin=clean_(data.pin),contact=clean_(data.contact);
 if(name.length<3||pin.length<8)throw new Error('ระบุชื่อหลักสูตรและรหัสผ่านอย่างน้อย 8 ตัวอักษร');
 const lock=LockService.getScriptLock();lock.waitLock(15000);
 try{
   if(rows_('Programs').some(p=>p.programName===name))throw new Error('มีชื่อหลักสูตรนี้แล้ว');
   const id='PRG-'+Utilities.getUuid().slice(0,8).toUpperCase();
   append_('Programs',{programId:id,programName:name,faculty:'คณะวิทยาการจัดการ',contact,createdAt:now_()});
   PropertiesService.getScriptProperties().setProperty('PIN_'+id,hash_(pin));
   return {programId:id,programName:name};
 }finally{lock.releaseLock();}
}
function login(programId,pin){const p=auth_(programId,pin);return {programId:p.programId,programName:p.programName};}
function loadProgram(programId,pin,cycle){auth_(programId,pin);const c=clean_(cycle);return {
  dimensions:DIMENSIONS,
  assessments:rows_('Assessments').filter(x=>x.programId===programId&&x.cycle===c),
  evidence:rows_('Evidence').filter(x=>x.programId===programId&&x.cycle===c),
  actions:rows_('Actions').filter(x=>x.programId===programId&&x.cycle===c)
};}
function saveAssessment(data){
 auth_(data.programId,data.pin);const cycle=clean_(data.cycle),dim=clean_(data.dimension),score=Number(data.score);
 if(!/^25\d{2}$/.test(cycle)||!DIMENSIONS.includes(dim)||!Number.isInteger(score)||score<1||score>5)throw new Error('ข้อมูลประเมินไม่ถูกต้อง');
 const lock=LockService.getScriptLock();lock.waitLock(15000);
 try{
   const sh=sheet_('Assessments'),all=sh.getDataRange().getValues();let found=0;
   for(let i=1;i<all.length;i++)if(all[i][1]===data.programId&&String(all[i][2])===cycle&&all[i][3]===dim){found=i+1;break;}
   const r={id:found?all[found-1][0]:Utilities.getUuid(),programId:data.programId,cycle,dimension:dim,score,note:clean_(data.note),updatedAt:now_()};
   if(found)sh.getRange(found,1,1,SHEETS.Assessments.length).setValues([SHEETS.Assessments.map(k=>r[k])]);else append_('Assessments',r);
   return {ok:true};
 }finally{lock.releaseLock();}
}
function addEvidence(data){
 auth_(data.programId,data.pin);if(!clean_(data.title))throw new Error('กรุณาระบุชื่อหลักฐาน');
 const url=clean_(data.url);if(url&&!/^https:\/\//i.test(url))throw new Error('ลิงก์หลักฐานต้องเริ่มด้วย https://');
 append_('Evidence',{id:Utilities.getUuid(),programId:data.programId,cycle:clean_(data.cycle),criterion:clean_(data.criterion),title:clean_(data.title),url,owner:clean_(data.owner),status:clean_(data.status)||'รอตรวจสอบ',updatedAt:now_()});return {ok:true};
}
function addAction(data){
 auth_(data.programId,data.pin);if(!clean_(data.issue)||!clean_(data.action))throw new Error('กรุณาระบุปัญหาและมาตรการ');
 append_('Actions',{id:Utilities.getUuid(),programId:data.programId,cycle:clean_(data.cycle),issue:clean_(data.issue),action:clean_(data.action),owner:clean_(data.owner),dueDate:clean_(data.dueDate),status:clean_(data.status)||'ยังไม่เริ่ม',result:clean_(data.result),updatedAt:now_()});return {ok:true};
}
function getDashboard(adminKey,cycle){
 const stored=PropertiesService.getScriptProperties().getProperty('ADMIN_KEY_HASH');
 if(!stored||hash_(clean_(adminKey))!==stored)throw new Error('รหัสผู้ดูแลระบบไม่ถูกต้อง');
 const c=clean_(cycle),programs=rows_('Programs'),a=rows_('Assessments').filter(x=>x.cycle===c),e=rows_('Evidence').filter(x=>x.cycle===c),act=rows_('Actions').filter(x=>x.cycle===c);
 return programs.map(p=>{const x=a.filter(v=>v.programId===p.programId);return {programId:p.programId,programName:p.programName,assessed:x.length,avg:x.length?Math.round(x.reduce((s,v)=>s+Number(v.score),0)/x.length*100)/100:null,evidence:e.filter(v=>v.programId===p.programId).length,actions:act.filter(v=>v.programId===p.programId).length};});
}
function setAdminKey(){const key='CHANGE-THIS-TO-A-LONG-RANDOM-SECRET';PropertiesService.getScriptProperties().setProperty('ADMIN_KEY_HASH',hash_(key));Logger.log('ตั้งรหัสผู้ดูแลแล้ว: เปลี่ยนค่าตัวแปร key ก่อนใช้จริง');}
