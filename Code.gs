const SHEET_ID='1B-DmdtY0FYZQsqbR_y1Nzvmp1p2NalsYcV1tVM0Yc2k';
const FOLDER_ID='1CsZRBf6lvsGCC-4IHoOWIwrYuZ5UCTDq';
const CSV='https://docs.google.com/spreadsheets/d/e/2PACX-1vQsKnkpIzLapkSe78mMmARPoBiA4gepYa9qQ4mNw4w1CjbVSDwH9ANIhUiR6Vcw44xMkSigbZtlq76J/pub?gid=0&single=true&output=csv';
const HEAD=['ID','Timestamp','Sesi','Tarikh','Bilangan Perjumpaan','Kategori','Unit','Tempat','Hadir','Sepatutnya','Peratus','Perincian','Modul PIKeBM','Aktiviti PIKeBM','Gambar 1 ID','Gambar 2 ID','Disediakan Oleh'];
const CATS=['KELAB PERSATUAN','SUKAN PERMAINAN','PASUKAN BADAN BERUNIFORM'];
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
function doGet(e){if(e && e.parameter && e.parameter.action==='getUnits'){try{return json_({ok:true,data:getUnits()});}catch(err){return json_({ok:false,error:String(err.message||err)});}}return json_({ok:true,service:'SMK Seri Ampang Kokurikulum API',version:'github-3'});}
function doPost(e){try{const p=JSON.parse(e.postData.contents);const args=Array.isArray(p.args)?p.args:[];let data;switch(p.action){case 'getUnits':data=getUnits();break;case 'listReports':data=listReports();break;case 'saveReport':data=saveReport(args[0]);break;case 'reportPhotos':data=reportPhotos(args[0]);break;default:throw Error('Tindakan tidak sah.');}return json_({ok:true,data:data});}catch(err){return json_({ok:false,error:String(err.message||err)});}}

function sheet_(){const s=SpreadsheetApp.openById(SHEET_ID);let t=s.getSheetByName('Laporan Perjumpaan');if(!t)t=s.insertSheet('Laporan Perjumpaan');if(!t.getLastRow()){t.appendRow(HEAD);t.setFrozenRows(1);}else if(t.getRange(1,1,1,HEAD.length).getDisplayValues()[0].join('|')!==HEAD.join('|'))throw Error('Header tab Laporan Perjumpaan tidak sepadan.');return t;}
function setup(){sheet_();DriveApp.getFolderById(FOLDER_ID).getName();return getUnits();}
function category_(x){x=String(x).toUpperCase().replace(/[^A-Z ]/g,' ').replace(/\s+/g,' ').trim();if(/BERUNIFORM|UNIFORM|^UB$|^PBB$/.test(x))return CATS[2];if(/SUKAN|PERMAINAN|^SP$/.test(x))return CATS[1];if(/KELAB|PERSATUAN|^KP$/.test(x))return CATS[0];return '';}
function getUnits(){
 const res=UrlFetchApp.fetch(CSV,{muteHttpExceptions:true,followRedirects:true});
 if(res.getResponseCode()!==200)throw Error('CSV unit gagal dimuat: HTTP '+res.getResponseCode());
 const rows=Utilities.parseCsv(res.getContentText().replace(/^\uFEFF/,''));
 const norm=x=>String(x||'').trim().toUpperCase().replace(/\s+/g,' ');
 const hi=rows.findIndex(r=>r.map(norm).includes('KATEGORI KOKURIKULUM')&&r.map(norm).includes('NAMA UNIT')); const h=(rows[hi]||[]).map(norm), ci=h.indexOf('KATEGORI KOKURIKULUM'), ui=h.indexOf('NAMA UNIT');
 if(ci<0||ui<0)throw Error('Header CSV mesti KATEGORI KOKURIKULUM dan NAMA UNIT.');
 const out={};CATS.forEach(c=>out[c]=[]);let last='';
 rows.slice(hi+1).forEach(r=>{const raw=String(r[ci]||'').trim();if(raw)last=category_(raw);const v=String(r[ui]||'').trim();if(last&&v&&!out[last].some(x=>norm(x)===norm(v)))out[last].push(v);});
 CATS.forEach(c=>out[c].sort((a,b)=>a.localeCompare(b)));
 if(!CATS.some(c=>out[c].length))throw Error('Tiada nama unit yang sepadan dalam CSV.');
 return out;
}
function testUnits(){const result=getUnits();Logger.log(JSON.stringify(result));return result;}
function safe_(v){v=String(v||'');return /^[=+\-@]/.test(v)?"'"+v:v;}
function saveReport(d){if(!d||!['PAGI','PETANG'].includes(d.sesi)||!/^\d{4}-\d{2}-\d{2}$/.test(d.tarikh)||!CATS.includes(d.kategori)||!d.unit||!d.tempat||!d.perincian||!d.nama||!String(d.bil).trim())throw Error('Lengkapkan semua maklumat wajib.');const a=Number(d.hadir),b=Number(d.sepatutnya);if(!Number.isInteger(a)||!Number.isInteger(b)||a<0||b<=0||a>b)throw Error('Kehadiran tidak sah.');if(!Array.isArray(d.photos)||d.photos.length!==2)throw Error('Dua gambar diperlukan.');const units=getUnits();if(!units[d.kategori].includes(d.unit))throw Error('Unit tidak sah. Muat semula senarai unit.');const lock=LockService.getScriptLock();lock.waitLock(30000);const files=[];try{const sh=sheet_();const id=String(d.id||'');if(!/^[a-zA-Z0-9-]{15,80}$/.test(id))throw Error('ID tidak sah.');if(sh.getLastRow()>1&&sh.getRange(2,1,sh.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).findNext())return {id:id};const folder=DriveApp.getFolderById(FOLDER_ID);d.photos.forEach((p,i)=>{const m=/^data:image\/jpeg;base64,([A-Za-z0-9+/=]+)$/.exec(p);if(!m||m[1].length>2500000)throw Error('Format atau saiz gambar tidak sah.');files.push(folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[1]),'image/jpeg',id+'_'+(i+1)+'.jpg')));});sh.appendRow([id,new Date(),d.sesi,"'"+d.tarikh,safe_(d.bil),d.kategori,safe_(d.unit),safe_(d.tempat),a,b,Math.round(a/b*10000)/100,safe_(d.perincian),safe_((d.modul||[]).join(' | ')),safe_(d.aktiviti),files[0].getId(),files[1].getId(),safe_(d.nama)]);return {id:id};}catch(e){files.forEach(f=>{try{f.setTrashed(true);}catch(_) {}});throw e;}finally{lock.releaseLock();}}
function listReports(){const sh=sheet_();if(sh.getLastRow()<2)return [];return sh.getRange(2,1,sh.getLastRow()-1,HEAD.length).getDisplayValues().map(r=>({id:r[0],sesi:r[2],tarikh:r[3],bil:r[4],kategori:r[5],unit:r[6],tempat:r[7],hadir:Number(r[8]),sepatutnya:Number(r[9]),peratus:Number(r[10]),perincian:r[11],modul:r[12]?r[12].split(' | '):[],aktiviti:r[13],nama:r[16]})).reverse();}
function reportPhotos(id){const sh=sheet_();if(sh.getLastRow()<2)throw Error('Rekod tidak dijumpai.');const f=sh.getRange(2,1,sh.getLastRow()-1,1).createTextFinder(String(id)).matchEntireCell(true).findNext();if(!f)throw Error('Rekod tidak dijumpai.');return sh.getRange(f.getRow(),15,1,2).getDisplayValues()[0].map(x=>{const b=DriveApp.getFileById(x).getBlob();return 'data:image/jpeg;base64,'+Utilities.base64Encode(b.getBytes());});}
