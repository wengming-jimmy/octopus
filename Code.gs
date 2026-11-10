// 歌唱比賽 章魚哥獎 後端（Google Apps Script，綁定在「章魚哥獎資料庫」試算表）
const T={settings:'設定',players:'選手',picks:'預測登記',champs:'冠軍',winners:'章魚哥獎'};
const GROUPS=['對唱','新人','獨唱','長青'];
const ZONES=['第一專區','第二專區','第三專區','第四專區','第五專區','第六專區','第七專區','第八專區','第九專區','第十專區','第十一專區','第十二專區','第十三專區'];
const PICK_HEAD=['時間','專區','分會','姓名','手機'].concat(GROUPS.map(g=>g+'組預測'),['兌獎編號','驗證碼']);

function ss(){return SpreadsheetApp.getActiveSpreadsheet()}
function sh(n,head){let s=ss().getSheetByName(n);if(!s){s=ss().insertSheet(n);if(head){s.getRange(1,1,1,head.length).setValues([head]).setFontWeight('bold');s.setFrozenRows(1)}}return s}
function rows(n){const s=ss().getSheetByName(n);if(!s)return[];const v=s.getDataRange().getDisplayValues();v.shift();return v.filter(r=>r.some(c=>c!==''))}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function write(n,head,data){const s=sh(n,head);s.clearContents();const all=[head].concat(data);
  const rg=s.getRange(1,1,all.length,head.length);rg.setNumberFormat('@');rg.setValues(all);s.getRange(1,1,1,head.length).setFontWeight('bold');s.setFrozenRows(1)}

function setup(){ // 第一次執行會建立分頁與預設值
  const s=sh(T.settings,['項目','值']);
  if(s.getLastRow()<2){s.getRange(2,1,9,2).setNumberFormat('@').setValues([
    ['eventName','2026-2027年度歌唱才藝總決賽'],['adminPassword','octo2127'],['open','TRUE'],
    ['minHits','2'],['published','FALSE'],['seq','0'],['deadline',''],['zones',ZONES.join(',')],['groups',GROUPS.join(',')]])}
  sh(T.players,['組別','姓名','分會']);sh(T.picks,PICK_HEAD).getRange('E:E').setNumberFormat('@');
  sh(T.champs,['組別','冠軍']);sh(T.winners,['兌獎編號','驗證碼','專區','分會','姓名','手機','猜中組數','猜中組別']);
  const d=ss().getSheetByName('工作表1');if(d&&ss().getSheets().length>1&&d.getLastRow()===0)ss().deleteSheet(d);
}
function settings(){if(!ss().getSheetByName(T.settings))setup();const o={};rows(T.settings).forEach(r=>o[r[0]]=r[1]);return o}
function setSetting(k,v){const s=sh(T.settings);const v2=s.getDataRange().getValues();
  for(let i=1;i<v2.length;i++)if(v2[i][0]===k){s.getRange(i+1,2).setNumberFormat('@').setValue(String(v));return}
  s.appendRow([k,String(v)])}
const TRUE=v=>String(v).toUpperCase()==='TRUE';
const label=p=>p.club?`${p.name}（${p.club}）`:p.name;
const mask=p=>String(p).replace(/^(\d{4})\d{3}(\d{3})$/,'$1***$2');

function players(){return rows(T.players).map(r=>({g:r[0].replace(/組$/,''),name:r[1],club:r[2]}))}
function champs(){const o={};rows(T.champs).forEach(r=>{if(r[1])o[r[0]]=r[1]});return o}
function picks(){return rows(T.picks).map(r=>{const p={time:r[0],zone:r[1],club:r[2],name:r[3],phone:r[4],picks:{}};GROUPS.forEach((g,i)=>p.picks[g]=r[5+i]||'');p.no=r[9]||'';p.code=r[10]||'';return p})}
function winners(st){const ch=champs(),min=Number(st.minHits)||2;
  return picks().map(p=>{const hit=GROUPS.filter(g=>ch[g]&&p.picks[g]===ch[g]);return {...p,hits:hit.length,hitGroups:hit}})
    .filter(p=>p.hits>=min).sort((a,b)=>b.hits-a.hits||(a.time<b.time?-1:1))}
function deadlineMs(st){const d=String(st.deadline||'').trim();if(!d)return 0;const t=new Date(d.replace(' ','T')+(d.length<=16?':00':'')+'+08:00').getTime();return isNaN(t)?0:t}
function isOpen(st){const dl=deadlineMs(st);return TRUE(st.open)&&(!dl||Date.now()<dl)}
function publicData(){const st=settings();const pub=TRUE(st.published);
  return {eventName:st.eventName,open:isOpen(st),openSwitch:TRUE(st.open),deadline:st.deadline||'',now:Date.now(),minHits:Number(st.minHits)||2,published:pub,
    zones:(st.zones||ZONES.join(',')).split(',').map(s=>s.trim()).filter(Boolean),groups:GROUPS,players:players(),count:rows(T.picks).length,
    champions:pub?champs():{},winners:pub?winners(st).map(w=>({no:w.no,zone:w.zone,club:w.club,name:w.name,phone:mask(w.phone),hits:w.hits,hitGroups:w.hitGroups})):[]}}
function writeWinners(st){write(T.winners,['兌獎編號','驗證碼','專區','分會','姓名','手機','猜中組數','猜中組別'],winners(st).map(w=>[w.no,w.code,w.zone,w.club,w.name,w.phone,String(w.hits),w.hitGroups.join('、')]))}

function submit(b){
  const st=settings();if(!isOpen(st))return {ok:false,error:'預測登記已截止'};
  const zones=(st.zones||'').split(',').map(s=>s.trim());
  const phone=String(b.phone||'').replace(/\D/g,'');const name=String(b.name||'').trim(),club=String(b.club||'').trim(),zone=String(b.zone||'').trim();
  if(zones.indexOf(zone)<0)return {ok:false,error:'請選擇專區'};
  if(!club||!name)return {ok:false,error:'請填寫分會與姓名'};
  if(!/^09\d{8}$/.test(phone))return {ok:false,error:'手機號碼格式不正確（09 開頭共 10 碼）'};
  const pl=players();const pk=b.picks||{};
  for(const g of GROUPS){if(!pl.some(p=>p.g===g&&label(p)===pk[g]))return {ok:false,error:`請選擇${g}組的預測冠軍`}}
  if(picks().some(p=>p.phone===phone))return {ok:false,error:'這支手機已經登記過了，每人限登記一次'};
  const t=Utilities.formatDate(new Date(),'Asia/Taipei','yyyy/MM/dd HH:mm:ss');
  const seq=(parseInt(st.seq)||0)+1;setSetting('seq',seq);const no='OC'+('000'+seq).slice(-4);
  const code=Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,no+phone+t).map(b=>('0'+(b&255).toString(16)).slice(-2)).join('').slice(0,6).toUpperCase();
  const s=sh(T.picks,PICK_HEAD);const row=s.getLastRow()+1;
  const rg=s.getRange(row,1,1,PICK_HEAD.length);rg.setNumberFormat('@');rg.setValues([[t,zone,club,name,phone].concat(GROUPS.map(g=>pk[g]),[no,code])]);
  return {ok:true,time:t,no,code};
}
function mine(b){const phone=String(b.phone||'').replace(/\D/g,'');const st=settings();
  const p=picks().find(x=>x.phone===phone);if(!p)return {ok:false,error:'查無這支手機的登記'};
  const r={ok:true,no:p.no,code:p.code,phone:mask(p.phone),zone:p.zone,club:p.club,name:p.name,time:p.time,picks:p.picks,published:TRUE(st.published)};
  if(r.published){const ch=champs();r.champions=ch;r.hits=GROUPS.filter(g=>ch[g]&&p.picks[g]===ch[g]);r.win=r.hits.length>=(Number(st.minHits)||2)}
  return r}

function doGet(){const c=CacheService.getScriptCache();let j=c.get('pub');if(!j){j=JSON.stringify(publicData());try{c.put('pub',j,20)}catch(e){}}
  const d=JSON.parse(j);d.now=Date.now();return out({ok:true,data:d})}
function bust(){try{CacheService.getScriptCache().remove('pub')}catch(e){}}
function doPost(e){
  const lock=LockService.getScriptLock();if(!lock.tryLock(30000))return out({ok:false,error:'目前登記人數較多，請再按一次送出'});
  try{
    const b=JSON.parse(e.postData.contents);
    if(b.action==='submit'){const r=submit(b);if(r.ok)bust();return out(r)}
    if(b.action==='mine')return out(mine(b));
    const st=settings();
    if(String(b.password)!==String(st.adminPassword))return out({ok:false,error:'auth'});
    if(b.action!=='getAll')bust();
    if(b.action==='check')return out({ok:true});
    if(b.action==='getAll')return out({ok:true,data:{...publicData(),champions:champs(),picks:picks(),winners:winners(st),published:TRUE(st.published)}});
    if(b.action==='saveSettings'){const d=b.data;
      if(d.eventName!==undefined)setSetting('eventName',d.eventName||'2026-2027年度歌唱才藝總決賽');
      if(d.open!==undefined)setSetting('open',d.open?'TRUE':'FALSE');
      if(d.deadline!==undefined)setSetting('deadline',String(d.deadline||'').replace('T',' '));
      if(d.published!==undefined)setSetting('published',d.published?'TRUE':'FALSE');
      if(d.minHits!==undefined){setSetting('minHits',Math.max(1,Math.min(4,parseInt(d.minHits)||2)));writeWinners(settings())}
      if(d.zones!==undefined)setSetting('zones',d.zones.map(s=>String(s).trim()).filter(Boolean).join(','));
      if(d.newPassword)setSetting('adminPassword',d.newPassword);
      return out({ok:true})}
    if(b.action==='savePlayers'){write(T.players,['組別','姓名','分會'],b.players.map(p=>[p.g,p.name,p.club]));return out({ok:true,count:b.players.length})}
    if(b.action==='saveChampions'){write(T.champs,['組別','冠軍'],GROUPS.map(g=>[g,b.champions[g]||'']));writeWinners(st);return out({ok:true,winners:winners(st).length})}
    if(b.action==='deletePick'){const s=sh(T.picks);const v=s.getDataRange().getDisplayValues();
      for(let i=v.length-1;i>=1;i--)if(v[i][4]===String(b.phone)){s.deleteRow(i+1);writeWinners(st);return out({ok:true})}
      return out({ok:false,error:'找不到這筆登記'})}
    if(b.action==='clearPicks'){write(T.picks,PICK_HEAD,[]);writeWinners(st);return out({ok:true})}
    return out({ok:false,error:'unknown action'});
  }catch(err){return out({ok:false,error:String(err)})}
  finally{lock.releaseLock()}
}
