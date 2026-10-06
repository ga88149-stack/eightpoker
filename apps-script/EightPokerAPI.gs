/**
 * Eight Poker Management API V1.4
 * Paste into EightPokerAPI.gs in the existing Apps Script project.
 * Existing doPost routes action starting with "eight." to eightPokerApi_(body).
 */
const EIGHT_DB_ID='1k_7hXkJoDP8nu9kiFkE-NxZhJf4dPQI40p7gBx6RiuI';
const EIGHT_MEMBER_SHEET='Form Responses 1';
const EIGHT_META_SHEET='EIGHT_MEMBER_META';
const EIGHT_EVENT_SHEET='EPC_賽事';
const EIGHT_PLAYER_SHEET='EPC_賽事玩家';
const EIGHT_SETTINGS_SHEET='EPC_系統設定';

function eightPokerApi_(req){
  try{
    switch(String(req.action||'')){
      case 'eight.ping': return {ok:true,service:'Eight Poker API',version:'V1.5'};
      case 'eight.bootstrap': return eightBootstrap_();
      case 'eight.members.list': return {ok:true,members:eightListMembers_()};
      case 'eight.members.create': return {ok:true,member:eightCreateMember_(req.member||{})};
      case 'eight.members.update': return {ok:true,member:eightUpdateMember_(req.memberKey,req.patch||{})};
      case 'eight.members.delete': return {ok:true,result:eightDeleteMember_(req.memberKey)};
      case 'eight.settings.update': return {ok:true,settings:eightUpdateSettings_(req.settings||{})};
      case 'eight.events.list': return {ok:true,events:eightListEvents_(req.businessDate)};
      case 'eight.events.create': return {ok:true,event:eightCreateEvent_(req.event||{})};
      case 'eight.events.update': return {ok:true,event:eightUpdateEvent_(req.eventId,req.patch||{})};
      case 'eight.events.delete': return {ok:true,result:eightDeleteEvent_(req.eventId)};
      case 'eight.eventPlayers.list': return {ok:true,players:eightListEventPlayers_(req.eventId)};
      case 'eight.eventPlayers.add': return {ok:true,player:eightAddEventPlayer_(req.eventId,req.memberKey)};
      case 'eight.eventPlayers.update': return {ok:true,player:eightUpdateEventPlayer_(req.eventId,req.memberKey,req.patch||{})};
      case 'eight.eventPlayers.delete': return {ok:true,result:eightDeleteEventPlayer_(req.eventId,req.memberKey)};
      case 'eight.eventPlayers.saveAll': return {ok:true,players:eightSaveAllEventPlayers_(req.eventId,req.players||[])};
      default: throw new Error('UNKNOWN_ACTION');
    }
  }catch(err){
    return {ok:false,error:String(err&&err.message?err.message:err),stack:String(err&&err.stack?err.stack:'')};
  }
}
function eightDb_(){return SpreadsheetApp.openById(EIGHT_DB_ID)}
function eightIdx_(h,names){for(const n of names){const i=h.indexOf(n);if(i>=0)return i}return -1}
function eightEnsureSheet_(name,headers){
  const ss=eightDb_();let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);
  if(sh.getLastRow()===0)sh.getRange(1,1,1,headers.length).setValues([headers]);
  else {const width=Math.max(sh.getLastColumn(),headers.length),cur=sh.getRange(1,1,1,width).getDisplayValues()[0];headers.forEach((x,i)=>{if(!cur[i])sh.getRange(1,i+1).setValue(x)})}
  sh.setFrozenRows(1);return sh;
}
function eightMeta_(){const sh=eightDb_().getSheetByName(EIGHT_META_SHEET);if(!sh)throw new Error('META_SHEET_MISSING');return sh}
function eightEvents_(){const sh=eightDb_().getSheetByName(EIGHT_EVENT_SHEET);if(!sh)throw new Error('EVENT_SHEET_MISSING');return sh}
function eightPlayers_(){const sh=eightDb_().getSheetByName(EIGHT_PLAYER_SHEET);if(!sh)throw new Error('PLAYER_SHEET_MISSING');return sh}
function eightSettings_(){
  const sh=eightDb_().getSheetByName(EIGHT_SETTINGS_SHEET),out={businessStart:'16:00',businessEnd:'07:00'};
  if(!sh)return out;const v=sh.getDataRange().getDisplayValues();for(let i=1;i<v.length;i++)if(v[i][0])out[v[i][0]]=v[i][1];return out
}
function eightUpdateSettings_(p){
  const lock=LockService.getScriptLock();lock.waitLock(10000);try{
    const sh=eightEnsureSheet_(EIGHT_SETTINGS_SHEET,['設定鍵','設定值','說明','最後更新時間']);
    const allowed={businessStart:'營業開始時間',businessEnd:'營業結束時間'},rows=sh.getDataRange().getDisplayValues();
    Object.keys(allowed).forEach(k=>{if(!(k in p))return;const val=String(p[k]||'');if(!/^([01]\d|2[0-3]):[0-5]\d$/.test(val))throw new Error('INVALID_TIME');let row=0;for(let i=1;i<rows.length;i++)if(rows[i][0]===k){row=i+1;break}if(row)sh.getRange(row,2,1,3).setValues([[val,allowed[k],new Date()]]);else sh.appendRow([k,val,allowed[k],new Date()])});
    return eightSettings_()
  }finally{lock.releaseLock()}
}
function eightMemberSource_(){
  const sh=eightDb_().getSheetByName(EIGHT_MEMBER_SHEET);if(!sh)throw new Error('MEMBER_SHEET_NOT_FOUND');
  const v=sh.getDataRange().getDisplayValues(),h=v[0]||[];return {sh,v,h,c:{ts:eightIdx_(h,['Timestamp','時間戳記']),name:eightIdx_(h,['姓名']),birth:eightIdx_(h,['出生年月日']),nationalId:eightIdx_(h,['身分證字號']),phone:eightIdx_(h,['手機號碼']),id:eightIdx_(h,['POKER FANS ID','POKERFANS ID']),address:eightIdx_(h,['地址'])}}
}
function eightMetaMap_(){
  const sh=eightMeta_(),map=new Map();if(sh.getLastRow()<2)return map;
  sh.getRange(2,1,sh.getLastRow()-1,8).getDisplayValues().forEach((r,i)=>{const id=String(r[1]||'').trim();if(id&&!map.has(id))map.set(id,{row:i+2,key:r[0],nickname:r[3]||'',group:r[4]||'',status:r[5]||'active'})});return map
}
function eightListMembers_(){
  const ss=eightDb_(),src=ss.getSheetByName(EIGHT_MEMBER_SHEET),metaSh=ss.getSheetByName(EIGHT_META_SHEET);
  if(!src)throw new Error('MEMBER_SHEET_NOT_FOUND');if(!metaSh)throw new Error('META_SHEET_MISSING');
  const sv=src.getDataRange().getDisplayValues(),h=sv[0]||[];
  const c={ts:eightIdx_(h,['Timestamp','時間戳記']),name:eightIdx_(h,['姓名']),birth:eightIdx_(h,['出生年月日']),phone:eightIdx_(h,['手機號碼']),id:eightIdx_(h,['POKER FANS ID','POKERFANS ID']),address:eightIdx_(h,['地址'])};
  if(c.id<0||c.name<0)throw new Error('MEMBER_HEADER_MISSING');
  const mv=metaSh.getLastRow()>1?metaSh.getRange(2,1,metaSh.getLastRow()-1,8).getDisplayValues():[],meta=new Map();
  mv.forEach(r=>{const id=String(r[1]||'').trim();if(id&&!meta.has(id))meta.set(id,{key:r[0],nickname:r[3]||'',group:r[4]||'',status:r[5]||'active'})});
  const out=[];
  for(let i=1;i<sv.length;i++){const r=sv[i],id=String(r[c.id]||'').trim();if(!id)continue;const m=meta.get(id);if(!m||m.status==='inactive')continue;
    out.push({memberKey:m.key,memberId:id,name:r[c.name]||'',nickname:m.nickname,group:m.group,birth:c.birth>=0?r[c.birth]:'',phone:c.phone>=0?r[c.phone]:'',address:c.address>=0?r[c.address]:'',timestamp:c.ts>=0?r[c.ts]:'',status:m.status,eventCount:0,totalEntries:0,pnl:0,spendShare:'',lastVisit:''});
  }return out
}
function eightGenerateMemberId_(source){
  const used=new Set(source.v.slice(1).map(r=>source.c.id>=0?String(r[source.c.id]||'').trim().toUpperCase():''));
  for(let i=0;i<500;i++){const id='A'+String(Math.floor(Math.random()*100000)).padStart(5,'0');if(!used.has(id))return id}
  throw new Error('MEMBER_ID_GENERATION_FAILED')
}
function eightAssertUniqueMemberId_(source,id,excludeRow){
  for(let i=1;i<source.v.length;i++)if(i+1!==excludeRow&&String(source.v[i][source.c.id]||'').trim().toUpperCase()===id.toUpperCase())throw new Error('MEMBER_ID_ALREADY_EXISTS')
}
function eightFindMember_(key){
  const meta=eightMeta_();if(meta.getLastRow()<2)throw new Error('MEMBER_NOT_FOUND');const v=meta.getRange(2,1,meta.getLastRow()-1,8).getDisplayValues();
  for(let i=0;i<v.length;i++)if(v[i][0]===key)return {metaRow:i+2,id:v[i][1],meta:v[i]};throw new Error('MEMBER_NOT_FOUND')
}
function eightFindSourceRowsById_(source,id){const a=[];for(let i=1;i<source.v.length;i++)if(String(source.v[i][source.c.id]||'').trim()===String(id||'').trim())a.push(i+1);return a}
function eightCreateMember_(m){
  const lock=LockService.getScriptLock();lock.waitLock(10000);try{
    const s=eightMemberSource_();let id=String(m.memberId||'').trim();if(!id)id=eightGenerateMemberId_(s);
    if(!/^A\d{5}$/i.test(id)&&!m.allowLegacyId)throw new Error('MEMBER_ID_FORMAT');id=id.toUpperCase();
    const name=String(m.name||'').trim();if(!name)throw new Error('MEMBER_NAME_REQUIRED');eightAssertUniqueMemberId_(s,id,0);
    const row=new Array(s.h.length).fill(''),put=(c,v)=>{if(c>=0)row[c]=v==null?'':v};put(s.c.ts,new Date());put(s.c.name,name);put(s.c.birth,m.birth);put(s.c.nationalId,m.nationalId);put(s.c.phone,m.phone);put(s.c.id,id);put(s.c.address,m.address);s.sh.appendRow(row);
    const key='M-'+Utilities.getUuid(),now=new Date();eightMeta_().appendRow([key,id,name,m.nickname||'',m.group||'','active',now,now]);
    return {memberKey:key,memberId:id,name,status:'active'}
  }finally{lock.releaseLock()}
}
function eightUpdateMember_(key,p){
  const lock=LockService.getScriptLock();lock.waitLock(10000);try{
    const f=eightFindMember_(key),s=eightMemberSource_(),rows=eightFindSourceRowsById_(s,f.id);if(rows.length!==1)throw new Error(rows.length?'MEMBER_SOURCE_DUPLICATE_ID':'MEMBER_NOT_FOUND');const row=rows[0];
    let newId=f.id;if('memberId' in p){newId=String(p.memberId||'').trim().toUpperCase();if(!newId)throw new Error('MEMBER_ID_REQUIRED');if(!/^A\d{5}$/i.test(newId)&&!p.allowLegacyId)throw new Error('MEMBER_ID_FORMAT');eightAssertUniqueMemberId_(s,newId,row)}
    const map={name:s.c.name,birth:s.c.birth,nationalId:s.c.nationalId,phone:s.c.phone,memberId:s.c.id,address:s.c.address};Object.keys(map).forEach(k=>{if(k in p&&map[k]>=0)s.sh.getRange(row,map[k]+1).setValue(k==='memberId'?newId:p[k])});
    const meta=eightMeta_();if(newId!==f.id)meta.getRange(f.metaRow,2).setValue(newId);if('name' in p)meta.getRange(f.metaRow,3).setValue(p.name);if('nickname' in p)meta.getRange(f.metaRow,4).setValue(p.nickname);if('group' in p)meta.getRange(f.metaRow,5).setValue(p.group);if('status' in p)meta.getRange(f.metaRow,6).setValue(p.status);meta.getRange(f.metaRow,8).setValue(new Date());
    return {memberKey:key,memberId:newId}
  }finally{lock.releaseLock()}
}
function eightDeleteMember_(key){
  const lock=LockService.getScriptLock();lock.waitLock(10000);try{const f=eightFindMember_(key),s=eightMemberSource_(),rows=eightFindSourceRowsById_(s,f.id);if(rows.length!==1)throw new Error(rows.length?'MEMBER_SOURCE_DUPLICATE_ID':'MEMBER_NOT_FOUND');s.sh.deleteRow(rows[0]);eightMeta_().deleteRow(f.metaRow);return {deleted:true}}finally{lock.releaseLock()}
}
function eightListEvents_(businessDate){
  const sh=eightEvents_();if(sh.getLastRow()<2)return [];const v=sh.getRange(2,1,sh.getLastRow()-1,19).getDisplayValues();
  const events=v.filter(r=>r[0]&&(!businessDate||r[1]===businessDate)&&r[9]!=='deleted').map(r=>({eventId:r[0],businessDate:r[1],name:r[2],startTime:r[3],regClose:r[4],level:r[5],buyin:Number(r[6]||0),fee:Number(r[7]||0),jpRate:Number(r[8]||0),status:r[9]||'open',buyinTotal:Number(r[12]||Number(r[6]||0)+Number(r[7]||0)),buyinAdmin:Number(r[13]||r[7]||0),rebuyTotal:Number(r[14]||Number(r[6]||0)+Number(r[7]||0)),rebuyAdmin:Number(r[15]||r[7]||0),freeAdminFrom:Number(r[16]||11),icmRate:Number(r[17]||3),icmRound:Number(r[18]||100)}));const psh=eightPlayers_(),pv=psh.getLastRow()>1?psh.getRange(2,1,psh.getLastRow()-1,18).getDisplayValues():[];events.forEach(e=>{const ps=pv.filter(r=>r[0]===e.eventId&&r[8]!=='deleted');let totalEntries=0,rebuyPeople=0,early=0,late=0,rebuyDisc=0,entryDisc=0,other=0;ps.forEach(r=>{const b=Math.max(0,Number(r[16]||1)),rb=Math.max(0,Number(r[17]||Math.max(0,Number(r[4]||1)-1))),n=b+rb;totalEntries+=n;if(rb>0)rebuyPeople++;early+=Number(r[11]||0);late+=Number(r[12]||0);rebuyDisc+=Number(r[13]||0);const autoRebuy=Math.max(0,rb)*e.rebuyAdmin/2,autoOverbuy=Math.max(0,n-10)*e.rebuyAdmin/2;rebuyDisc+=autoRebuy;entryDisc+=autoOverbuy;other+=Number(r[15]||0)});const totalGross=ps.length*e.buyinTotal+Math.max(0,totalEntries-ps.length)*e.rebuyTotal;const prizeBase=ps.length*Math.max(0,e.buyinTotal-e.buyinAdmin)+Math.max(0,totalEntries-ps.length)*Math.max(0,e.rebuyTotal-e.rebuyAdmin);const unit=Math.max(1,e.icmRound||100),prizePool=Math.floor((prizeBase*(1-(e.icmRate||0)/100))/unit)*unit;const jp=Math.floor(prizeBase*(e.jpRate||0)/100);const discounts=early+late+rebuyDisc+entryDisc+other;const adminGross=ps.length*e.buyinAdmin+Math.max(0,totalEntries-ps.length)*e.rebuyAdmin;e.summary={participants:ps.length,rebuyPeople,totalEntries,totalGross,earlyDiscount:early,lateDiscount:late,rebuyDiscount:rebuyDisc,entryDiscount:entryDisc,otherDiscount:other,prizePool,adminNet:Math.max(0,adminGross-discounts),jp}});return events
}
function eightCreateEvent_(e){
  const lock=LockService.getScriptLock();if(!lock.tryLock(3000))throw new Error('SYSTEM_BUSY_RETRY');
  try{
    if(!String(e.name||'').trim())throw new Error('EVENT_NAME_REQUIRED');
    if(!/^\d{4}-\d{2}-\d{2}$/.test(String(e.businessDate||'')))throw new Error('EVENT_DATE_REQUIRED');
    const id='E-'+Utilities.getUuid(),now=new Date(),sh=eightEvents_(),row=Math.max(2,sh.getLastRow()+1);
    sh.getRange(row,1,1,19).setValues([[id,e.businessDate,String(e.name).trim(),e.startTime||'',e.regClose||'',e.level||'custom',Number(e.buyin||0),Number(e.fee||0),Number(e.jpRate||0),'open',now,now,Number(e.buyinTotal||0),Number(e.buyinAdmin||0),Number(e.rebuyTotal||0),Number(e.rebuyAdmin||0),Number(e.freeAdminFrom||11),Number(e.icmRate||3),Number(e.icmRound||100)]]);
    SpreadsheetApp.flush();
    return {eventId:id,businessDate:e.businessDate,name:String(e.name||'').trim(),startTime:e.startTime||'',regClose:e.regClose||'',level:e.level||'custom',buyin:Number(e.buyin||0),fee:Number(e.fee||0),jpRate:Number(e.jpRate||0),buyinTotal:Number(e.buyinTotal||0),buyinAdmin:Number(e.buyinAdmin||0),rebuyTotal:Number(e.rebuyTotal||0),rebuyAdmin:Number(e.rebuyAdmin||0),freeAdminFrom:Number(e.freeAdminFrom||11),icmRate:Number(e.icmRate||3),icmRound:Number(e.icmRound||100),status:'open'}
  }finally{lock.releaseLock()}
}
function eightFindEventRow_(id){const sh=eightEvents_(),v=sh.getDataRange().getDisplayValues();for(let i=1;i<v.length;i++)if(v[i][0]===id)return {sh,row:i+1,data:v[i]};throw new Error('EVENT_NOT_FOUND')}
function eightUpdateEvent_(id,p){const lock=LockService.getScriptLock();lock.waitLock(10000);try{const f=eightFindEventRow_(id),map={businessDate:2,name:3,startTime:4,regClose:5,level:6,buyin:7,fee:8,jpRate:9,status:10,buyinTotal:13,buyinAdmin:14,rebuyTotal:15,rebuyAdmin:16,freeAdminFrom:17,icmRate:18,icmRound:19};Object.keys(map).forEach(k=>{if(k in p)f.sh.getRange(f.row,map[k]).setValue(p[k])});f.sh.getRange(f.row,12).setValue(new Date());return {eventId:id}}finally{lock.releaseLock()}}
function eightDeleteEvent_(id){const lock=LockService.getScriptLock();lock.waitLock(10000);try{const f=eightFindEventRow_(id);f.sh.getRange(f.row,10).setValue('deleted');f.sh.getRange(f.row,12).setValue(new Date());return {deleted:true}}finally{lock.releaseLock()}}
function eightListEventPlayers_(eventId){
  if(!eventId)return [];const sh=eightPlayers_();if(sh.getLastRow()<2)return [];return sh.getRange(2,1,sh.getLastRow()-1,18).getDisplayValues().filter(r=>r[0]===eventId&&r[8]!=='deleted').map(r=>({eventId:r[0],memberKey:r[1],memberId:r[2],name:r[3],entries:Number(r[4]||1),discount:Number(r[5]||0),group:r[6]||'',chips:Number(r[7]||0),status:r[8]||'active',earlyDiscount:Number(r[11]||0),lateDiscount:Number(r[12]||0),rebuyDiscount:Number(r[13]||0),entryDiscount:Number(r[14]||0),otherDiscount:Number(r[15]||0),buyin:Number(r[16]||1),rebuy:Number(r[17]||Math.max(0,Number(r[4]||1)-1))}))
}
function eightAddEventPlayer_(eventId,memberKey){
  const lock=LockService.getScriptLock();lock.waitLock(10000);try{eightFindEventRow_(eventId);const f=eightFindMember_(memberKey),existing=eightListEventPlayers_(eventId);if(existing.some(x=>x.memberKey===memberKey))throw new Error('PLAYER_ALREADY_IN_EVENT');const s=eightMemberSource_(),rows=eightFindSourceRowsById_(s,f.id);if(rows.length!==1)throw new Error('MEMBER_NOT_FOUND');const src=s.v[rows[0]-1],name=s.c.name>=0?src[s.c.name]:'',now=new Date();eightPlayers_().appendRow([eventId,memberKey,f.id,name,1,0,f.meta[4]||'',0,'active',now,now,0,0,0,0,0,1,0]);return {eventId,memberKey,memberId:f.id,name,entries:1,buyin:1,rebuy:0,discount:0,group:f.meta[4]||'',chips:0}}finally{lock.releaseLock()}
}
function eightFindPlayerRow_(eventId,key){const sh=eightPlayers_(),v=sh.getDataRange().getDisplayValues();for(let i=1;i<v.length;i++)if(v[i][0]===eventId&&v[i][1]===key&&v[i][8]!=='deleted')return {sh,row:i+1};throw new Error('EVENT_PLAYER_NOT_FOUND')}
function eightUpdateEventPlayer_(eventId,key,p){
  const lock=LockService.getScriptLock();lock.waitLock(10000);try{
    const f=eightFindPlayerRow_(eventId,key),sh=f.sh,row=f.row;
    const cur=sh.getRange(row,1,1,18).getValues()[0];
    let buyin=('buyin' in p)?Number(p.buyin||0):Number(cur[16]||1);
    let rebuy=('rebuy' in p)?Number(p.rebuy||0):Number(cur[17]||0);
    if(buyin<0||rebuy<0)throw new Error('INVALID_NUMBER');
    if('buyin' in p||'rebuy' in p||'entries' in p){
      cur[16]=buyin;cur[17]=rebuy;cur[4]=buyin+rebuy;
    }
    const map={discount:5,group:6,chips:7,status:8,earlyDiscount:11,lateDiscount:12,rebuyDiscount:13,entryDiscount:14,otherDiscount:15};
    Object.keys(map).forEach(k=>{if(k in p){let v=p[k];if(['discount','chips','earlyDiscount','lateDiscount','rebuyDiscount','entryDiscount','otherDiscount'].includes(k)){v=Number(v||0);if(v<0)throw new Error('INVALID_NUMBER')}cur[map[k]]=v}});
    cur[10]=new Date();sh.getRange(row,1,1,18).setValues([cur]);
    SpreadsheetApp.flush();
    return {eventId,memberKey:key,buyin,rebuy,entries:buyin+rebuy}
  }finally{lock.releaseLock()}
}
function eightDeleteEventPlayer_(eventId,key){const lock=LockService.getScriptLock();if(!lock.tryLock(3000))throw new Error('SYSTEM_BUSY_RETRY');try{const f=eightFindPlayerRow_(eventId,key);f.sh.deleteRow(f.row);SpreadsheetApp.flush();return {deleted:true}}finally{lock.releaseLock()}}
function eightBootstrap_(){
  const members=eightListMembers_(),settings=eightSettings_(),now=new Date();let monthNew=0;
  members.forEach(m=>{const d=new Date(m.timestamp);if(!isNaN(d)&&d.getFullYear()===now.getFullYear()&&d.getMonth()===now.getMonth())monthNew++});
  return {ok:true,settings:settings,members:members,summary:{memberCount:members.length,monthNewMembers:monthNew},apiVersion:'V1.5'}
}
function eightSaveAllEventPlayers_(eventId,players){
  const lock=LockService.getScriptLock();lock.waitLock(15000);try{
    eightFindEventRow_(eventId);const sh=eightPlayers_(),data=sh.getDataRange().getValues(),byKey=new Map();
    for(let i=1;i<data.length;i++)if(String(data[i][0])===String(eventId)&&data[i][8]!=='deleted')byKey.set(String(data[i][1]),i+1);
    players.forEach(p=>{const row=byKey.get(String(p.memberKey));if(!row)return;const cur=sh.getRange(row,1,1,18).getValues()[0];
      const b=Math.max(0,Number(p.buyin??cur[16]??1)),rb=Math.max(0,Number(p.rebuy??cur[17]??0));cur[4]=b+rb;cur[16]=b;cur[17]=rb;
      if('earlyDiscount' in p)cur[11]=Number(p.earlyDiscount||0);if('lateDiscount' in p)cur[12]=Number(p.lateDiscount||0);
      if('otherDiscount' in p)cur[15]=Number(p.otherDiscount||0);if('group' in p)cur[6]=p.group||'';if('chips' in p)cur[7]=Number(p.chips||0);
      cur[10]=new Date();sh.getRange(row,1,1,18).setValues([cur]);
    });SpreadsheetApp.flush();return eightListEventPlayers_(eventId)
  }finally{lock.releaseLock()}
}
