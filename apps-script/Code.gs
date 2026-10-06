/**
 * Eight Poker Management - Apps Script API V1
 * Spreadsheet ID is configured in Script Properties as DB_ID.
 */
const MEMBER_SHEET = 'Form Responses 1';
const SETTINGS_SHEET = 'EPC_系統設定';
const MEMBER_META_SHEET = 'EIGHT_MEMBER_META';

function doGet() {
  return respond({ok:true, service:'Eight Poker API', version:'1.0'});
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    // Eight Poker V1 authoritative API. Keep legacy actions below for migration only.
    if (String(req.action || '').indexOf('eight.') === 0) return respond(eightPokerApi_(req));
    switch (req.action) {
      case 'bootstrap': return respond(bootstrap());
      case 'members.meta.init': return respond(initMemberMeta());
      case 'members.list': return respond({ok:true, members:listMembers()});
      case 'members.create': return respond({ok:true, member:createMember(req.member || {})});
      case 'members.update': return respond({ok:true, member:updateMember(req.memberKey, req.patch || {})});
      case 'members.deactivate': return respond({ok:true, result:setMemberStatus(req.memberKey, 'inactive')});
      default: throw new Error('UNKNOWN_ACTION');
    }
  } catch (err) {
    return respond({ok:false, error:String(err && err.message || err)});
  }
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function db() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (!id) throw new Error('DB_ID_NOT_CONFIGURED');
  return SpreadsheetApp.openById(id);
}

function columnIndex(headers, names) {
  for (const name of names) {
    const i = headers.indexOf(name);
    if (i >= 0) return i;
  }
  return -1;
}

function getSettings() {
  const sh = db().getSheetByName(SETTINGS_SHEET);
  const rows = sh ? sh.getDataRange().getDisplayValues() : [];
  const out = {businessStart:'16:00', businessEnd:'07:00'};
  for (let i=1; i<rows.length; i++) if (rows[i][0]) out[rows[i][0]] = rows[i][1];
  return out;
}

function ensureMemberColumns(sh) {
  let headers = sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getDisplayValues()[0];
  ['綽號','分帳群組','MEMBER_KEY','會員狀態'].forEach(name => {
    if (!headers.includes(name)) {
      sh.getRange(1, sh.getLastColumn()+1).setValue(name);
      headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
    }
  });
  return headers;
}

function listMembers() {
  const sh = db().getSheetByName(MEMBER_SHEET);
  if (!sh) throw new Error('MEMBER_SHEET_NOT_FOUND');
  const rows = sh.getDataRange().getDisplayValues();
  const h = rows[0] || [];
  const c = {
    ts:columnIndex(h,['Timestamp','時間戳記']),
    name:columnIndex(h,['姓名']),
    id:columnIndex(h,['POKER FANS ID','POKERFANS ID']),
    nick:columnIndex(h,['綽號']),
    group:columnIndex(h,['分帳群組']),
    key:columnIndex(h,['MEMBER_KEY']),
    status:columnIndex(h,['會員狀態'])
  };
  return rows.slice(1).map((r,i) => {
    if (!r.some(Boolean)) return null;
    const memberId = c.id >= 0 ? r[c.id] : '';
    return {
      memberKey:c.key >= 0 && r[c.key] ? r[c.key] : 'FORM-'+(i+2)+'-'+memberId,
      memberId,
      name:c.name >= 0 ? r[c.name] : '',
      nickname:c.nick >= 0 ? r[c.nick] : '',
      group:c.group >= 0 ? r[c.group] : '',
      timestamp:c.ts >= 0 ? r[c.ts] : '',
      status:c.status >= 0 && r[c.status] ? r[c.status] : 'active',
      eventCount:0,totalEntries:0,pnl:0,spendShare:'',lastVisit:''
    };
  }).filter(Boolean);
}

function findMemberRow(sh, key) {
  const h = ensureMemberColumns(sh);
  const keyCol = h.indexOf('MEMBER_KEY');
  const idCol = columnIndex(h,['POKER FANS ID','POKERFANS ID']);
  const rows = sh.getDataRange().getDisplayValues();
  for (let i=1; i<rows.length; i++) {
    const fallback = 'FORM-'+(i+1)+'-'+(idCol >= 0 ? rows[i][idCol] : '');
    if (rows[i][keyCol] === key || fallback === key) return i+1;
  }
  throw new Error('MEMBER_NOT_FOUND');
}

function createMember(m) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = db().getSheetByName(MEMBER_SHEET);
    const h = ensureMemberColumns(sh);
    const row = new Array(h.length).fill('');
    const key = 'M-'+Utilities.getUuid();
    const put = (names,val) => { const i=columnIndex(h,names); if(i>=0) row[i]=val == null ? '' : val; };
    put(['Timestamp','時間戳記'],new Date());
    put(['姓名'],m.name); put(['出生年月日'],m.birth);
    put(['身分證字號'],m.nationalId); put(['手機號碼'],m.phone);
    put(['POKER FANS ID','POKERFANS ID'],m.memberId); put(['地址'],m.address);
    put(['綽號'],m.nickname); put(['分帳群組'],m.group);
    put(['MEMBER_KEY'],key); put(['會員狀態'],'active');
    sh.appendRow(row);
    return {memberKey:key, memberId:m.memberId || '', name:m.name || '', status:'active'};
  } finally { lock.releaseLock(); }
}

function updateMember(key, patch) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = db().getSheetByName(MEMBER_SHEET);
    const h = ensureMemberColumns(sh);
    const row = findMemberRow(sh,key);
    const map = {
      name:['姓名'], birth:['出生年月日'], nationalId:['身分證字號'],
      phone:['手機號碼'], memberId:['POKER FANS ID','POKERFANS ID'],
      address:['地址'], nickname:['綽號'], group:['分帳群組'], status:['會員狀態']
    };
    Object.keys(patch).forEach(k => {
      if (!map[k]) return;
      const c = columnIndex(h,map[k]);
      if (c >= 0) sh.getRange(row,c+1).setValue(patch[k]);
    });
    return {memberKey:key};
  } finally { lock.releaseLock(); }
}

function setMemberStatus(key,status) {
  return updateMember(key,{status});
}

function bootstrap() {
  const members = listMembers();
  const active = members.filter(m => m.status !== 'inactive');
  const now = new Date();
  const monthNew = active.filter(m => {
    const d = new Date(m.timestamp);
    return !isNaN(d) && d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth();
  }).length;
  return {ok:true, settings:getSettings(), members:active,
    summary:{memberCount:active.length, monthNewMembers:monthNew}};
}

/**
 * One-time, idempotent member metadata migration.
 * Keeps Google Form response columns untouched.
 */
function initMemberMeta() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = db();
    const memberSh = ss.getSheetByName(MEMBER_SHEET);
    if (!memberSh) throw new Error('MEMBER_SHEET_NOT_FOUND');

    let metaSh = ss.getSheetByName(MEMBER_META_SHEET);
    if (!metaSh) metaSh = ss.insertSheet(MEMBER_META_SHEET);

    const headers = ['MEMBER_KEY','POKER_FANS_ID','姓名快照','綽號','分帳群組','會員狀態','建立時間','更新時間'];
    if (metaSh.getLastRow() === 0) metaSh.getRange(1,1,1,headers.length).setValues([headers]);
    else {
      const current = metaSh.getRange(1,1,1,Math.max(metaSh.getLastColumn(),headers.length)).getDisplayValues()[0];
      headers.forEach((h,i) => { if (current[i] !== h) metaSh.getRange(1,i+1).setValue(h); });
    }

    const source = memberSh.getDataRange().getDisplayValues();
    const h = source[0] || [];
    const idCol = columnIndex(h,['POKER FANS ID','POKERFANS ID']);
    const nameCol = columnIndex(h,['姓名']);
    if (idCol < 0) throw new Error('MEMBER_ID_COLUMN_NOT_FOUND');

    const existingRows = metaSh.getLastRow() > 1
      ? metaSh.getRange(2,1,metaSh.getLastRow()-1,headers.length).getDisplayValues() : [];
    const byId = new Map();
    existingRows.forEach((r,i) => {
      const id = String(r[1] || '').trim();
      if (id && !byId.has(id)) byId.set(id,{row:i+2,key:r[0]});
    });

    const append = [];
    let existing = 0, skipped = 0;
    const now = new Date();
    for (let i=1; i<source.length; i++) {
      const row = source[i];
      if (!row.some(Boolean)) continue;
      const id = String(row[idCol] || '').trim();
      if (!id) { skipped++; continue; }
      if (byId.has(id)) { existing++; continue; }
      const key = 'M-'+Utilities.getUuid();
      append.push([key,id,nameCol>=0?row[nameCol]:'','','','active',now,now]);
      byId.set(id,{row:null,key:key});
    }
    if (append.length) metaSh.getRange(metaSh.getLastRow()+1,1,append.length,headers.length).setValues(append);
    metaSh.setFrozenRows(1);
    return {ok:true,sheet:MEMBER_META_SHEET,created:append.length,existing:existing,skipped:skipped,totalMeta:metaSh.getLastRow()-1};
  } finally { lock.releaseLock(); }
}

function readMemberMetaMap_() {
  const sh = db().getSheetByName(MEMBER_META_SHEET);
  const map = new Map();
  if (!sh || sh.getLastRow() < 2) return map;
  const rows = sh.getRange(2,1,sh.getLastRow()-1,Math.max(8,sh.getLastColumn())).getDisplayValues();
  rows.forEach(r => {
    const id=String(r[1]||'').trim();
    if(id) map.set(id,{memberKey:r[0],nickname:r[3]||'',group:r[4]||'',status:r[5]||'active'});
  });
  return map;
}
