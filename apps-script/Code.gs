const SHEET_ID = '1k_7hXkJoDP8nu9kiFkE-NxZhJf4dPQI40p7gBx6RiuI';



const DATA_SHEET = 'EPC_DATA';



const FORM_SHEET = 'Form Responses 1';



const CHUNK_SIZE = 45000;















function doGet(e) {



  try {



    const action = String((e && e.parameter && e.parameter.action) || 'ping').trim();







    if (action === 'ping') {



      return json_({



        ok: true,



        service: 'EPC MAIN API',



        version: '1.2.4',



        mode: 'READ_WRITE'



      });



    }







    if (action === 'load') {



      return json_(loadData_());



    }







    if (action === 'revision') {

      return json_(loadRevision_());

    }



    if (action === 'members') {



      const loaded = loadData_();



      return json_({



        ok: true,



        members: loaded.members || [],



        revision: loaded.revision,



        updatedAt: loaded.updatedAt,



        writeId: loaded.writeId



      });



    }







    return json_({ok:false, error:'UNKNOWN_ACTION'});



  } catch (err) {



    return json_({ok:false, error:errorText_(err)});



  }



}







function doPost(e) {



  try {



    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');



    const action = String(body.action || '').trim();







    if (action === 'save') {



      // V1.2.1: multi-device clients send baseData. Route those saves

      // through the three-way merge path while keeping legacy save clients compatible.

      if (body.baseData && typeof body.baseData === 'object') {

        return json_(mergeSaveData_(body));

      }

      return json_(saveData_(body));



    }



    if (action === 'mergeSave') {



      return json_(mergeSaveData_(body));



    }







    if (action === 'addMember') {



      return json_(addMember_(body));



    }
    if (action === 'updateMember') {
      return json_(updateMember_(body));
    }









    if (action === 'deleteMember') {



      return json_(deleteMember_(body));



    }







    return json_({ok:false, error:'UNKNOWN_ACTION'});



  } catch (err) {



    return json_({ok:false, error:errorText_(err)});



  }



}















function loadRevision_() {

  const sh = getDataSheet_();

  const row = sh.getRange(2, 3, 1, 3).getValues()[0];

  return {

    ok: true,

    revision: Number(row[0]) || 0,

    updatedAt: dateText_(row[1]),

    writeId: String(row[2] || '')

  };

}





function loadData_() {
  // Fast path: one cache entry avoids repeatedly rebuilding the same ~240KB EPC_DB
  // and re-reading/merging the entire member form on every page load.
  const cache = CacheService.getScriptCache();
  const cached = cache.get('EPC_LOAD_V2');
  if (cached) {
    try { return JSON.parse(cached); } catch (ignore) {}
  }

  const sh = getDataSheet_();

  const state = readState_(sh);







  let data;



  try {



    data = JSON.parse(state.raw);



  } catch (err) {



    throw new Error('EPC_DB_INVALID_JSON');



  }







  if (!data || typeof data !== 'object') {



    throw new Error('EPC_DB_INVALID_DATA');



  }



  // V1.1.4: merge current Form Responses members into every load response.

  // This changes only the response object; it does not write EPC_DATA.

  data.members = mergeFormMembersForRead_(

    Array.isArray(data.members) ? data.members : []

  );







  const response = {
    ok: true,
    data: data,
    members: Array.isArray(data.members) ? data.members : [],
    revision: state.revision,
    updatedAt: state.updatedAt,
    writeId: state.writeId,
    chunks: state.chunkCount
  };
  // Apps Script cache has a per-value size ceiling. Cache only when payload fits.
  try {
    const txt = JSON.stringify(response);
    if (txt.length < 95000) cache.put('EPC_LOAD_V2', txt, 60);
  } catch (ignore) {}
  return response;
}







function readState_(sh) {



  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('EPC_DATA_EMPTY');

  // EPC_DATA is a compact chunk table. Bound the scan to avoid accidental
  // formatting/old rows making every startup read much larger than necessary.
  const scanRows = Math.min(Math.max(lastRow - 1, 1), 100);
  const rows = sh.getRange(2, 1, scanRows, 5).getValues();



  const chunks = [];







  rows.forEach(function(row, offset) {



    const key = String(row[0] || '').trim();



    let index = null;







    if (key === 'EPC_DB') {



      index = 0;



    } else {



      const m = key.match(/^EPC_DB_(\d+)$/);



      if (m) index = Number(m[1]);



    }







    if (index !== null) {



      chunks.push({



        row: offset + 2,



        index: index,



        json: String(row[1] == null ? '' : row[1]),



        revision: Number(row[2]) || 0,



        updatedAt: dateText_(row[3]),



        writeId: String(row[4] || '')



      });



    }



  });







  if (!chunks.length) throw new Error('EPC_DB_NOT_FOUND');







  chunks.sort(function(a,b){ return a.index - b.index; });







  for (let i = 0; i < chunks.length; i++) {



    if (chunks[i].index !== i) {



      throw new Error('EPC_DB_CHUNK_MISSING_' + i);



    }



  }







  const revision = chunks[0].revision;



  const writeId = chunks[0].writeId;



  const updatedAt = chunks[0].updatedAt;







  for (let i = 1; i < chunks.length; i++) {



    if (chunks[i].revision !== revision) {



      throw new Error('EPC_DB_REVISION_MISMATCH');



    }



    if (writeId && chunks[i].writeId && chunks[i].writeId !== writeId) {



      throw new Error('EPC_DB_WRITE_ID_MISMATCH');



    }



  }







  return {



    raw: chunks.map(function(c){ return c.json; }).join(''),



    revision: revision,



    updatedAt: updatedAt,



    writeId: writeId,



    chunkCount: chunks.length,



    firstRow: chunks[0].row



  };



}

















// V1.2.3 Multi-device merge save.

// The client sends the last cloud snapshot (baseData) plus its current local data.

// Under one ScriptLock we merge only the fields/records that changed locally into

// the newest server state, so another computer's unrelated changes are preserved.

function mergeSaveData_(body) {

  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    const sh = getDataSheet_();

    const current = readState_(sh);

    let remote;

    try { remote = JSON.parse(current.raw); } catch (err) { throw new Error('EPC_DB_INVALID_JSON'); }

    if (!body.data || typeof body.data !== 'object') throw new Error('INVALID_DATA');

    const local = body.data;

    const base = (body.baseData && typeof body.baseData === 'object') ? body.baseData : remote;

    const merged = merge3_(base, local, remote, 'root');

    const raw = JSON.stringify(merged);

    JSON.parse(raw);

    const parts = splitString_(raw, CHUNK_SIZE);

    if (!parts.length) throw new Error('EMPTY_DATA');

    const newRevision = current.revision + 1;

    const updatedAt = new Date().toISOString();

    const writeId = String(body.writeId || makeWriteId_()).trim();

    if (!writeId) throw new Error('INVALID_WRITE_ID');

    const rowCount = Math.max(current.chunkCount, parts.length);

    const values = [];

    for (let i=0;i<rowCount;i++) {

      if (i < parts.length) values.push([i===0?'EPC_DB':'EPC_DB_'+i, parts[i], newRevision, updatedAt, writeId]);

      else values.push(['','','','','']);

    }

    sh.getRange(current.firstRow,1,rowCount,5).setValues(values);

    SpreadsheetApp.flush();

    const verify = readState_(sh);

    if (verify.revision !== newRevision || verify.writeId !== writeId || verify.raw !== raw) throw new Error('CLOUD_WRITE_VERIFY_FAILED');

    return {ok:true, data:merged, revision:newRevision, updatedAt:updatedAt, writeId:writeId, chunks:parts.length, merged:true};

  } finally { lock.releaseLock(); }

}



function sameJson_(a,b){ try{return JSON.stringify(a)===JSON.stringify(b);}catch(e){return false;} }

function isPlainObject_(v){ return v && typeof v==='object' && !Array.isArray(v); }

function cloneJson_(v){ return v===undefined?undefined:JSON.parse(JSON.stringify(v)); }

function recordKey_(x){

  if(!x || typeof x!=='object' || Array.isArray(x)) return '';

  const keys=['id','memberId','playerId','eventId','reportId','writeId','date','key'];

  for(let i=0;i<keys.length;i++){ const k=keys[i]; if(x[k]!==undefined && x[k]!==null && String(x[k])!=='') return k+':'+String(x[k]); }

  return '';

}

function keyedArray_(arr){

  if(!Array.isArray(arr) || !arr.length) return false;

  for(let i=0;i<arr.length;i++) if(!recordKey_(arr[i])) return false;

  return true;

}

function mergeArray3_(base,local,remote,path){

  base=Array.isArray(base)?base:[]; local=Array.isArray(local)?local:[]; remote=Array.isArray(remote)?remote:[];

  if(!keyedArray_(base.concat(local,remote).filter(function(x){return x!==undefined;}))) {

    if(sameJson_(local,base)) return cloneJson_(remote);

    if(sameJson_(remote,base)) return cloneJson_(local);

    return cloneJson_(local); // same primitive/list field changed on both: local edit wins.

  }

  const bm={},lm={},rm={},order=[];

  function put(arr,map){arr.forEach(function(x){const k=recordKey_(x); if(!map[k]) order.push(k); map[k]=x;});}

  put(remote,rm); put(local,lm); put(base,bm);

  const seen={},out=[];

  order.forEach(function(k){

    if(seen[k])return; seen[k]=1;

    const b=bm[k], l=lm[k], r=rm[k];

    if(b!==undefined && l===undefined){

      // Local deletion only wins when remote record was not independently changed.

      if(r!==undefined && !sameJson_(r,b)) out.push(cloneJson_(r));

      return;

    }

    if(l!==undefined && r===undefined){

      if(b!==undefined && sameJson_(l,b)) return; // remote deletion

      out.push(cloneJson_(l)); return;

    }

    if(l===undefined && r!==undefined){out.push(cloneJson_(r));return;}

    if(l!==undefined && r!==undefined) out.push(merge3_(b,l,r,path+'['+k+']'));

  });

  return out;

}

function merge3_(base,local,remote,path){

  if(sameJson_(local,base)) return cloneJson_(remote);

  if(sameJson_(remote,base)) return cloneJson_(local);

  if(Array.isArray(local) || Array.isArray(remote) || Array.isArray(base)) return mergeArray3_(base,local,remote,path);

  if(isPlainObject_(local) && isPlainObject_(remote)){

    const b=isPlainObject_(base)?base:{}; const out={}; const keys={};

    Object.keys(b).forEach(function(k){keys[k]=1;}); Object.keys(local).forEach(function(k){keys[k]=1;}); Object.keys(remote).forEach(function(k){keys[k]=1;});

    Object.keys(keys).forEach(function(k){

      const hasB=Object.prototype.hasOwnProperty.call(b,k), hasL=Object.prototype.hasOwnProperty.call(local,k), hasR=Object.prototype.hasOwnProperty.call(remote,k);

      if(hasB && !hasL){ if(hasR && !sameJson_(remote[k],b[k])) out[k]=cloneJson_(remote[k]); return; }

      if(hasL && !hasR){ if(hasB && sameJson_(local[k],b[k])) return; out[k]=cloneJson_(local[k]); return; }

      if(!hasL && hasR){out[k]=cloneJson_(remote[k]);return;}

      if(hasL && hasR) out[k]=merge3_(hasB?b[k]:undefined,local[k],remote[k],path+'.'+k);

    });

    return out;

  }

  return cloneJson_(local); // same scalar changed on both: latest local edit wins.

}



function saveData_(body) {



  const lock = LockService.getScriptLock();



  lock.waitLock(30000);







  try {



    const sh = getDataSheet_();



    const current = readState_(sh);







    const clientRevision = Number(body.revision);



    const force = body.force === true;







    if (!force && clientRevision !== current.revision) {



      throw new Error(



        'CLOUD_CONFLICT:SERVER_REVISION_' +



        current.revision +



        ':CLIENT_REVISION_' +



        clientRevision



      );



    }







    if (!body.data || typeof body.data !== 'object') {



      throw new Error('INVALID_DATA');



    }







    const raw = JSON.stringify(body.data);



    // \u5BEB\u5165\u524D\u5148\u9A57\u8B49 JSON \u53EF\u5B8C\u6574\u9084\u539F.



    JSON.parse(raw);







    const parts = splitString_(raw, CHUNK_SIZE);



    if (!parts.length) throw new Error('EMPTY_DATA');







    const newRevision = current.revision + 1;



    const updatedAt = new Date().toISOString();



    const writeId = String(body.writeId || makeWriteId_()).trim();







    if (!writeId) throw new Error('INVALID_WRITE_ID');











    const rowCount = Math.max(current.chunkCount, parts.length);



    const values = [];







    for (let i = 0; i < rowCount; i++) {



      if (i < parts.length) {



        values.push([



          i === 0 ? 'EPC_DB' : 'EPC_DB_' + i,



          parts[i],



          newRevision,



          updatedAt,



          writeId



        ]);



      } else {



        values.push(['','','','','']);



      }



    }







    sh.getRange(current.firstRow, 1, rowCount, 5).setValues(values);



    SpreadsheetApp.flush();







    // \u5BEB\u5165\u5F8C\u7ACB\u5373\u7531 Sheet \u91CD\u65B0\u8B80\u56DE\u9A57\u8B49.



    const verify = readState_(sh);







    if (



      verify.revision !== newRevision ||



      verify.writeId !== writeId ||



      verify.raw !== raw



    ) {



      throw new Error('CLOUD_WRITE_VERIFY_FAILED');



    }







    return {



      ok: true,



      revision: newRevision,



      updatedAt: updatedAt,



      writeId: writeId,



      chunks: parts.length



    };



  } finally {



    lock.releaseLock();



  }



}















function addMember_(body) {

  const member = body.member || {};

  const requestedId = String(member.id || body.memberId || '').trim();

  const name = String(member.name || body.name || '').trim();



  if (!name) throw new Error('MEMBER_NAME_REQUIRED');



  const lock = LockService.getScriptLock();

  lock.waitLock(30000);



  try {

    const sh = getFormSheet_();

    const map = getFormColumnMap_(sh);

    const lastRow = sh.getLastRow();



    let existingIds = [];

    if (lastRow >= 2) {

      existingIds = sh.getRange(2, map.idCol, lastRow - 1, 1)

        .getDisplayValues()

        .map(function(r){ return String(r[0] || '').trim(); })

        .filter(Boolean);

    }



    let id = requestedId;

    if (!id) id = generateMemberId_(existingIds);



    if (existingIds.indexOf(id) !== -1) {

      throw new Error('MEMBER_EXISTS');

    }



    const width = Math.max(sh.getLastColumn(), map.maxCol);

    const row = new Array(width).fill('');



    if (map.timestampCol) row[map.timestampCol - 1] = new Date();

    row[map.nameCol - 1] = name;

    row[map.idCol - 1] = id;



    sh.appendRow(row);

    SpreadsheetApp.flush();



    return {

      ok: true,

      member: {

        id: id,

        name: name,

        createdAt: null,

        createdDate: new Date().toISOString(),

        source: 'Form Responses 1'

      }

    };

  } finally {

    lock.releaseLock();

  }

}



function updateMember_(body) {
  const member = body.member || {};
  const id = String(member.id || body.memberId || '').trim();
  const name = String(member.name || body.name || '').trim();

  if (!id) throw new Error('MEMBER_ID_REQUIRED');
  if (!name) throw new Error('MEMBER_NAME_REQUIRED');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = getFormSheet_();
    const map = getFormColumnMap_(sh);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return {ok:false, updated:false, memberId:id, error:'MEMBER_NOT_FOUND'};

    const ids = sh.getRange(2, map.idCol, lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim() !== id) continue;
      const row = i + 2;
      sh.getRange(row, map.nameCol).setValue(name);
      SpreadsheetApp.flush();

      const verifyId = String(sh.getRange(row, map.idCol).getDisplayValue() || '').trim();
      const verifyName = String(sh.getRange(row, map.nameCol).getDisplayValue() || '').trim();
      if (verifyId !== id || verifyName !== name) throw new Error('MEMBER_UPDATE_VERIFY_FAILED');

      return {ok:true, updated:true, member:{id:id, name:name, source:'Form Responses 1'}};
    }
    return {ok:false, updated:false, memberId:id, error:'MEMBER_NOT_FOUND'};
  } finally {
    lock.releaseLock();
  }
}


function generateMemberId_(existingIds) {

  const used = {};

  (existingIds || []).forEach(function(v) {

    const s = String(v || '').trim();

    if (s) used[s] = true;

  });



  for (let i = 0; i < 200; i++) {

    const id = String(Math.floor(100000 + Math.random() * 900000));

    if (!used[id]) return id;

  }



  throw new Error('MEMBER_ID_GENERATION_FAILED');

}





function deleteMember_(body) {



  const id = String(body.memberId || (body.member && body.member.id) || '').trim();



  if (!id) throw new Error('MEMBER_ID_REQUIRED');







  const lock = LockService.getScriptLock();



  lock.waitLock(30000);







  try {



    const sh = getFormSheet_();



    const map = getFormColumnMap_(sh);



    const lastRow = sh.getLastRow();







    if (lastRow < 2) {



      return {ok:true, deleted:false, memberId:id};



    }







    const ids = sh.getRange(2, map.idCol, lastRow - 1, 1).getDisplayValues();







    for (let i = ids.length - 1; i >= 0; i--) {



      if (String(ids[i][0] || '').trim() === id) {



        sh.deleteRow(i + 2);



        SpreadsheetApp.flush();



        return {ok:true, deleted:true, memberId:id};



      }



    }







    return {ok:true, deleted:false, memberId:id};



  } finally {



    lock.releaseLock();



  }



}











function getFormColumnMap_(sh) {

  const lastCol = sh.getLastColumn();

  if (lastCol < 1) throw new Error('FORM_HEADER_NOT_FOUND');



  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0]

    .map(function(v){ return String(v || '').trim(); });



  let timestampCol = 0;

  let nameCol = 0;

  let idCol = 0;



  headers.forEach(function(h, i) {

    const x = h.toLowerCase().replace(/\s+/g, '');



    if (!timestampCol && (

      x === 'timestamp' ||

      x.includes('\u6642\u9593\u6233') ||

      x.includes('\u6642\u9593\u6233\u8a18') ||

      x.includes('\u63d0\u4ea4\u6642\u9593')

    )) timestampCol = i + 1;



    if (!nameCol && (

      x === 'name' ||

      x === '\u59d3\u540d' ||

      x.includes('\u6703\u54e1\u59d3\u540d') ||

      x.includes('\u73a9\u5bb6\u59d3\u540d')

    )) nameCol = i + 1;



    if (!idCol && (

      x === 'id' ||

      x === 'memberid' ||

      x === '\u6703\u54e1id' ||

      x === '\u6703\u54e1\u7de8\u865f' ||

      x.includes('\u6703\u54e1\u7de8\u865f') ||

      x.includes('\u6703\u54e1id') ||

      x === 'pokerfansid' ||

      x.includes('pokerfansid')

    )) idCol = i + 1;

  });



  if (!nameCol || !idCol) {

    throw new Error(

      'FORM_COLUMNS_NOT_FOUND:HEADERS=' +

      headers.join('|').slice(0, 500)

    );

  }



  return {

    timestampCol: timestampCol,

    nameCol: nameCol,

    idCol: idCol,

    maxCol: Math.max(timestampCol || 0, nameCol, idCol)

  };

}





function mergeFormMembersForRead_(baseMembers) {

  const map = new Map();



  (baseMembers || []).forEach(function(m) {

    if (!m || typeof m !== 'object') return;

    const id = String(m.id || '').trim();

    if (!id) return;

    map.set(id, m);

  });



  try {

    const sh = getFormSheet_();

    const cols = getFormColumnMap_(sh);

    const lastRow = sh.getLastRow();

    if (lastRow < 2) return Array.from(map.values());



    const width = Math.max(sh.getLastColumn(), cols.maxCol);

    const rows = sh.getRange(2, 1, lastRow - 1, width).getValues();



    rows.forEach(function(row) {

      const id = String(row[cols.idCol - 1] || '').trim();

      const name = String(row[cols.nameCol - 1] || '').trim();

      if (!id || !name) return;



      const old = map.get(id) || {};

      const ts = cols.timestampCol ? row[cols.timestampCol - 1] : '';

      const createdDate = ts instanceof Date ? ts.toISOString() : String(ts || '');



      map.set(id, Object.assign({}, old, {

        id: id,

        name: name,

        group: old.group || '',

        note: old.note || '',

        history: Array.isArray(old.history) ? old.history : [],

        createdAt: old.createdAt || null,

        createdDate: old.createdDate || createdDate,

        source: 'Form Responses 1'

      }));

    });

  } catch (err) {

    // Member-form refresh must never make EPC_DATA unreadable.

    console.error('FORM_MEMBER_MERGE_FAILED', err);

  }



  return Array.from(map.values());

}





function getDataSheet_() {



  const ss = SpreadsheetApp.openById(SHEET_ID);



  const sh = ss.getSheetByName(DATA_SHEET);



  if (!sh) throw new Error('EPC_DATA_NOT_FOUND');



  return sh;



}







function getFormSheet_() {



  const ss = SpreadsheetApp.openById(SHEET_ID);



  const sh = ss.getSheetByName(FORM_SHEET);



  if (!sh) throw new Error('FORM_RESPONSES_NOT_FOUND');



  return sh;



}







function splitString_(text, size) {



  const out = [];



  for (let i = 0; i < text.length; i += size) {



    out.push(text.slice(i, i + size));



  }



  return out;



}







function makeWriteId_() {



  return 'EPC-' + Date.now() + '-' + Utilities.getUuid().slice(0, 8);



}







function dateText_(value) {



  if (value instanceof Date) return value.toISOString();



  return String(value || '');



}







function errorText_(err) {



  return String(err && err.message || err);



}







function json_(obj) {



  return ContentService



    .createTextOutput(JSON.stringify(obj))



    .setMimeType(ContentService.MimeType.JSON);



}
