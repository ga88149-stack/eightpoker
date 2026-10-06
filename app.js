const CONFIG={apiUrl:localStorage.getItem('eightApiUrl')||'https://script.google.com/macros/s/AKfycbwj-jLaSTiljcw15_UnhUFv4QxVq5HbNyers3ybU-6IAkwSj_A5X36x-852Xu5MxobS/exec',businessStart:localStorage.getItem('eightBusinessStart')||'16:00',businessEnd:localStorage.getItem('eightBusinessEnd')||'07:00'};
let MEMBER_ROWS=[];let MEMBER_PAGE=1;const MEMBER_PAGE_SIZE=100;let MEMBER_SEARCH_TIMER=null;
const pages={dashboard:'總覽',members:'會員資料',events:'賽事管理',settlement:'分帳報表',accounting:'帳務管理',activities:'活動專區',devices:'設備管理',settings:'系統設定'};
const pad=n=>String(n).padStart(2,'0'),money=n=>new Intl.NumberFormat('zh-TW').format(Number(n||0));
function localISO(d){return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())}
function businessDate(now=new Date(),start=CONFIG.businessStart,end=CONFIG.businessEnd){const [sh,sm]=start.split(':').map(Number),[eh,em]=end.split(':').map(Number),mins=now.getHours()*60+now.getMinutes(),s=sh*60+sm,e=eh*60+em,d=new Date(now);if(e<s&&mins<e)d.setDate(d.getDate()-1);return localISO(d)}
async function api(action,payload={}){
  if(!CONFIG.apiUrl)throw new Error('尚未設定 API URL');
  const ctrl=new AbortController(),isRead=['eight.ping','eight.bootstrap','eight.members.list','eight.events.list','eight.eventPlayers.list','eight.events.snapshot'].includes(action),timeout=action==='eight.bootstrap'?35000:(isRead?30000:20000),timer=setTimeout(()=>ctrl.abort(),timeout);
  try{
    const res=await fetch(CONFIG.apiUrl,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action,...payload}),signal:ctrl.signal});
    const raw=await res.text();let data;
    try{data=JSON.parse(raw)}catch(_){throw new Error('後端連線失敗（HTTP '+res.status+'）')}
    if(!data.ok)throw new Error(data.error||'API error');return data
  }catch(err){if(err&&err.name==='AbortError')throw new Error('後端逾時，請重試');throw err}
  finally{clearTimeout(timer)}
}
function setSync(t,bad=false){const e=document.querySelector('#syncState');e.textContent=t;e.style.color=bad?'var(--bad)':''}
function refreshBusinessDay(){CONFIG.businessStart=document.querySelector('#businessStart')?.value||CONFIG.businessStart;CONFIG.businessEnd=document.querySelector('#businessEnd')?.value||CONFIG.businessEnd;const d=businessDate();document.querySelector('#businessDayLabel').textContent='營業時間：每日 '+CONFIG.businessStart+'–翌日 '+CONFIG.businessEnd;document.querySelector('#todayDate').textContent=d;document.querySelector('#globalDate').value=d}
function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function memberSearchText(v){return String(v??'').normalize('NFKC').toLowerCase().replace(/\s+/g,'')}
function renderMembers(rows=MEMBER_ROWS){const q=memberSearchText(document.querySelector('#memberSearch')?.value||'');const filtered=q?rows.filter(m=>[m.memberId,m.name,m.nickname,m.group,m.phone].some(v=>memberSearchText(v).includes(q))):rows;const pages=Math.max(1,Math.ceil(filtered.length/MEMBER_PAGE_SIZE));MEMBER_PAGE=Math.min(MEMBER_PAGE,pages);const start=(MEMBER_PAGE-1)*MEMBER_PAGE_SIZE,list=filtered.slice(start,start+MEMBER_PAGE_SIZE);const tb=document.querySelector('#members tbody');const info=document.querySelector('#memberResultInfo');if(info)info.textContent='共 '+filtered.length+' 筆 · 第 '+MEMBER_PAGE+' / '+pages+' 頁';if(!list.length){tb.innerHTML='<tr><td colspan="10" class="empty">目前沒有符合的會員資料</td></tr>';return}tb.innerHTML=list.map(m=>'<tr data-key="'+esc(m.memberKey)+'"><td>'+esc(m.memberId)+'</td><td>'+esc(m.name)+'</td><td>'+esc(m.nickname)+'</td><td>'+esc(m.group)+'</td><td>'+money(m.eventCount)+'</td><td>'+money(m.totalEntries)+'</td><td>'+money(m.pnl)+'</td><td>'+esc(m.spendShare||'—')+'</td><td>'+esc(m.lastVisit||'—')+'</td><td class="row-actions"><button class="edit-member" data-key="'+esc(m.memberKey)+'">編輯</button><button class="delete-member danger" data-key="'+esc(m.memberKey)+'">刪除</button></td></tr>').join('');document.querySelector('#memberPrev').disabled=MEMBER_PAGE<=1;document.querySelector('#memberNext').disabled=MEMBER_PAGE>=pages}
function nextMemberId(){const used=new Set(MEMBER_ROWS.map(m=>String(m.memberId||'').trim().toUpperCase()));for(let i=0;i<100;i++){const id='A'+String(Math.floor(Math.random()*100000)).padStart(5,'0');if(!used.has(id))return id}return 'A'+String(Date.now()).slice(-5)}
function openMemberModal(member=null){document.querySelector('#memberForm').reset();document.querySelector('#memberKey').value=member?.memberKey||'';document.querySelector('#memberId').value=member?.memberId||nextMemberId();document.querySelector('#memberName').value=member?.name||'';document.querySelector('#memberNickname').value=member?.nickname||'';document.querySelector('#memberGroup').value=member?.group||'';document.querySelector('#memberBirth').value=member?.birth||'';document.querySelector('#memberPhone').value=member?.phone||'';document.querySelector('#memberAddress').value=member?.address||'';document.querySelector('#memberModalTitle').textContent=member?'編輯會員':'新增會員';document.querySelector('#memberFormState').textContent='';document.querySelector('#memberFormState').className='form-state';document.querySelector('#memberModal').hidden=false}
function closeMemberModal(){document.querySelector('#memberModal').hidden=true}
async function saveMember(e){e.preventDefault();const btn=document.querySelector('#memberSaveBtn'),state=document.querySelector('#memberFormState'),memberKey=document.querySelector('#memberKey').value;const data={memberId:document.querySelector('#memberId').value.trim(),name:document.querySelector('#memberName').value.trim(),nickname:document.querySelector('#memberNickname').value.trim(),group:document.querySelector('#memberGroup').value.trim(),birth:document.querySelector('#memberBirth').value.trim(),phone:document.querySelector('#memberPhone').value.trim(),address:document.querySelector('#memberAddress').value.trim()};if(memberKey){const original=MEMBER_ROWS.find(x=>x.memberKey===memberKey);['birth','phone','address'].forEach(k=>{if(original&&original[k]===undefined)delete data[k]})}if(!data.memberId||!data.name){state.textContent='POKER FANS ID 與姓名為必填';state.className='form-state bad';return}btn.disabled=true;state.textContent='儲存中…';try{if(memberKey)await api('eight.members.update',{memberKey,patch:data});else await api('eight.members.create',{member:data});state.textContent='儲存成功';state.className='form-state good';closeMemberModal();boot()}catch(err){console.error(err);const msg={MEMBER_ID_ALREADY_EXISTS:'POKER FANS ID 已存在',MEMBER_ID_REQUIRED:'POKER FANS ID 為必填',MEMBER_NAME_REQUIRED:'姓名為必填',MEMBER_NOT_FOUND:'找不到此會員'}[err.message]||err.message;state.textContent='儲存失敗：'+msg;state.className='form-state bad'}finally{btn.disabled=false}}
async function boot(){
 refreshBusinessDay();if(!CONFIG.apiUrl){setSync('資料庫：等待 Apps Script 部署');return}
 let cached=[];try{cached=JSON.parse(localStorage.getItem('eightMemberCache')||'[]')}catch(_){}
 if(cached.length){MEMBER_ROWS=cached;renderMembers();document.querySelector('#kMembers').textContent=money(cached.length);setSync('資料庫：同步中…')}
 loadEvents();
 try{
   const r=await api('eight.bootstrap');
   if(r.settings){CONFIG.businessStart=r.settings.businessStart||CONFIG.businessStart;CONFIG.businessEnd=r.settings.businessEnd||CONFIG.businessEnd;document.querySelector('#businessStart').value=CONFIG.businessStart;document.querySelector('#businessEnd').value=CONFIG.businessEnd}
   MEMBER_ROWS=r.members||[];try{localStorage.setItem('eightMemberCache',JSON.stringify(MEMBER_ROWS))}catch(_){}
   renderMembers();document.querySelector('#kMembers').textContent=money(r.summary?.memberCount);document.querySelector('#kNewMembers').textContent=money(r.summary?.monthNewMembers);refreshBusinessDay();setSync('資料庫：已連線')
 }catch(e){
   console.error(e)
   if(cached.length){MEMBER_ROWS=cached;renderMembers();document.querySelector('#kMembers').textContent=money(cached.length);setSync('資料庫：暫時無法連線（顯示上次資料）',true)}
   else setSync('資料庫：連線失敗',true)
 }
}
function goPage(page){document.querySelectorAll('#nav button').forEach(x=>x.classList.toggle('active',x.dataset.page===page));document.querySelectorAll('.page').forEach(x=>x.classList.toggle('active',x.id===page));const t=document.querySelector('#pageTitle');if(t&&pages[page])t.textContent=pages[page]}
document.querySelectorAll('#nav button').forEach(b=>b.addEventListener('click',()=>{goPage(b.dataset.page);if(b.dataset.page==='events')loadEvents()}));
document.querySelector('#themeBtn').addEventListener('click',()=>{const dark=document.body.dataset.theme==='dark';document.body.dataset.theme=dark?'light':'dark';document.querySelector('#themeBtn').textContent=dark?'☾ 深色模式':'☀ 一般模式';localStorage.setItem('eightTheme',document.body.dataset.theme)});
document.body.dataset.theme=localStorage.getItem('eightTheme')||'dark';
document.querySelector('#businessStart').addEventListener('change',refreshBusinessDay);document.querySelector('#businessEnd').addEventListener('change',refreshBusinessDay);
document.querySelector('#saveBusinessHours').addEventListener('click',async()=>{const state=document.querySelector('#businessSaveState');const start=document.querySelector('#businessStart').value,end=document.querySelector('#businessEnd').value;state.textContent='儲存中…';try{const r=await api('eight.settings.update',{settings:{businessStart:start,businessEnd:end}});CONFIG.businessStart=r.settings.businessStart;CONFIG.businessEnd=r.settings.businessEnd;localStorage.removeItem('eightBusinessStart');localStorage.removeItem('eightBusinessEnd');refreshBusinessDay();state.textContent='已同步到資料庫'}catch(err){state.textContent='儲存失敗：'+err.message}setTimeout(()=>state.textContent='',2200)});
const now=new Date(),first=new Date(now.getFullYear(),now.getMonth(),1);document.querySelector('#rangeFrom').value=localISO(first);document.querySelector('#rangeTo').value=localISO(now);
document.querySelector('#refreshBtn').addEventListener('click',boot);boot();
document.querySelector('#addMemberBtn').addEventListener('click',()=>openMemberModal());
document.querySelectorAll('[data-close-member]').forEach(x=>x.addEventListener('click',closeMemberModal));
document.querySelector('#memberForm').addEventListener('submit',saveMember);
document.querySelector('#memberSearch').addEventListener('input',()=>{clearTimeout(MEMBER_SEARCH_TIMER);MEMBER_SEARCH_TIMER=setTimeout(()=>{MEMBER_PAGE=1;renderMembers()},120)});
document.querySelector('#memberPrev').addEventListener('click',()=>{if(MEMBER_PAGE>1){MEMBER_PAGE--;renderMembers()}});
document.querySelector('#memberNext').addEventListener('click',()=>{MEMBER_PAGE++;renderMembers()});
document.querySelector('#members tbody').addEventListener('click',async e=>{const edit=e.target.closest('.edit-member');if(edit){const m=MEMBER_ROWS.find(x=>x.memberKey===edit.dataset.key);if(m)openMemberModal(m);return}const del=e.target.closest('.delete-member');if(!del)return;const m=MEMBER_ROWS.find(x=>x.memberKey===del.dataset.key);if(!m)return;if(!confirm('確定永久刪除會員「'+m.name+'」（'+m.memberId+'）？\n\n這會刪除會員主表與會員 META 資料。'))return;if(!confirm('再次確認：永久刪除後無法從會員頁復原。確定刪除？'))return;del.disabled=true;try{await api('eight.members.delete',{memberKey:m.memberKey});MEMBER_ROWS=MEMBER_ROWS.filter(x=>x.memberKey!==m.memberKey);renderMembers();document.querySelector('#kMembers').textContent=money(MEMBER_ROWS.filter(x=>x.status!=='inactive').length)}catch(err){const msg={MEMBER_SOURCE_DUPLICATE_ID:'此 POKER FANS ID 有重複資料，為避免刪錯已停止刪除',MEMBER_NOT_FOUND:'找不到此會員'}[err.message]||err.message;alert('刪除失敗：'+msg);del.disabled=false}});

const EVENT_PRESETS={3400:[3400,400,3400,400],6600:[6600,600,6600,600],11000:[11000,1000,11000,1000],21500:[21500,1500,21500,1500],32000:[32000,2000,32000,2000]};
function openEventModal(){document.querySelector('#eventForm').reset();const d=document.querySelector('#eventDate').value||businessDate();document.querySelector('#eventBusinessDate').value=d;document.querySelector('#eventLevel').value='3400';const p=EVENT_PRESETS['3400'];document.querySelector('#eventBuyinTotal').value=p[0];document.querySelector('#eventBuyinAdmin').value=p[1];document.querySelector('#eventRebuyTotal').value=p[2];document.querySelector('#eventRebuyAdmin').value=p[3];document.querySelector('#eventFreeAdminFrom').value=11;document.querySelector('#eventJP').value=3;document.querySelector('#eventICMRate').value=3;document.querySelector('#eventICMRound').value=100;const count=(window.EIGHT_EVENTS||[]).length+1;document.querySelector('#eventName').value='EPC#'+count+' 3400 限時錦標賽';document.querySelector('#eventFormState').textContent='';document.querySelector('#eventModal').hidden=false}
function closeEventModal(){document.querySelector('#eventModal').hidden=true}
document.querySelector('#eventDate').value=businessDate();
document.querySelector('#eventTodayBtn').addEventListener('click',()=>document.querySelector('#eventDate').value=businessDate());
document.querySelector('#createEventBtn').addEventListener('click',openEventModal);
document.querySelectorAll('[data-close-event]').forEach(x=>x.addEventListener('click',closeEventModal));
document.querySelector('#eventLevel').addEventListener('change',e=>{const p=EVENT_PRESETS[e.target.value];if(p){document.querySelector('#eventBuyinTotal').value=p[0];document.querySelector('#eventBuyinAdmin').value=p[1];document.querySelector('#eventRebuyTotal').value=p[2];document.querySelector('#eventRebuyAdmin').value=p[3]}const n=document.querySelector('#eventName'),count=(window.EIGHT_EVENTS||[]).length+1;if(/^EPC#\d+\s/.test(n.value)||!n.value.trim())n.value='EPC#'+count+' '+(e.target.value==='custom'?'自訂':e.target.value)+' 限時錦標賽'});
function renderEvents(rows=[]){window.EIGHT_EVENTS=rows;const el=document.querySelector('#eventList');if(!rows.length){el.className='empty';el.innerHTML='目前營業日尚無賽事';return}el.className='event-list';const stat=(k,v,moneyFmt=false)=>'<div class="event-stat"><small>'+k+'</small><b>'+(moneyFmt?money(v):esc(v??0))+'</b></div>';el.innerHTML=rows.map(x=>{const z=x.summary||{};return '<div class="event-row event-row-rich" data-event-id="'+esc(x.eventId)+'"><div class="event-main"><div><b>'+esc(x.name||'未命名賽事')+'</b><small>'+esc(x.businessDate||'')+' · '+esc(x.startTime||'')+' · '+esc(x.level||'自訂')+'</small></div><div class="event-actions"><button class="secondary enter-event" data-id="'+esc(x.eventId)+'">進入</button><button class="danger delete-event" data-id="'+esc(x.eventId)+'">刪除</button></div></div><div class="event-stats">'+stat('參賽人數',z.participants||0)+stat('重買人數',z.rebuyPeople||0)+stat('總組數',z.totalEntries||0)+stat('總買入',z.totalGross||0,true)+stat('早鳥',z.earlyDiscount||0,true)+stat('晚鳥',z.lateDiscount||0,true)+stat('重買優惠',z.rebuyDiscount||0,true)+stat('組數優惠',z.entryDiscount||0,true)+stat('其他優惠',z.otherDiscount||0,true)+stat('總獎金',z.prizePool||0,true)+stat('實收行政費',z.adminNet||0,true)+stat('JP',z.jp||0,true)+'</div></div>'}).join('')}
let EVENTS_CACHE_DATE='',EVENTS_LAST_SYNC=0;
async function loadEvents(){
 const date=document.querySelector('#eventDate').value||businessDate(),cacheKey='eightEvents:'+date;
 if(EVENTS_CACHE_DATE!==date){EVENTS_CACHE_DATE=date;let cached=[];try{cached=JSON.parse(localStorage.getItem(cacheKey)||'[]')}catch(_){}renderEvents(cached)}
 if(EVENTS_CACHE_DATE===date&&Date.now()-EVENTS_LAST_SYNC<30000&&(window.EIGHT_EVENTS||[]).length)return;
 try{const r=await api('eight.events.list',{businessDate:date}),rows=r.events||[];if(EVENTS_CACHE_DATE!==date)return;renderEvents(rows);EVENTS_LAST_SYNC=Date.now();try{localStorage.setItem(cacheKey,JSON.stringify(rows))}catch(_){}}
 catch(err){if(err.message!=='UNKNOWN_ACTION')console.error(err)}
}
document.querySelector('#eventDate').addEventListener('change',loadEvents);
document.querySelector('#eventForm').addEventListener('submit',async e=>{e.preventDefault();const state=document.querySelector('#eventFormState');const level=document.querySelector('#eventLevel').value,count=(window.EIGHT_EVENTS||[]).length+1;let eventName=document.querySelector('#eventName').value.trim();if(!eventName)eventName='EPC#'+count+' '+(level==='custom'?'自訂':level)+' 限時錦標賽';const buyinTotal=Number(document.querySelector('#eventBuyinTotal').value||0),buyinAdmin=Number(document.querySelector('#eventBuyinAdmin').value||0),rebuyTotal=Number(document.querySelector('#eventRebuyTotal').value||0),rebuyAdmin=Number(document.querySelector('#eventRebuyAdmin').value||0);const event={name:eventName,businessDate:document.querySelector('#eventBusinessDate').value,startTime:document.querySelector('#eventStartTime').value,regClose:document.querySelector('#eventRegClose').value,level,buyin:Math.max(0,buyinTotal-buyinAdmin),fee:buyinAdmin,buyinTotal,buyinAdmin,rebuyTotal,rebuyAdmin,freeAdminFrom:Number(document.querySelector('#eventFreeAdminFrom').value||11),jpRate:Number(document.querySelector('#eventJP').value||0),icmRate:Number(document.querySelector('#eventICMRate').value||0),icmRound:Number(document.querySelector('#eventICMRound').value||100)};state.textContent='建立中…';try{const created=await api('eight.events.create',{event});state.textContent='建立成功';state.className='form-state good';document.querySelector('#eventDate').value=event.businessDate;closeEventModal();if(created.event){window.EIGHT_EVENTS=[...(window.EIGHT_EVENTS||[]),created.event];renderEvents(window.EIGHT_EVENTS)}loadEvents()}catch(err){state.textContent='建立失敗：'+err.message;state.className='form-state bad'}});
let ACTIVE_EVENT=null,EVENT_PLAYERS=[];
function openPlayersModal(eventId){ACTIVE_EVENT=eventId;const ev=(window.EIGHT_EVENTS||[]).find(x=>x.eventId===eventId);document.querySelector('#eventPlayersTitle').textContent=ev?.name||'賽事玩家';document.querySelector('#eventPlayersMeta').textContent=(ev?.businessDate||'')+' '+(ev?.startTime||'');document.querySelector('#eventMemberSearch').value='';document.querySelector('#eventMemberMatches').innerHTML='';document.querySelector('#eventPlayersModal').hidden=false;loadEventPlayers()}
function closePlayersModal(){document.querySelector('#eventPlayersModal').hidden=true;ACTIVE_EVENT=null}
async function loadEventPlayers(){if(!ACTIVE_EVENT)return;try{const r=await api('eight.eventPlayers.list',{eventId:ACTIVE_EVENT});EVENT_PLAYERS=r.players||[];renderEventPlayers()}catch(err){EVENT_PLAYERS=[];renderEventPlayers()}}
function renderEventPlayers(){const el=document.querySelector('#eventPlayerList');document.querySelector('#eventPlayerCount').textContent=EVENT_PLAYERS.length+' 人';if(!EVENT_PLAYERS.length){el.className='empty';el.innerHTML='尚未加入玩家';return}el.className='player-list';el.innerHTML=EVENT_PLAYERS.map(p=>'<div class="player-row player-row-wide"><div><b>'+esc(p.name||'')+'</b><small>'+esc(p.memberId||'')+'</small></div><div class="player-fields"><label>組數<input class="entry-count" data-key="'+esc(p.memberKey)+'" type="number" min="1" value="'+esc(p.entries||1)+'"></label><label>早鳥<input class="early-discount" data-key="'+esc(p.memberKey)+'" type="number" min="0" value="'+esc(p.earlyDiscount||0)+'"></label><label>晚鳥<input class="late-discount" data-key="'+esc(p.memberKey)+'" type="number" min="0" value="'+esc(p.lateDiscount||0)+'"></label><label>重買優惠<input class="rebuy-discount" data-key="'+esc(p.memberKey)+'" type="number" min="0" value="'+esc(p.rebuyDiscount||0)+'"></label><label>組數優惠<input class="entry-discount" data-key="'+esc(p.memberKey)+'" type="number" min="0" value="'+esc(p.entryDiscount||0)+'"></label><label>其他優惠<input class="other-discount" data-key="'+esc(p.memberKey)+'" type="number" min="0" value="'+esc(p.otherDiscount||0)+'"></label><label>分帳群組<input class="settlement-group" data-key="'+esc(p.memberKey)+'" value="'+esc(p.group||'')+'"></label><label>籌碼<input class="chip-count" data-key="'+esc(p.memberKey)+'" type="number" min="0" value="'+esc(p.chips||0)+'"></label><button class="danger remove-player" data-key="'+esc(p.memberKey)+'">移除</button></div></div>').join('')}
document.querySelectorAll('[data-close-players]').forEach(x=>x.addEventListener('click',closePlayersModal));

let WORKSPACE_EVENT=null,WORKSPACE_PLAYERS=[];
function eventCalcPlayer(p,e){
 const buyin=Math.max(0,Number(p.buyin??1)),rebuy=Math.max(0,Number(p.rebuy??0)),groups=buyin+rebuy;
 const gross=buyin*Number(e.buyinTotal||0)+rebuy*Number(e.rebuyTotal||0);
 const early=Number(p.earlyDiscount||0),late=Number(p.lateDiscount||0),other=Number(p.otherDiscount||0);
 const rd=Math.max(0,rebuy)*Number(e.rebuyAdmin||0)/2;
 const overbuy=Math.max(0,groups-Math.max(0,Number(e.freeAdminFrom||11)-1))*Number(e.rebuyAdmin||0)/2;
 const discount=early+late+rd+overbuy+other;
 return {buyin,rebuy,groups,gross,early,late,rd,overbuy,other,discount,paid:Math.max(0,gross-discount)};
}
function showOnlyPage(id){goPage(id)}
async function openEventWorkspace(id){
 WORKSPACE_EVENT=(window.EIGHT_EVENTS||[]).find(x=>x.eventId===id);if(!WORKSPACE_EVENT)return;
 ACTIVE_EVENT=id;try{WORKSPACE_PLAYERS=JSON.parse(localStorage.getItem('eightEventPlayers:'+id)||'[]')}catch(_){WORKSPACE_PLAYERS=[]}
 showOnlyPage('eventWorkspace');document.querySelector('#workspaceTitle').textContent=WORKSPACE_EVENT.name;
 document.querySelector('#workspaceMeta').textContent=(WORKSPACE_EVENT.businessDate||'')+' · '+(WORKSPACE_EVENT.startTime||'')+' · '+(WORKSPACE_EVENT.level||'');
 document.querySelector('#workspaceMemberSearch').value='';document.querySelector('#workspaceGroup').value='';filterWorkspaceMembers('');renderWorkspace();
 loadWorkspacePlayers();
}
async function loadWorkspacePlayers(force=false){if(!ACTIVE_EVENT)return;const id=ACTIVE_EVENT;if(!force&&Date.now()-(WS_PLAYER_SYNC.get(id)||0)<30000&&WORKSPACE_PLAYERS.length)return;try{const r=await api('eight.eventPlayers.list',{eventId:id});if(ACTIVE_EVENT!==id)return;WORKSPACE_PLAYERS=r.players||[];WS_PLAYER_SYNC.set(id,Date.now());try{localStorage.setItem('eightEventPlayers:'+id,JSON.stringify(WORKSPACE_PLAYERS))}catch(_){}renderWorkspace();filterWorkspaceMembers(document.querySelector('#workspaceMemberSearch').value)}catch(err){if(ACTIVE_EVENT===id&&!WORKSPACE_PLAYERS.length)alert('載入玩家失敗：'+err.message)}}
function filterWorkspaceMembers(q){
 q=memberSearchText(q);const joined=new Set(WORKSPACE_PLAYERS.map(x=>x.memberKey));
 const list=MEMBER_ROWS.filter(m=>!joined.has(m.memberKey)&&(!q||[m.memberId,m.name,m.nickname,m.group].some(v=>memberSearchText(v).includes(q)))).slice(0,30);
 const sel=document.querySelector('#workspaceMemberSelect');sel.innerHTML=list.length?list.map(m=>'<option value="'+esc(m.memberKey)+'">'+esc(m.name)+' ('+esc(m.memberId)+')'+(m.nickname?' · '+esc(m.nickname):'')+'</option>').join(''):'<option value="">找不到會員</option>';
 if(list.length===1)sel.value=list[0].memberKey;
}
function renderWorkspace(){
 const e=WORKSPACE_EVENT;if(!e)return;const body=document.querySelector('#workspacePlayerRows');document.querySelector('#workspacePlayerCount').textContent=WORKSPACE_PLAYERS.length+' 人';
 let totals={buyin:0,rebuy:0,groups:0,gross:0,early:0,late:0,rd:0,overbuy:0,other:0,paid:0};
 body.innerHTML=WORKSPACE_PLAYERS.map(p=>{const c=eventCalcPlayer(p,e);Object.keys(totals).forEach(k=>totals[k]+=Number(c[k]||0));
 return '<tr data-key="'+esc(p.memberKey)+'"><td class="name">'+esc(p.name||'')+'</td><td>'+esc(p.memberId||'')+'</td>'+
 '<td><input class="ws-buyin" type="number" min="0" value="'+c.buyin+'"></td><td><input class="ws-rebuy" type="number" min="0" value="'+c.rebuy+'"></td><td class="ws-groups">'+c.groups+'</td>'+
 '<td><input class="ws-early" type="number" min="0" value="'+c.early+'"></td><td><input class="ws-late" type="number" min="0" value="'+c.late+'"></td><td class="ws-auto">'+money(c.rd)+'</td>'+
 '<td class="ws-auto">'+money(c.overbuy)+'</td><td><input class="ws-other" type="number" min="0" value="'+c.other+'"></td><td class="ws-paid">'+money(c.paid)+'</td>'+
 '<td><input class="ws-chips" type="number" min="0" value="'+Number(p.chips||0)+'"></td><td><button class="danger ws-remove">移除</button></td><td><input class="ws-group" value="'+esc(p.group||'')+'"></td></tr>'}).join('');
 document.querySelector('#workspacePlayerTotals').innerHTML='<tr><td colspan="2">合計</td><td>'+totals.buyin+'</td><td>'+totals.rebuy+'</td><td>'+totals.groups+'</td><td>'+money(totals.early)+'</td><td>'+money(totals.late)+'</td><td>'+money(totals.rd)+'</td><td>'+money(totals.groupDiscount)+'</td><td>'+money(totals.other)+'</td><td>'+money(totals.paid)+'</td><td colspan="3"></td></tr>';
 const adminGross=totals.buyin*Number(e.buyinAdmin||0)+totals.rebuy*Number(e.rebuyAdmin||0),discounts=totals.early+totals.late+totals.rd+totals.overbuy+totals.other;
 const prizeBase=totals.buyin*Math.max(0,Number(e.buyinTotal||0)-Number(e.buyinAdmin||0))+totals.rebuy*Math.max(0,Number(e.rebuyTotal||0)-Number(e.rebuyAdmin||0));
 const jp=Math.floor(adminGross*Number(e.jpRate||0)/100),unit=Math.max(1,Number(e.icmRound||100)),prize=Math.floor((prizeBase*(1-Number(e.icmRate||0)/100))/unit)*unit;
 const k=[['參賽人數',WORKSPACE_PLAYERS.length],['重買人數',WORKSPACE_PLAYERS.filter(p=>Number(p.rebuy||0)>0).length],['總組數',totals.groups],['總買入',money(totals.gross)],['總優惠',money(discounts)],['總獎金',money(prize)],['實收行政費',money(Math.max(0,adminGross-discounts))],['JP',money(jp)]];
 document.querySelector('#workspaceKpis').innerHTML=k.map(x=>'<div class="card event-kpi"><small>'+x[0]+'</small><b>'+x[1]+'</b></div>').join('');
}
async function addWorkspacePlayer(){
 const key=document.querySelector('#workspaceMemberSelect').value;if(!key)return;
 const btn=document.querySelector('#workspaceAddPlayer'),m=MEMBER_ROWS.find(x=>x.memberKey===key),group=document.querySelector('#workspaceGroup').value.trim()||m?.group||'';
 if(!m||WORKSPACE_PLAYERS.some(x=>x.memberKey===key))return;
 const optimistic={memberKey:key,memberId:m.memberId,name:m.name,buyin:1,rebuy:0,group,chips:0};
 WORKSPACE_PLAYERS.push(optimistic);renderWorkspace();filterWorkspaceMembers('');try{localStorage.setItem('eightEventPlayers:'+ACTIVE_EVENT,JSON.stringify(WORKSPACE_PLAYERS))}catch(_){}
 document.querySelector('#workspaceMemberSearch').value='';document.querySelector('#workspaceGroup').value='';btn.disabled=true;
 try{
   await api('eight.eventPlayers.add',{eventId:ACTIVE_EVENT,memberKey:key});
   if(group)api('eight.eventPlayers.update',{eventId:ACTIVE_EVENT,memberKey:key,patch:{group}}).catch(console.error);
   btn.disabled=false;
   loadWorkspacePlayers();loadEvents();
 }catch(err){
   btn.disabled=false;
   const eventId=ACTIVE_EVENT;
   api('eight.eventPlayers.list',{eventId}).then(r=>{
     if(ACTIVE_EVENT!==eventId)return;
     const rows=r.players||[],committed=rows.some(x=>x.memberKey===key);
     WORKSPACE_PLAYERS=rows;renderWorkspace();filterWorkspaceMembers('');
     try{localStorage.setItem('eightEventPlayers:'+eventId,JSON.stringify(rows))}catch(_){}
     if(!committed)alert('加入失敗：'+err.message)
   }).catch(()=>alert('加入狀態尚未確認，請稍後按重新整理'));
 }
}
document.querySelector('#eventList').addEventListener('click',e=>{
 const del=e.target.closest('.delete-event');if(del)return;
 const enter=e.target.closest('.enter-event'),row=e.target.closest('.event-row[data-event-id]'),id=enter?.dataset.id||row?.dataset.eventId;
 if(id)openEventWorkspace(id)
});
function returnToEventList(){
 const eventId=ACTIVE_EVENT,keys=[...WS_PENDING_PATCH.keys()],events=document.querySelector('#events'),workspace=document.querySelector('#eventWorkspace');
 document.querySelectorAll('.page').forEach(x=>x.classList.remove('active'));if(events)events.classList.add('active');if(workspace)workspace.classList.remove('active');
 document.querySelectorAll('#nav button').forEach(x=>x.classList.toggle('active',x.dataset.page==='events'));const t=document.querySelector('#pageTitle');if(t)t.textContent=pages.events;
 ACTIVE_EVENT=null;WORKSPACE_EVENT=null;
 if(eventId&&keys.length)Promise.allSettled(keys.map(key=>flushWorkspacePlayerForEvent(eventId,key))).catch(console.error)
}
document.querySelector('#backToEvents').onclick=returnToEventList;
document.querySelector('#workspaceRefresh').addEventListener('click',()=>loadWorkspacePlayers(true));
document.querySelector('#workspaceMemberSearch').addEventListener('input',e=>filterWorkspaceMembers(e.target.value));
document.querySelector('#workspaceMemberSearch').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();addWorkspacePlayer()}});
document.querySelector('#workspaceAddPlayer').addEventListener('click',addWorkspacePlayer);
async function flushWorkspacePlayerForEvent(eventId,key){
 const patch=WS_PENDING_PATCH.get(key);if(!patch||!eventId)return;WS_PENDING_PATCH.delete(key);WS_SAVE_TIMERS.delete(key);
 try{const p=WORKSPACE_PLAYERS.find(x=>x.memberKey===key);if(p&&p.revision)patch.expectedRevision=p.revision;const r=await api('eight.eventPlayers.update',{eventId:eventId,memberKey:key,patch:patch});if(p&&r.player&&r.player.revision)p.revision=r.player.revision}
 catch(err){if(ACTIVE_EVENT===eventId){if(err.message==='STALE_WRITE')alert('資料已被其他裝置更新，已重新載入最新資料');else alert('更新失敗：'+err.message);await loadWorkspacePlayers()}else console.error('背景儲存失敗',err)}
}
async function flushWorkspacePlayer(key){return flushWorkspacePlayerForEvent(ACTIVE_EVENT,key)}
function queueWorkspacePlayerSave(key,patch){
 WS_PENDING_PATCH.set(key,Object.assign({},WS_PENDING_PATCH.get(key)||{},patch));
 if(WS_SAVE_TIMERS.has(key))clearTimeout(WS_SAVE_TIMERS.get(key));
 WS_SAVE_TIMERS.set(key,setTimeout(()=>flushWorkspacePlayer(key),250));
}
document.querySelector('#workspacePlayerRows').addEventListener('change',e=>{
 const tr=e.target.closest('tr[data-key]');if(!tr)return;const key=tr.dataset.key,p=WORKSPACE_PLAYERS.find(x=>x.memberKey===key);if(!p)return;let patch={};
 if(e.target.classList.contains('ws-buyin')){p.buyin=Math.max(0,Number(e.target.value||0));patch.buyin=p.buyin;patch.entries=p.buyin+Number(p.rebuy||0)}
 else if(e.target.classList.contains('ws-rebuy')){p.rebuy=Math.max(0,Number(e.target.value||0));patch.rebuy=p.rebuy;patch.entries=Number(p.buyin??1)+p.rebuy}
 else if(e.target.classList.contains('ws-early')){p.earlyDiscount=Number(e.target.value||0);patch.earlyDiscount=p.earlyDiscount}
 else if(e.target.classList.contains('ws-late')){p.lateDiscount=Number(e.target.value||0);patch.lateDiscount=p.lateDiscount}
 else if(e.target.classList.contains('ws-other')){p.otherDiscount=Number(e.target.value||0);patch.otherDiscount=p.otherDiscount}
 else if(e.target.classList.contains('ws-group')){p.group=e.target.value.trim();patch.group=p.group}
 else if(e.target.classList.contains('ws-chips')){p.chips=Number(e.target.value||0);patch.chips=p.chips}else return;
 renderWorkspace();try{localStorage.setItem('eightEventPlayers:'+ACTIVE_EVENT,JSON.stringify(WORKSPACE_PLAYERS))}catch(_){}queueWorkspacePlayerSave(key,patch)
});
document.querySelector('#workspacePlayerRows').addEventListener('click',async e=>{const b=e.target.closest('.ws-remove');if(!b)return;const tr=b.closest('tr[data-key]');if(!confirm('確定移除此玩家？'))return;const key=tr.dataset.key,old=[...WORKSPACE_PLAYERS];if(WS_SAVE_TIMERS.has(key))clearTimeout(WS_SAVE_TIMERS.get(key));WS_SAVE_TIMERS.delete(key);WS_PENDING_PATCH.delete(key);WORKSPACE_PLAYERS=WORKSPACE_PLAYERS.filter(x=>x.memberKey!==key);renderWorkspace();filterWorkspaceMembers(document.querySelector('#workspaceMemberSearch').value);api('eight.eventPlayers.delete',{eventId:ACTIVE_EVENT,memberKey:key}).then(()=>{loadEvents()}).catch(err=>{WORKSPACE_PLAYERS=old;renderWorkspace();alert('移除失敗：'+err.message)})});
document.querySelector('#workspaceSettle').addEventListener('click',async()=>{
 const btn=document.querySelector('#workspaceSettle');btn.disabled=true;
 try{
   for(const key of [...WS_PENDING_PATCH.keys()])await flushWorkspacePlayer(key);
   const r=await api('eight.events.snapshot',{eventId:ACTIVE_EVENT});
   window.EIGHT_SETTLEMENT_SNAPSHOT=r.snapshot;
   alert('賽事資料已確認同步，可以進入 ICM / 結算。');
 }catch(err){alert('結算前同步失敗：'+err.message)}
 finally{btn.disabled=false}
});

document.querySelector('#eventList').addEventListener('click',e=>{const b=e.target.closest('.delete-event');if(!b)return;if(!confirm('確定刪除此賽事？'))return;const id=b.dataset.id,old=[...(window.EIGHT_EVENTS||[])];window.EIGHT_EVENTS=old.filter(x=>x.eventId!==id);renderEvents(window.EIGHT_EVENTS);api('eight.events.delete',{eventId:id}).catch(err=>{window.EIGHT_EVENTS=old;renderEvents(old);alert('刪除失敗：'+err.message)})});
loadEvents();

window.addEventListener('offline',()=>{const el=document.querySelector('#dbStatus');if(el){el.textContent='資料庫：網路離線';el.className='bad'}});
window.addEventListener('online',()=>{boot()});
