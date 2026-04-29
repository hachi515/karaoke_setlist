// ============================== Data placeholders ==============================
__COOL_DATA_LINE__
__RANK_DATA_LINE__
__TREND_DATA_LINE__
__HISTORY_DATA_LINE__
const ROOMS = __ROOMS_JSON__;
const CATS = __CATS_JSON__;
const TREND_PERIODS = __TREND_PERIODS_JSON__;
const TREND_TARGET_CAT = __TREND_TARGET_JSON__;
const UPDATE_TS = "__UPDATE_TS__";
const MYLIST_GAS_URL = "https://script.google.com/macros/s/AKfycbwhx-_zhE7IJvRLgOegHpQ5WYqC6cBG3fENWtMa0sEX0Rwnh__dPoS25OPkkU9xvIDp/exec";

document.getElementById('updateLine').innerText = UPDATE_TS + ' 更新';

// ============================== Utilities ==============================
function escHtml(s){return String(s||"").replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function escAttr(s){return String(s||"").replace(/'/g,"&#39;").replace(/"/g,'&quot;');}
function jsNormalize(s){if(!s)return"";s=String(s).normalize('NFKC');s=s.replace(/\.[a-zA-Z0-9]{3,4}$/,'');s=s.replace(/[\[\(\{【].*?[\]\)\}】]/g,' ');s=s.replace(/(key|KEY)?\s*[\+\-]\s*[0-9]+/g,' ');s=s.replace(/原キー/g,' ');s=s.replace(/(キー)?変更[:：]?/g,' ');s=s.replace(/[~〜～\-_=,.]/g,' ');s=s.replace(/\s+/g,' ').trim();return s.toUpperCase();}
function cleanSearchKeyword(s){
  if(!s) return '';
  s = String(s);
  s = s.replace(/\([^)]*\)/g,' ').replace(/（[^）]*）/g,' ');
  s = s.replace(/[～〜~／\/]/g,' ');
  s = s.replace(/\s+/g,' ').trim();
  return s;
}

// ============================== Dark mode ==============================
(function(){
  const tg = document.getElementById('themeToggle');
  if(!tg) return;
  function apply(d){
    document.body.classList.toggle('dark', d);
    const ic = tg.querySelector('i');
    if(ic) ic.className = d ? 'fas fa-sun' : 'fas fa-moon';
  }
  let dark = false;
  try { dark = localStorage.getItem('darkMode') === '1'; } catch(e){}
  apply(dark);
  tg.addEventListener('click', ()=>{
    dark = !dark;
    try { localStorage.setItem('darkMode', dark ? '1' : '0'); } catch(e){}
    apply(dark);
  });
})();

// ============================== Tabs ==============================
document.querySelectorAll('.tab-btn').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    const tab = btn.getAttribute('data-tab');
    document.querySelectorAll('.tab-btn').forEach(b=>b.classList.toggle('active', b===btn));
    document.querySelectorAll('.tab-content').forEach(c=>c.classList.toggle('active', c.id==='tab-'+tab));
  });
});

// ============================== Singer (top) ==============================
const viewerSingerInput = document.getElementById('viewerSinger');
try {
  const saved = localStorage.getItem('viewerSinger');
  if(saved) viewerSingerInput.value = saved;
} catch(e){}
viewerSingerInput.addEventListener('input', ()=>{
  try { localStorage.setItem('viewerSinger', viewerSingerInput.value); } catch(e){}
});

// ============================== Room (port) selection ==============================
let viewerPort = 11059;
let viewerSearchType = 'ykr'; // 'ykr' or 'eve'
try {
  const p = parseInt(localStorage.getItem('viewerPort'));
  if(p && p > 0) viewerPort = p;
  const t = localStorage.getItem('viewerSearchType');
  if(t==='ykr' || t==='eve') viewerSearchType = t;
} catch(e){}

function effectiveSearchType(){
  if(viewerPort === 11058 || viewerPort === 11060) return 'eve';
  return viewerSearchType;
}

function buildSearchUrl(anime, song){
  const kw = cleanSearchKeyword((anime ? anime + ' ' : '') + song);
  const path = effectiveSearchType()==='ykr'
    ? 'search_listerdb_filelist.php?anyword='
    : 'search.php?searchword=';
  return 'http://ykr.moe:' + viewerPort + '/' + path + encodeURIComponent(kw);
}

function refreshRoomBtnLabel(){
  const name = ROOMS[viewerPort] || 'カスタム部屋';
  document.getElementById('roomPickLabel').innerText = name;
  document.getElementById('roomPickPort').innerText = viewerPort;
}
refreshRoomBtnLabel();

const roomModal = document.getElementById('roomModalOverlay');
function openRoomModal(){
  // populate radio
  const rs = document.querySelectorAll('input[name="rmSearchType"]');
  rs.forEach(r=>{ r.checked = (r.value === viewerSearchType); });
  // populate grid
  const grid = document.getElementById('roomGridBtns');
  let html = '';
  Object.entries(ROOMS).forEach(([port, name])=>{
    const sel = (parseInt(port)===viewerPort) ? ' selected' : '';
    html += '<button type="button" class="room-btn'+sel+'" data-port="'+port+'">'+escHtml(name)+'<span class="port">'+port+'</span></button>';
  });
  grid.innerHTML = html;
  grid.querySelectorAll('.room-btn').forEach(b=>{
    b.addEventListener('click', ()=>{
      const p = parseInt(b.getAttribute('data-port'));
      selectPort(p);
    });
  });
  document.getElementById('roomPortInput').value = viewerPort;
  roomModal.classList.add('active');
}
function closeRoomModal(){ roomModal.classList.remove('active'); }
document.getElementById('roomPickBtn').addEventListener('click', openRoomModal);
document.getElementById('roomModalClose').addEventListener('click', closeRoomModal);
roomModal.addEventListener('click', e=>{ if(e.target===roomModal) closeRoomModal(); });

document.querySelectorAll('input[name="rmSearchType"]').forEach(r=>{
  r.addEventListener('change', ()=>{
    viewerSearchType = r.value;
    try { localStorage.setItem('viewerSearchType', viewerSearchType); } catch(e){}
  });
});

function selectPort(p){
  if(!p || p<=0) return;
  viewerPort = p;
  try { localStorage.setItem('viewerPort', String(viewerPort)); } catch(e){}
  refreshRoomBtnLabel();
  closeRoomModal();
}

document.getElementById('roomPortSet').addEventListener('click', ()=>{
  const v = parseInt(document.getElementById('roomPortInput').value);
  if(v>0) selectPort(v);
  else showDialog('有効なポート番号を入力してください', 'alert');
});
document.getElementById('roomPortInput').addEventListener('keydown', e=>{
  if(e.key==='Enter'){ e.preventDefault(); document.getElementById('roomPortSet').click(); }
});

// ============================== Cool / Ranking / Trend rendering ==============================
// Singer modal (existing)
function findHistoryMatches(w, s){
  const sn = jsNormalize(s);
  const wn = jsNormalize(w);
  if(!sn && !wn) return [];
  return HISTORY.filter(h=>{
    let so=false, wo=false;
    if(sn){
      if(/^[A-Z0-9 ]+$/.test(sn)){
        const re = new RegExp('(?:^|[^A-Z0-9])' + sn.replace(/[.*+?^${}()|[\]\\]/g,'\\$&') + '(?:[^A-Z0-9]|$)','i');
        so = re.test(h.sn);
      } else {
        so = h.sn.indexOf(sn) >= 0;
      }
    }
    if(wn){ wo = (h.sn.indexOf(wn) >= 0) || (h.wn.indexOf(wn) >= 0); }
    if(sn && wn) return so && wo;
    if(sn) return so;
    if(wn) return wo;
    return false;
  });
}
function openSingersModal(w, s){
  const m = findHistoryMatches(w, s);
  const u = {};
  m.forEach(x=>{ if(!u[x.u]) u[x.u]={c:0, d:x.d, r:x.rm}; u[x.u].c++; if(x.d>u[x.u].d){u[x.u].d=x.d; u[x.u].r=x.rm;} });
  const us = Object.entries(u).sort((a,b)=>b[1].c-a[1].c);
  document.getElementById('modalTitle').innerHTML = escHtml(s) + '<small>' + escHtml(w) + ' - ' + m.length + '件 / ' + us.length + '人</small>';
  let b = '<div class="modal-summary"><i class="fas fa-users"></i> ' + us.length + '人がこの曲を歌っています（合計' + m.length + '回）</div>';
  if(us.length===0) b += '<div class="empty">履歴なし</div>';
  else us.forEach(([uu,info])=>{
    b += '<div class="modal-row"><div class="top"><div class="user"><i class="fas fa-microphone"></i> ' + escHtml(uu) + ' <span style="color:var(--accent);font-size:11.5px;margin-left:4px">×'+info.c+'</span></div><div class="date">' + escHtml(info.d) + '</div></div><div class="meta">最新: ' + escHtml(info.r) + '</div></div>';
  });
  document.getElementById('modalBody').innerHTML = b;
  document.getElementById('modalOverlay').classList.add('active');
}
function closeModal(){ document.getElementById('modalOverlay').classList.remove('active'); }
document.getElementById('modalOverlay').addEventListener('click', e=>{ if(e.target.id==='modalOverlay') closeModal(); });
function toggleCard(h){ h.parentElement.classList.toggle('expanded'); }

// Type chips/pills
function typeChipHtml(t){
  if(t==='OP') return '<div class="type-chip tc-op">OP</div>';
  if(t==='ED') return '<div class="type-chip tc-ed">ED</div>';
  if(t==='IN') return '<div class="type-chip tc-in">IN</div>';
  return '<div class="type-chip tc-other">'+escHtml(t||'')+'</div>';
}
function typePillHtml(t){
  if(t==='OP') return '<span class="type-pill tp-op">OP</span>';
  if(t==='ED') return '<span class="type-pill tp-ed">ED</span>';
  if(t==='IN') return '<span class="type-pill tp-in">IN</span>';
  return '<span class="type-pill tp-other">'+escHtml(t||'')+'</span>';
}
function flatMetricHtml(label, val, kind){
  return '<div class="flat-metric flat-'+kind+'"><div class="lbl">'+label+'</div><div class="val">'+val+'</div></div>';
}

// Search button + Mylist register button (for cool/ranking/trend rows)
function rowActionsHtml(anime, song, artist){
  const data = 'data-anime="'+escAttr(anime)+'" data-song="'+escAttr(song)+'" data-artist="'+escAttr(artist||'')+'"';
  const search = '<button type="button" class="viewer-search-btn js-search" '+data+' title="検索"><i class="fas fa-search"></i></button>';
  const mylist = '<button type="button" class="viewer-mylist-btn js-mylist-add" '+data+' title="マイリストに登録"><i class="fas fa-bookmark"></i> マイリスト</button>';
  return {search, mylist};
}

function buildCoolCard(w, rank, cat){
  const opVal = (w.op_n||0)>0 ? (w.op_created||0) : '-';
  const edVal = (w.ed_n||0)>0 ? (w.ed_created||0) : '-';
  const inVal = (w.in_n||0)>0 ? (w.in_created||0) : '-';
  const opTag = '<span class="type-pill tp-op">OP <b>'+opVal+'</b></span>';
  const edTag = '<span class="type-pill tp-ed">ED <b>'+edVal+'</b></span>';
  const inTag = '<span class="type-pill tp-in">IN <b>'+inVal+'</b></span>';
  const createdTag = (w.total_created||0) > 0
    ? '<span class="type-pill tp-created" title="作成数の合計"><b>'+w.total_created+'</b>／'+w.songs.length+' 作成</span>'
    : '';
  let songsHtml = '';
  w.songs.forEach(s=>{
    const acts = rowActionsHtml(w.anime, s.song, s.artist);
    const createdMark = (s.creation_count||0) > 0
      ? '<span class="song-created-mark" title="offline_list 内の作成数">作成'+s.creation_count+'</span>'
      : '';
    songsHtml += '<div class="song-row">'
      + typeChipHtml(s.type)
      + '<div class="song-info-wrap">'
        + '<div class="song-name">'+escHtml(s.song)+createdMark+'</div>'
        + '<div class="song-artist">'+escHtml(s.artist)+'</div>'
      + '</div>'
      + '<div class="viewer-row-actions">'+acts.search+'</div>'
      + '<div class="viewer-metrics-stack">'
        + acts.mylist
        + '<div class="song-metrics">'+flatMetricHtml('人数',s.user_count,'user')+flatMetricHtml('歌唱数',s.count,'song')+'</div>'
      + '</div>'
      + '</div>';
  });
  return '<div class="card">'
    + '<div class="cool-head" onclick="toggleCard(this)">'
    + '<div class="num-badge'+(rank===1?' gold':rank===2?' silver':rank===3?' bronze':'')+'">'+rank+'</div>'
    + '<div class="cool-info"><div class="cool-anime">'+escHtml(w.anime)+'</div><div class="cool-types">'+opTag+edTag+inTag+createdTag+'</div></div>'
    + '<div class="cool-metrics">'+flatMetricHtml('人数',w.total_user,'user')+flatMetricHtml('歌唱数',w.total_count,'song')+'</div>'
    + '<i class="fas fa-chevron-down card-chev"></i>'
    + '</div>'
    + '<div class="card-detail">'+songsHtml+'</div>'
    + '</div>';
}

const coolCatSel = document.getElementById('coolCat');
const coolSortSel = document.getElementById('coolSort');
CATS.forEach(c=>{ const o=document.createElement('option'); o.value=c; o.innerText=c; coolCatSel.appendChild(o); });
coolCatSel.addEventListener('change', renderCool);
coolSortSel.addEventListener('change', renderCool);
function renderCool(){
  const cat = coolCatSel.value;
  const sort = coolSortSel.value;
  const data = COOL_DATA[cat] || {works:[]};
  const works = data.works.slice();
  if(sort==='count') works.sort((a,b)=>b.total_count-a.total_count || b.total_user-a.total_user);
  else if(sort==='user') works.sort((a,b)=>b.total_user-a.total_user || b.total_count-a.total_count);
  else if(sort==='name') works.sort((a,b)=>a.anime.localeCompare(b.anime,'ja'));
  else if(sort==='created') works.sort((a,b)=>{
    const ac = a.songs.reduce((s,x)=>s+(x.creation_count||0),0);
    const bc = b.songs.reduce((s,x)=>s+(x.creation_count||0),0);
    return bc-ac;
  });
  document.getElementById('coolCount').innerText = works.length;
  const list = document.getElementById('coolList');
  if(works.length===0){ list.innerHTML='<div class="empty">データがありません</div>'; return; }
  let html = '';
  works.forEach((w,i)=> html += buildCoolCard(w, i+1, cat));
  list.innerHTML = html;
}

// Ranking
const rankCatSel = document.getElementById('rankCat');
const rankModeSel = document.getElementById('rankMode');
CATS.forEach(c=>{ const o=document.createElement('option'); o.value=c; o.innerText=c; rankCatSel.appendChild(o); });
rankCatSel.addEventListener('change', renderRanking);
rankModeSel.addEventListener('change', renderRanking);
function buildRankingTop20(cat, mode){
  const all = (RANK_DATA[cat]||[]).slice();
  if(mode==='count') all.sort((a,b)=>b.count-a.count || b.user_count-a.user_count);
  else all.sort((a,b)=>b.user_count-a.user_count || b.count-a.count);
  const top20=[]; let pv=null, cr=0;
  for(let i=0;i<all.length;i++){
    const v = mode==='count' ? all[i].count : all[i].user_count;
    if(v!==pv){ cr=i+1; pv=v; }
    if(cr>20) break;
    top20.push(Object.assign({}, all[i], {rank:cr}));
  }
  return top20;
}
function buildRankCardHtml(r, mode){
  const grade = r.rank===1?'gold':r.rank===2?'silver':r.rank===3?'bronze':'';
  const acts = rowActionsHtml(r.anime, r.song, r.artist);
  return '<div class="rank-row-flat '+grade+'">'
    + '<div class="num-badge '+grade+'">'+r.rank+'</div>'
    + '<div class="rank-info"><div class="rank-anime">'+escHtml(r.anime)+'</div><div class="rank-sub">'+escHtml(r.song)+' / '+escHtml(r.artist)+'</div><div class="rank-types-inline">'+typePillHtml(r.type)+'</div></div>'
    + '<div class="viewer-row-actions">'+acts.search+'</div>'
    + '<div class="viewer-metrics-stack">'+acts.mylist
    + '<div class="song-metrics">'+flatMetricHtml('人数',r.user_count,'user')+flatMetricHtml('歌唱数',r.count,'song')+'</div></div>'
    + '</div>';
}
function renderRanking(){
  const cat = rankCatSel.value;
  const mode = rankModeSel.value;
  const top20 = buildRankingTop20(cat, mode);
  document.getElementById('rankCount').innerText = top20.length;
  const list = document.getElementById('rankList');
  if(top20.length===0){ list.innerHTML='<div class="empty">ランキング対象データがありません</div>'; return; }
  let html = '';
  top20.forEach(r=> html += buildRankCardHtml(r, mode));
  list.innerHTML = html;
}

// Trend
const trendPeriodSel = document.getElementById('trendPeriod');
const trendSortSel = document.getElementById('trendSort');
const trendModeSel = document.getElementById('trendMode');
TREND_PERIODS.forEach(p=>{
  const o = document.createElement('option');
  o.value = String(p); o.innerText = '直近'+p+'日';
  if(p===7) o.selected = true;
  trendPeriodSel.appendChild(o);
});
trendPeriodSel.addEventListener('change', renderTrend);
trendSortSel.addEventListener('change', renderTrend);
if(trendModeSel) trendModeSel.addEventListener('change', renderTrend);
function trendCurValue(it, mode){ return mode==='user' ? (it.cur_user||0) : (it.cur_count||0); }
function trendPrevValue(it, mode){ if(mode==='user') return (it.prev_user!=null ? it.prev_user : 0); return (it.prev_count!=null ? it.prev_count : 0); }
function trendDelta(it, mode){ return trendCurValue(it, mode) - trendPrevValue(it, mode); }
function growthBadgeHtml(it, mode){
  const cur = trendCurValue(it, mode);
  const prev = trendPrevValue(it, mode);
  const d = cur - prev;
  const unit = mode==='user' ? '人' : '曲';
  if(prev<=0 && cur>0) return '<span class="growth-badge new">NEW +'+cur+unit+'</span>';
  if(prev<=0 && cur<=0) return '';
  if(d===0) return '<span class="growth-badge flat">±0'+unit+'</span>';
  if(d>0) return '<span class="growth-badge up">+'+d+unit+'</span>';
  return '<span class="growth-badge down">'+d+unit+'</span>';
}
function buildTrendItems(){
  const sort = trendSortSel.value;
  const mode = trendModeSel ? trendModeSel.value : 'count';
  const period = trendPeriodSel.value;
  const td = TREND_DATA[period] || {kpi:{surge_count:0,new_in:0,max_delta:0}, items:[]};
  const items = td.items.map(it=>{
    const d = trendDelta(it, mode);
    return Object.assign({}, it, {_delta_m:d, _cur_m:trendCurValue(it,mode), _is_new_m:(trendPrevValue(it,mode)<=0 && trendCurValue(it,mode)>0)});
  });
  if(sort==='surge') items.sort((a,b)=>b._delta_m-a._delta_m || b._cur_m-a._cur_m);
  else if(sort==='new') items.sort((a,b)=>(b._is_new_m?1:0)-(a._is_new_m?1:0) || b._delta_m-a._delta_m);
  else items.sort((a,b)=>b._delta_m-a._delta_m);
  const surge_count = items.filter(x=>x._delta_m>0).length;
  const new_in = items.filter(x=>x._is_new_m).length;
  const max_delta = items.reduce((m,x)=>Math.max(m, x._delta_m), 0);
  return {kpi:{surge_count,new_in,max_delta}, items, mode};
}
function buildTrendHtml(){
  const r = buildTrendItems();
  const kpi = r.kpi, items = r.items, mode = r.mode;
  const unit = mode==='user' ? '人' : '曲';
  let html = '';
  html += '<div class="trend-stats">'
    + '<div class="trend-stat"><div class="ico up"><i class="fas fa-arrow-trend-up"></i></div><div class="lbl">急上昇</div><div class="val">'+kpi.surge_count+'<small>'+unit+'</small></div></div>'
    + '<div class="trend-stat"><div class="ico new"><i class="far fa-star"></i></div><div class="lbl">新規ランクイン</div><div class="val">'+kpi.new_in+'<small>'+unit+'</small></div></div>'
    + '<div class="trend-stat"><div class="ico max"><i class="fas fa-arrow-up"></i></div><div class="lbl">最大伸び</div><div class="val">+'+kpi.max_delta+'</div></div>'
    + '</div>';
  if(items.length>0){
    const p = items[0];
    const acts = rowActionsHtml(p.anime, p.song, p.artist);
    html += '<div class="trend-pickup">'
      + '<div class="trend-pickup-head"><i class="fas fa-fire"></i> 急上昇ピックアップ</div>'
      + '<div class="trend-pickup-row">'
        + '<div class="num-badge">1</div>'
        + '<div class="pickup-info"><div class="pickup-anime">'+escHtml(p.anime)+'</div><div class="pickup-sub">'+escHtml(p.song)+' / '+escHtml(p.artist)+'</div><div class="pickup-types-inline">'+typePillHtml(p.type)+'</div></div>'
        + '<div class="viewer-row-actions">'+acts.search+'</div>'
        + '<div class="viewer-metrics-stack">'+acts.mylist
          + '<div class="song-metrics">'+flatMetricHtml('人数',p.cur_user||0,'user')+flatMetricHtml('歌唱数',p.cur_count||0,'song')+'</div>'
          + growthBadgeHtml(p, mode)
        + '</div>'
      + '</div></div>';
  }
  if(items.length>1){
    html += '<div class="notable-head"><i class="fas fa-arrow-trend-up"></i> 注目の上昇曲</div>';
    items.slice(1, 10).forEach((it, idx)=>{
      const rank = idx+2;
      const acts = rowActionsHtml(it.anime, it.song, it.artist);
      html += '<div class="notable-row">'
        + '<div class="num-badge">'+rank+'</div>'
        + '<div class="notable-info"><div class="notable-anime">'+escHtml(it.anime)+'</div><div class="notable-artist">'+escHtml(it.song)+' / '+escHtml(it.artist)+'</div><div class="notable-types">'+typePillHtml(it.type)+'</div></div>'
        + '<div class="viewer-row-actions">'+acts.search+'</div>'
        + '<div class="viewer-metrics-stack">'+acts.mylist
          + '<div class="song-metrics">'+flatMetricHtml('人数',it.cur_user||0,'user')+flatMetricHtml('歌唱数',it.cur_count||0,'song')+'</div>'
          + growthBadgeHtml(it, mode)
        + '</div>'
      + '</div>';
    });
  }
  if(items.length===0) html += '<div class="empty">急上昇データがありません</div>';
  return html;
}
function renderTrend(){ document.getElementById('trendBody').innerHTML = buildTrendHtml(); }

// Delegate clicks for search/mylist buttons (and modal-row click)
document.body.addEventListener('click', function(e){
  const sb = e.target.closest('.js-search');
  if(sb){
    e.stopPropagation();
    const a = sb.getAttribute('data-anime') || '';
    const s = sb.getAttribute('data-song') || '';
    // Same tab, NOT new tab
    window.location.href = buildSearchUrl(a, s);
    return;
  }
  const mb = e.target.closest('.js-mylist-add');
  if(mb){
    e.stopPropagation();
    const a = mb.getAttribute('data-anime') || '';
    const s = mb.getAttribute('data-song') || '';
    const ar = mb.getAttribute('data-artist') || '';
    openMylistRegister(a, s, ar);
    return;
  }
  // For cool song-row body click -> show singer modal (existing behavior)
  const sr = e.target.closest('.song-row');
  if(sr){
    if(e.target.closest('.viewer-search-btn, .viewer-mylist-btn, button, a')) return;
    const nameEl = sr.querySelector('.song-name');
    if(!nameEl) return;
    // Find anime by walking up to .card
    const card = sr.closest('.card');
    if(!card) return;
    const anime = (card.querySelector('.cool-anime')||{}).innerText || '';
    // song name is innerText minus possible "作成N" mark
    let songText = nameEl.cloneNode(true);
    const mk = songText.querySelector('.song-created-mark');
    if(mk) mk.remove();
    openSingersModal(anime, songText.innerText);
  }
});

// ============================== Mylist register modal ==============================
let mlRegContext = null;
function openMylistRegister(anime, song, artist){
  mlRegContext = {anime: anime, song: song, artist: artist};
  document.getElementById('mlRegSinger').value = (viewerSingerInput.value || '').trim();
  document.getElementById('mlRegWork').innerText = anime;
  document.getElementById('mlRegArtist').innerText = artist || '-';
  document.getElementById('mlRegSong').innerText = song;
  document.getElementById('mlRegPart').checked = false;
  document.getElementById('mlRegPrac').checked = false;
  document.getElementById('mlRegRev').checked = false;
  document.getElementById('mlRegOverlay').classList.add('active');
}
document.getElementById('mlRegClose').addEventListener('click', ()=>document.getElementById('mlRegOverlay').classList.remove('active'));
document.getElementById('mlRegCancel').addEventListener('click', ()=>document.getElementById('mlRegOverlay').classList.remove('active'));
document.getElementById('mlRegOverlay').addEventListener('click', e=>{ if(e.target.id==='mlRegOverlay') document.getElementById('mlRegOverlay').classList.remove('active'); });
document.getElementById('mlRegSubmit').addEventListener('click', ()=>{
  if(!mlRegContext) return;
  const singer = document.getElementById('mlRegSinger').value.trim();
  const item = {
    id: Date.now() + Math.random(),
    singer: singer,
    work: mlRegContext.anime || '',
    artist: mlRegContext.artist || '',
    song: mlRegContext.song || '',
    isPartDivision: document.getElementById('mlRegPart').checked,
    isPracticing: document.getElementById('mlRegPrac').checked,
    isReviewNeeded: document.getElementById('mlRegRev').checked,
    isPrivate: false,
    viewPassword: ''
  };
  if(!item.song){ showDialog('曲名がありません', 'alert'); return; }
  // Persist singer back to top input
  if(singer && singer !== viewerSingerInput.value){
    viewerSingerInput.value = singer;
    try { localStorage.setItem('viewerSinger', singer); } catch(e){}
  }
  mylistAddItem(item);
  document.getElementById('mlRegOverlay').classList.remove('active');
  showDialog('マイリストに登録しました。\n「同期して保存」でサーバーに反映できます。', 'alert');
});

// ============================== Mylist (full feature) ==============================
const ML_ITEMS_PER_PAGE = 10;
let mlSongs = [];
let mlDeletedIds = [];
let mlActiveTab = 'ALL';
let mlSearchQuery = '';
let mlCurrentPage = 1;
let mlEditingId = null;
let mlFilterTags = {isPartDivision:false, isPracticing:false, isReviewNeeded:false};
let mlForceOverwrite = false;

try { mlSongs = JSON.parse(localStorage.getItem('myMusicList')||'[]'); } catch(e){ mlSongs = []; }
if(!Array.isArray(mlSongs)) mlSongs = [];
try { mlDeletedIds = JSON.parse(localStorage.getItem('deletedIds')||'[]'); } catch(e){ mlDeletedIds = []; }

const mlEls = {
  inputSinger: document.getElementById('mlInputSinger'),
  inputWork: document.getElementById('mlInputWork'),
  inputArtist: document.getElementById('mlInputArtist'),
  inputSong: document.getElementById('mlInputSong'),
  inputPart: document.getElementById('mlInputPart'),
  inputPrac: document.getElementById('mlInputPrac'),
  inputRev: document.getElementById('mlInputRev'),
  inputPrivate: document.getElementById('mlInputPrivate'),
  inputPw: document.getElementById('mlInputPw'),
  privatePwWrap: document.getElementById('mlPrivatePwWrap'),
  submit: document.getElementById('mlSubmitBtn'),
  cancel: document.getElementById('mlCancelBtn'),
  formIcon: document.getElementById('mlFormIcon'),
  formTitle: document.getElementById('mlFormTitle'),
  singerSel: document.getElementById('mlSingerSelect'),
  singerCount: document.getElementById('mlSingerCount'),
  searchInput: document.getElementById('mlSearchInput'),
  searchBtn: document.getElementById('mlSearchBtn'),
  filterPart: document.getElementById('mlFilterPart'),
  filterPrac: document.getElementById('mlFilterPrac'),
  filterRev: document.getElementById('mlFilterRev'),
  viewPw: document.getElementById('mlViewPw'),
  viewPwApply: document.getElementById('mlViewPwApply'),
  list: document.getElementById('mlList'),
  pager: document.getElementById('mlPager'),
  count: document.getElementById('mlListCount'),
  title: document.getElementById('mlListTitle'),
  status: document.getElementById('mlStatus'),
  csvBtn: document.getElementById('mlCsvBtn'),
  importJsonBtn: document.getElementById('mlImportJsonBtn'),
  importJson: document.getElementById('mlImportJson'),
  importCsv: document.getElementById('mlImportCsv'),
  syncBtn: document.getElementById('mlSyncBtn')
};

function mlCleanForSearch(str){
  if(!str) return '';
  let s = str;
  s = s.replace(/\s*[\(（【\[~～].*$/, '');
  s = s.replace(/[!！?？]/g, '');
  return s.trim();
}
function mlNormalizeStr(str){ if(!str) return ''; return str.normalize('NFKC').replace(/[\s\u3000]+/g, ' ').trim(); }

function mylistAddItem(item){
  mlSongs.push(item);
  localStorage.setItem('myMusicList', JSON.stringify(mlSongs));
  if(item.singer) mlActiveTab = item.singer;
  mlRender();
}

function mlRender(){
  try { localStorage.setItem('myMusicList', JSON.stringify(mlSongs)); } catch(e){}
  const uniqueSingers = Array.from(new Set(mlSongs.map(s=>s.singer).filter(Boolean)));
  let optionsHtml = '<option value="ALL">全ての歌唱者を表示 ('+mlSongs.length+')</option>';
  uniqueSingers.forEach(sg=>{ const c = mlSongs.filter(s=>s.singer===sg).length; optionsHtml += '<option value="'+escAttr(sg)+'">'+escHtml(sg)+' ('+c+')</option>'; });
  mlEls.singerSel.innerHTML = optionsHtml;
  mlEls.singerSel.value = mlActiveTab;
  mlEls.singerCount.innerText = '登録数: ' + uniqueSingers.length + '人';

  const vpw = mlEls.viewPw.value.trim();
  let filtered = mlSongs.filter(s=>{
    if(s.isPrivate){ if(!vpw || s.viewPassword !== vpw) return false; }
    return true;
  });
  if(mlActiveTab !== 'ALL') filtered = filtered.filter(s=>s.singer===mlActiveTab);
  if(mlSearchQuery){
    const queries = mlSearchQuery.toLowerCase().replace(/　/g,' ').split(' ').filter(q=>q);
    filtered = filtered.filter(s=>{ const t = (s.song+' '+s.work+' '+s.artist).toLowerCase(); return queries.every(q=>t.includes(q)); });
  }
  if(mlFilterTags.isPartDivision) filtered = filtered.filter(s=>s.isPartDivision);
  if(mlFilterTags.isPracticing) filtered = filtered.filter(s=>s.isPracticing);
  if(mlFilterTags.isReviewNeeded) filtered = filtered.filter(s=>s.isReviewNeeded);

  const display = filtered.slice().reverse();
  const totalPages = Math.ceil(display.length / ML_ITEMS_PER_PAGE) || 1;
  if(mlCurrentPage > totalPages) mlCurrentPage = totalPages;
  if(mlCurrentPage < 1) mlCurrentPage = 1;
  const startIdx = (mlCurrentPage - 1) * ML_ITEMS_PER_PAGE;
  const slice = display.slice(startIdx, startIdx + ML_ITEMS_PER_PAGE);

  mlEls.count.innerText = display.length + ' tracks';
  mlEls.title.innerText = mlActiveTab === 'ALL' ? 'All Songs' : mlActiveTab;

  if(slice.length === 0){
    mlEls.list.innerHTML = '<div class="ml-empty"><i class="far fa-folder-open" style="font-size:32px;opacity:0.3"></i><div style="margin-top:8px">登録された曲はありません。</div></div>';
  } else {
    let html = '';
    slice.forEach(song=>{
      const displayIndex = mlSongs.indexOf(song) + 1;
      const cleanSong = mlCleanForSearch(song.song);
      const cleanWork = mlCleanForSearch(song.work);
      const searchUrl = buildSearchUrl(cleanWork, cleanSong);
      let tags = '';
      if(song.isPartDivision) tags += '<span class="ml-tag pd">パート分け</span>';
      if(song.isPracticing) tags += '<span class="ml-tag pr">練習中</span>';
      if(song.isReviewNeeded) tags += '<span class="ml-tag rv">要復習</span>';
      if(song.isPrivate) tags += '<span class="ml-tag pv">非公開</span>';
      const editing = song.id === mlEditingId;
      html += '<div class="ml-row'+(editing?' editing':'')+'" data-id="'+escAttr(song.id)+'">'
        + '<div class="num">'+displayIndex+'</div>'
        + '<div class="info">'
          + '<div class="song">'+escHtml(song.song)+tags+'</div>'
          + (song.singer ? '<div class="singer"><i class="fas fa-microphone" style="font-size:9px"></i> '+escHtml(song.singer)+'</div>' : '')
          + ((song.artist||song.work) ? '<div class="meta">'+escHtml(song.artist||'')+(song.work?' / '+escHtml(song.work):'')+'</div>' : '')
        + '</div>'
        + '<div class="ml-actions">'
          + '<a class="search" href="'+escAttr(searchUrl)+'" target="_self" rel="noopener" title="検索"><i class="fas fa-search"></i></a>'
          + '<button class="js-ml-edit" data-id="'+escAttr(song.id)+'" title="編集"><i class="fas fa-edit"></i></button>'
          + '<button class="js-ml-up" data-id="'+escAttr(song.id)+'" title="上に"><i class="fas fa-arrow-up"></i></button>'
          + '<button class="js-ml-down" data-id="'+escAttr(song.id)+'" title="下に"><i class="fas fa-arrow-down"></i></button>'
          + '<button class="del js-ml-del" data-id="'+escAttr(song.id)+'" title="削除"><i class="fas fa-trash"></i></button>'
        + '</div>'
      + '</div>';
    });
    mlEls.list.innerHTML = html;
  }

  if(totalPages > 1){
    mlEls.pager.innerHTML = '<button id="mlPagePrev"'+(mlCurrentPage===1?' disabled':'')+'><i class="fas fa-chevron-left"></i></button>'
      + '<span style="font-family:monospace;color:var(--text-sub)">Page '+mlCurrentPage+' / '+totalPages+'</span>'
      + '<button id="mlPageNext"'+(mlCurrentPage===totalPages?' disabled':'')+'><i class="fas fa-chevron-right"></i></button>';
    document.getElementById('mlPagePrev').addEventListener('click', ()=>{ mlCurrentPage--; mlRender(); });
    document.getElementById('mlPageNext').addEventListener('click', ()=>{ mlCurrentPage++; mlRender(); });
    mlEls.pager.style.display = 'flex';
  } else {
    mlEls.pager.innerHTML = '';
    mlEls.pager.style.display = 'none';
  }
}

mlEls.list.addEventListener('click', e=>{
  const ed = e.target.closest('.js-ml-edit');
  if(ed){ mlEdit(parseFloat(ed.getAttribute('data-id'))); return; }
  const up = e.target.closest('.js-ml-up');
  if(up){ mlMove(parseFloat(up.getAttribute('data-id')), 'up'); return; }
  const dn = e.target.closest('.js-ml-down');
  if(dn){ mlMove(parseFloat(dn.getAttribute('data-id')), 'down'); return; }
  const dl = e.target.closest('.js-ml-del');
  if(dl){ mlDelete(parseFloat(dl.getAttribute('data-id'))); return; }
});

mlEls.singerSel.addEventListener('change', ()=>{ mlActiveTab = mlEls.singerSel.value; mlCurrentPage = 1; mlRender(); });
mlEls.searchInput.addEventListener('keydown', e=>{ if(e.key==='Enter') mlTriggerSearch(); });
mlEls.searchBtn.addEventListener('click', mlTriggerSearch);
function mlTriggerSearch(){ mlSearchQuery = mlEls.searchInput.value; mlCurrentPage = 1; mlRender(); }
mlEls.filterPart.addEventListener('change', ()=>{ mlFilterTags.isPartDivision = mlEls.filterPart.checked; mlCurrentPage = 1; mlRender(); });
mlEls.filterPrac.addEventListener('change', ()=>{ mlFilterTags.isPracticing = mlEls.filterPrac.checked; mlCurrentPage = 1; mlRender(); });
mlEls.filterRev.addEventListener('change', ()=>{ mlFilterTags.isReviewNeeded = mlEls.filterRev.checked; mlCurrentPage = 1; mlRender(); });
mlEls.viewPwApply.addEventListener('click', ()=>{ mlCurrentPage = 1; mlRender(); });
mlEls.viewPw.addEventListener('keydown', e=>{ if(e.key==='Enter'){ mlCurrentPage = 1; mlRender(); } });

mlEls.inputPrivate.addEventListener('change', ()=>{ mlEls.privatePwWrap.style.display = mlEls.inputPrivate.checked ? 'block' : 'none'; });

mlEls.submit.addEventListener('click', mlSave);
mlEls.cancel.addEventListener('click', mlCancelEdit);
function mlSave(){
  const singer = mlEls.inputSinger.value.trim();
  const songName = mlEls.inputSong.value.trim();
  if(!songName){ showDialog('曲名は必須です', 'alert'); return; }
  const data = {
    singer: singer,
    work: mlEls.inputWork.value.trim(),
    artist: mlEls.inputArtist.value.trim(),
    song: songName,
    isPartDivision: mlEls.inputPart.checked,
    isPracticing: mlEls.inputPrac.checked,
    isReviewNeeded: mlEls.inputRev.checked,
    isPrivate: mlEls.inputPrivate.checked,
    viewPassword: mlEls.inputPw.value.trim()
  };
  if(mlEditingId !== null){
    const idx = mlSongs.findIndex(s=>s.id === mlEditingId);
    if(idx !== -1){ mlSongs[idx] = Object.assign({}, mlSongs[idx], data); showDialog('曲情報を更新しました。', 'alert'); }
  } else {
    mlSongs.push(Object.assign({id: Date.now() + Math.random()}, data));
  }
  if(singer) mlActiveTab = singer;
  mlResetForm();
  mlRender();
}
function mlEdit(id){
  const s = mlSongs.find(x=>x.id===id);
  if(!s) return;
  mlEls.inputSinger.value = s.singer || '';
  mlEls.inputWork.value = s.work || '';
  mlEls.inputArtist.value = s.artist || '';
  mlEls.inputSong.value = s.song || '';
  mlEls.inputPart.checked = !!s.isPartDivision;
  mlEls.inputPrac.checked = !!s.isPracticing;
  mlEls.inputRev.checked = !!s.isReviewNeeded;
  mlEls.inputPrivate.checked = !!s.isPrivate;
  mlEls.inputPw.value = s.viewPassword || '';
  mlEls.privatePwWrap.style.display = mlEls.inputPrivate.checked ? 'block' : 'none';
  mlEditingId = id;
  mlEls.formTitle.innerText = '登録情報を編集';
  mlEls.formIcon.className = 'fas fa-edit';
  mlEls.submit.innerText = '更新する';
  mlEls.submit.classList.add('editing');
  mlEls.cancel.classList.add('show');
  mlRender();
  document.getElementById('mlFormCard').scrollIntoView({behavior:'smooth', block:'center'});
}
function mlCancelEdit(){ mlResetForm(); mlRender(); }
function mlResetForm(){
  mlEls.inputSong.value = ''; mlEls.inputWork.value = ''; mlEls.inputArtist.value = '';
  mlEls.inputPart.checked = false; mlEls.inputPrac.checked = false; mlEls.inputRev.checked = false;
  mlEls.inputPrivate.checked = false; mlEls.inputPw.value = '';
  mlEls.privatePwWrap.style.display = 'none';
  mlEditingId = null;
  mlEls.formTitle.innerText = '新規登録';
  mlEls.formIcon.className = 'fas fa-plus';
  mlEls.submit.innerText = 'リストに追加';
  mlEls.submit.classList.remove('editing');
  mlEls.cancel.classList.remove('show');
}
function mlDelete(id){
  showDialog('本当に削除しますか？', 'confirm', ()=>{
    mlSongs = mlSongs.filter(s=>s.id !== id);
    if(!mlDeletedIds.includes(id)){ mlDeletedIds.push(id); localStorage.setItem('deletedIds', JSON.stringify(mlDeletedIds)); }
    if(mlEditingId === id) mlCancelEdit();
    mlRender();
  });
}
function mlMove(id, direction){
  const globalIndex = mlSongs.findIndex(s=>s.id === id);
  if(globalIndex === -1) return;
  const vpw = mlEls.viewPw.value.trim();
  let cur = mlSongs.filter(s=>{ if(s.isPrivate && (!vpw || s.viewPassword !== vpw)) return false; return true; });
  if(mlActiveTab !== 'ALL') cur = cur.filter(s=>s.singer===mlActiveTab);
  if(mlSearchQuery){ const qs = mlSearchQuery.toLowerCase().replace(/　/g,' ').split(' ').filter(q=>q); cur = cur.filter(s=>{ const t=(s.song+' '+s.work+' '+s.artist).toLowerCase(); return qs.every(q=>t.includes(q)); }); }
  if(mlFilterTags.isPartDivision) cur = cur.filter(s=>s.isPartDivision);
  if(mlFilterTags.isPracticing) cur = cur.filter(s=>s.isPracticing);
  if(mlFilterTags.isReviewNeeded) cur = cur.filter(s=>s.isReviewNeeded);
  const visualIndex = cur.findIndex(s=>s.id===id);
  if(visualIndex === -1) return;
  let target = -1;
  if(direction==='up' && visualIndex < cur.length - 1) target = visualIndex + 1;
  if(direction==='down' && visualIndex > 0) target = visualIndex - 1;
  if(target !== -1){
    const swap = cur[target];
    const swapIndex = mlSongs.findIndex(s=>s.id === swap.id);
    [mlSongs[globalIndex], mlSongs[swapIndex]] = [mlSongs[swapIndex], mlSongs[globalIndex]];
    mlRender();
  }
}

// CSV modal
mlEls.csvBtn.addEventListener('click', ()=>document.getElementById('mlCsvOverlay').classList.add('active'));
document.getElementById('mlCsvClose').addEventListener('click', ()=>document.getElementById('mlCsvOverlay').classList.remove('active'));
document.getElementById('mlCsvOverlay').addEventListener('click', e=>{ if(e.target.id==='mlCsvOverlay') document.getElementById('mlCsvOverlay').classList.remove('active'); });
document.getElementById('mlCsvExportAll').addEventListener('click', ()=>mlExportCSV(true));
document.getElementById('mlCsvExportTpl').addEventListener('click', ()=>mlExportCSV(false));
document.getElementById('mlCsvImportBtn').addEventListener('click', ()=>mlEls.importCsv.click());
function mlExportCSV(withData){
  const BOM = '\uFEFF';
  const header = ['歌唱者','作品名','歌手','曲名','パート分け','練習中','要復習'];
  const rows = [header.join(',')];
  if(withData){
    const exportable = mlSongs.filter(s=>!s.isPrivate);
    exportable.forEach(s=>{
      const row = [s.singer||'', s.work||'', s.artist||'', s.song||'', s.isPartDivision?'1':'', s.isPracticing?'1':'', s.isReviewNeeded?'1':''].map(f=>{
        const v = String(f);
        return (v.includes(',')||v.includes('"')||v.includes('\n')) ? '"'+v.replace(/"/g,'""')+'"' : v;
      });
      rows.push(row.join(','));
    });
  }
  const csv = BOM + rows.join('\r\n');
  const blob = new Blob([csv], {type:'text/csv'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = withData ? ('karaoke_list_backup_'+new Date().toISOString().slice(0,10)+'.csv') : 'karaoke_list_template.csv';
  a.click();
  URL.revokeObjectURL(url);
  document.getElementById('mlCsvOverlay').classList.remove('active');
}

mlEls.importJsonBtn.addEventListener('click', ()=>mlEls.importJson.click());
mlEls.importJson.addEventListener('change', function(){
  const file = this.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = e=>{
    try {
      const d = JSON.parse(e.target.result);
      if(Array.isArray(d)){
        showDialog('ファイルから '+d.length+' 件のデータを読み込みました (JSON)。\n反映するには「同期して保存」を押してください。\n(次回保存時にサーバーのデータを完全に上書きします)', 'alert');
        mlSongs = d;
        mlForceOverwrite = true;
        mlRender();
      } else {
        showDialog('JSONデータの形式が正しくありません。', 'alert');
      }
    } catch(err){ showDialog('ファイルの読み込みに失敗しました。', 'alert'); }
  };
  reader.readAsText(file);
  this.value = '';
});

mlEls.importCsv.addEventListener('change', function(){
  const file = this.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = e=>{
    try { mlImportCSV(e.target.result); } catch(err){ showDialog('ファイルの読み込みに失敗しました。', 'alert'); }
  };
  reader.readAsText(file);
  this.value = '';
  document.getElementById('mlCsvOverlay').classList.remove('active');
});
function mlImportCSV(csvText){
  const rows = csvText.split(/\r\n|\n/);
  const dataRows = rows.slice(1).filter(r=>r.trim() !== '');
  let added = 0, updated = 0;
  dataRows.forEach(rowStr=>{
    let row = []; let inQ = false; let cur = '';
    for(let i=0;i<rowStr.length;i++){
      const c = rowStr[i];
      if(c === '"'){
        if(i+1 < rowStr.length && rowStr[i+1] === '"'){ cur += '"'; i++; }
        else inQ = !inQ;
      } else if(c === ',' && !inQ){ row.push(cur); cur = ''; }
      else { cur += c; }
    }
    row.push(cur);
    row = row.map(f=>f.trim());
    if(row.length < 4) return;
    const newSinger = row[0], newWork = row[1], newArtist = row[2];
    let newSong = row[3];
    if(!newSong) return;
    let isP = (row[4]==='1' || (row[4]||'').toLowerCase()==='true');
    let isPr = (row[5]==='1' || (row[5]||'').toLowerCase()==='true');
    let isR = (row[6]==='1' || (row[6]||'').toLowerCase()==='true');
    const patterns = [
      {regex: /[\(（]パート分け[\)）]/, key:'p'},
      {regex: /[\(（]練習中[\)）]/, key:'pr'},
      {regex: /[\(（]要復習[\)）]/, key:'r'}
    ];
    patterns.forEach(p=>{
      if(p.regex.test(newSong)){
        if(p.key==='p') isP = true;
        if(p.key==='pr') isPr = true;
        if(p.key==='r') isR = true;
        newSong = newSong.replace(p.regex, '').trim();
      }
    });
    const ns = mlNormalizeStr(newSinger), nso = mlNormalizeStr(newSong);
    const idx = mlSongs.findIndex(s=>mlNormalizeStr(s.singer)===ns && mlNormalizeStr(s.song)===nso);
    if(idx !== -1){
      mlSongs[idx] = Object.assign({}, mlSongs[idx], {work:newWork, artist:newArtist, isPartDivision:isP, isPracticing:isPr, isReviewNeeded:isR});
      updated++;
    } else {
      mlSongs.push({id: Date.now() + Math.random(), singer:newSinger, work:newWork, artist:newArtist, song:newSong, isPartDivision:isP, isPracticing:isPr, isReviewNeeded:isR, isPrivate:false, viewPassword:''});
      added++;
    }
  });
  showDialog('CSVインポート完了:\n追加: '+added+'件\n更新: '+updated+'件\n反映するには「同期して保存」を押してください。', 'alert');
  mlRender();
}

// Sync
mlEls.syncBtn.addEventListener('click', mlSync);
async function mlSync(){
  if(!MYLIST_GAS_URL){ showDialog('GASのURLが設定されていません。', 'alert'); return; }
  mlUpdateStatus('サーバーと同期中...', true);
  try {
    let merged = mlSongs;
    if(!mlForceOverwrite){
      let server = [];
      try {
        const res = await fetch(MYLIST_GAS_URL + '?t=' + Date.now());
        if(!res.ok) throw new Error('Fetch failed');
        const text = await res.text();
        const data = JSON.parse(text);
        if(Array.isArray(data)) server = data;
      } catch(e){
        showDialog('サーバーからのデータ取得に失敗しました。保存を中止します。\n(インターネット接続やGASの状態を確認してください)', 'alert');
        mlUpdateStatus('保存中止');
        return;
      }
      const localIds = new Set(mlSongs.map(s=>s.id));
      const newServer = server.filter(s=>!localIds.has(s.id) && !mlDeletedIds.includes(s.id));
      merged = mlSongs.concat(newServer);
    }
    mlSongs = merged; mlRender();
    await fetch(MYLIST_GAS_URL, {method:'POST', body: JSON.stringify(merged), headers:{'Content-Type':'text/plain;charset=utf-8'}});
    mlDeletedIds = []; localStorage.setItem('deletedIds', JSON.stringify(mlDeletedIds));
    mlForceOverwrite = false;
    mlUpdateStatus('同期・保存完了！');
    setTimeout(()=>mlUpdateStatus(''), 3000);
  } catch(err){
    showDialog('保存エラー: '+err.message, 'alert');
    mlUpdateStatus('保存失敗');
  }
}
async function mlLoadFromGAS(silent){
  if(!MYLIST_GAS_URL) return;
  mlUpdateStatus('読み込み中...', true);
  try {
    const res = await fetch(MYLIST_GAS_URL + '?t=' + Date.now());
    const text = await res.text();
    let data;
    try { data = JSON.parse(text); } catch(e){ throw new Error('サーバーからの応答が不正です。'); }
    if(Array.isArray(data)){
      mlSongs = data; mlDeletedIds = []; localStorage.setItem('deletedIds', JSON.stringify(mlDeletedIds));
      mlActiveTab = 'ALL'; mlCurrentPage = 1; mlRender();
      mlUpdateStatus('読み込み完了'); setTimeout(()=>mlUpdateStatus(''), 3000);
    } else throw new Error('データ形式が不正です');
  } catch(err){
    if(!silent) showDialog('【通信エラー】\n'+err.message, 'alert');
    mlUpdateStatus('エラー');
  }
}
function mlUpdateStatus(msg, animate){
  mlEls.status.innerText = msg || '';
}

// ============================== Custom dialog ==============================
function showDialog(message, type, onConfirm){
  const ov = document.getElementById('cdOverlay');
  document.getElementById('cdMsg').innerText = message;
  const actions = document.getElementById('cdActions');
  actions.innerHTML = '';
  const close = ()=>ov.classList.remove('active');
  if(type === 'confirm'){
    const cancel = document.createElement('button');
    cancel.className = 'cancel'; cancel.innerText = 'キャンセル'; cancel.onclick = close;
    const ok = document.createElement('button');
    ok.className = 'ok'; ok.innerText = 'OK'; ok.onclick = ()=>{ close(); if(onConfirm) onConfirm(); };
    actions.appendChild(cancel); actions.appendChild(ok);
  } else {
    const ok = document.createElement('button');
    ok.className = 'ok'; ok.innerText = 'OK'; ok.onclick = close;
    actions.appendChild(ok);
  }
  ov.classList.add('active');
}

// ============================== Initial render ==============================
renderCool();
renderRanking();
renderTrend();
mlRender();
// Initial mylist load from GAS if local empty
if(mlSongs.length === 0){ mlLoadFromGAS(true); }
