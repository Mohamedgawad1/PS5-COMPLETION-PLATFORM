(() => {
  const LS = 'live_itr_state';
  function fmt(n){ return (n||0).toLocaleString('en-US'); }
  function build(update){
    const e = document.querySelector('#itr-live-badge');
    if(!e) return;
    const cb = update.closed, eb = (update.closed_by_discipline||{}).E||0,
          ib = (update.closed_by_discipline||{}).I||0, tb = (update.closed_by_discipline||{}).T||0;
    e.innerHTML =
      '<div class="itr-badge-title">ITR Closed &middot; Live</div>' +
      '<div class="itr-badge-big">' + fmt(cb) + '</div>' +
      '<div class="itr-badge-sub">E ' + fmt(eb) + ' &middot; I ' + fmt(ib) + ' &middot; T ' + fmt(tb) +
      ' &middot; Total E&I&T ' + fmt(update.eit_total) + '</div>' +
      '<div class="itr-badge-time">Updated ' + (update.updated||'') + '</div>';
  }
  function syncCards(update){
    const set = (id, v) => { const el = document.getElementById(id); if(el) el.textContent = v; };
    const tot = update.closed||0, all = update.eit_total||0, today = update.today_closed||0;
    set('todayClosed', fmt(today));
    set('totalClosed', fmt(tot));
    set('totalPct', all ? (tot/all*100).toFixed(2) : '0.00');
    document.querySelectorAll('.kpi').forEach(c => {
      const lbl = c.querySelector('.lbl'); if(!lbl) return;
      const val = c.querySelector('.val'); if(!val) return;
      const t = lbl.textContent||'';
      if(t.indexOf('CPP AGI EIT') !== -1 && t.indexOf('Total ITRs Closed') !== -1){
        val.textContent = fmt(tot) + ' / ' + fmt(all);
      } else if(t.trim() === 'Today Closed'){
        val.textContent = fmt(today);
      } else if(t.trim() === 'Closed This Week'){
        val.textContent = fmt(update.closed_this_week);
      } else if(t.trim() === 'Closed This Month'){
        val.textContent = fmt(update.closed_this_month);
      }
    });
  }
  function esc(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  function discChip(d){
    const c = (d||'').toLowerCase();
    const cls = c.indexOf('e')===0 ? '#065f46' : c.indexOf('i')===0 ? '#78350f' : c.indexOf('t')===0 ? '#831843' : '#1e293b';
    const col = c.indexOf('e')===0 ? '#a7f3d0' : c.indexOf('i')===0 ? '#fde68a' : c.indexOf('t')===0 ? '#fbcfe8' : '#cbd5e1';
    return '<span style="display:inline-block;padding:1px 8px;border-radius:10px;font-size:10px;background:'+cls+';color:'+col+'">'+esc(d)+'</span>';
  }
  function syncRecent(update){
    const list = Array.isArray(update.recent_closed) ? update.recent_closed : [];
    let sec = document.getElementById('live-recent-sec');
    if(!sec){
      sec = document.createElement('section');
      sec.id = 'live-recent-sec';
      sec.className = 'section active';
      sec.style.border = '1px solid var(--border)';
      sec.style.background = 'var(--panel2)';
      sec.style.borderRadius = '8px';
      sec.style.padding = '14px';
      sec.style.marginBottom = '16px';
      sec.innerHTML =
        '<div class="section-title" style="margin:0 0 10px">🕒 Latest ITR Closures &mdash; Live (Subsystem / Asset / Task)</div>' +
        '<div class="wrap" style="overflow-x:auto">' +
        '<table id="live-recent-tbl" style="width:100%;border-collapse:collapse;font-size:12px;white-space:nowrap">' +
        '<thead><tr style="background:var(--panel);text-transform:uppercase;font-size:10px;letter-spacing:.5px;color:var(--teal)">' +
        '<th style="padding:7px 9px;border:1px solid var(--border);text-align:left">Task ID</th>' +
        '<th style="padding:7px 9px;border:1px solid var(--border);text-align:left">Asset Tag</th>' +
        '<th style="padding:7px 9px;border:1px solid var(--border);text-align:left">Subsystem</th>' +
        '<th style="padding:7px 9px;border:1px solid var(--border);text-align:left">Loop</th>' +
        '<th style="padding:7px 9px;border:1px solid var(--border);text-align:left">Disc</th>' +
        '<th style="padding:7px 9px;border:1px solid var(--border);text-align:left">Category</th>' +
        '<th style="padding:7px 9px;border:1px solid var(--border);text-align:left">Approved</th>' +
        '</tr></thead><tbody></tbody></table></div>' +
        '<div style="font-size:11px;color:var(--muted);margin-top:8px">Live from <b>itr_live_state.json</b> &mdash; updates automatically every 2 min. Time shown: <b id="live-recent-time">&mdash;</b></div>';
      const kpi = document.getElementById('sec-kpi');
      if(kpi && kpi.parentElement) kpi.parentElement.insertBefore(sec, kpi.nextSibling);
      else document.body.appendChild(sec);
    }
    const tb = sec.querySelector('#live-recent-tbl tbody');
    const tm = sec.querySelector('#live-recent-time');
    if(tm) tm.textContent = update.updated || '—';
    if(tb){
      tb.innerHTML = list.map(r =>
        '<tr>' +
        '<td style="padding:6px 9px;border:1px solid var(--border)">'+esc(r.task)+'</td>' +
        '<td style="padding:6px 9px;border:1px solid var(--border);color:var(--teal);font-weight:600">'+esc(r.tag)+'</td>' +
        '<td style="padding:6px 9px;border:1px solid var(--border)">'+esc(r.system)+'</td>' +
        '<td style="padding:6px 9px;border:1px solid var(--border)">'+esc(r.loop)+'</td>' +
        '<td style="padding:6px 9px;border:1px solid var(--border)">'+discChip(r.disc)+'</td>' +
        '<td style="padding:6px 9px;border:1px solid var(--border)">'+esc(r.cat)+'</td>' +
        '<td style="padding:6px 9px;border:1px solid var(--border);color:var(--muted)">'+esc(r.approved)+'</td>' +
        '</tr>'
      ).join('') || '<tr><td colspan="7" style="padding:10px;text-align:center;color:var(--muted)">No closures yet &mdash; waiting for live data&hellip;</td></tr>';
    }
  }
  function syncChart(update){
    const daily = Array.isArray(update.daily) ? update.daily : null;
    if(!daily || !daily.length) return;
    const labels = daily.map(d=>d.label);
    const tot = daily.map(d=>d.Total||0);
    const upd = (id, apply) => {
      const cv = document.getElementById(id);
      if(!cv) return;
      let ch = null;
      try{ ch = window.Chart && Chart.getChart ? Chart.getChart(cv) : null; }catch(e){}
      if(!ch || !ch.data || !ch.data.datasets) return;
      ch.data.labels = labels;
      apply(ch.data.datasets);
      ch.update();
    };
    upd('chartCombinedDaily', ds => { if(ds[0]) ds[0].data = tot; });
    upd('chartDaily', ds => {
      ['E','I','T'].forEach((k,i)=>{ if(ds[i]) ds[i].data = daily.map(d=>d[k]||0); });
    });
  }
  function syncMilestone(update){
    const t = document.getElementById('todayMsTotal');
    if(t) t.textContent = fmt(update.today_milestone_total||0);
    const body = document.getElementById('todayMsBody');
    if(!body) return;
    const tm = Array.isArray(update.today_milestone) ? update.today_milestone : [];
    if(!tm.length){
      if(!body.dataset._live) body.innerHTML = '<tr><td colspan="5" style="padding:10px;text-align:center;color:var(--muted)">No closures today yet &mdash; waiting for live data&hellip;</td></tr>';
      return;
    }
    body.dataset._live = '1';
    const msMap = {};
    tm.forEach(r => {
      if(!msMap[r.milestone]) msMap[r.milestone] = {E:0, I:0, T:0};
      const d = String(r.disc||'');
      if(msMap[r.milestone][d] !== undefined) msMap[r.milestone][d] = r.count;
    });
    body.innerHTML = Object.keys(msMap).sort().map(ms => {
      const d = msMap[ms];
      const rowTotal = d.E + d.I + d.T;
      return '<tr>' +
        '<td style="text-align:left;border:1px solid var(--border);font-weight:bold;color:var(--red);background:var(--card3);">'+esc(ms)+'</td>' +
        '<td style="text-align:center;border:1px solid var(--border);background:#FCE4D6;color:var(--red);">'+fmt(d.E)+'</td>' +
        '<td style="text-align:center;border:1px solid var(--border);background:#FFF2CC;color:var(--red);">'+fmt(d.I)+'</td>' +
        '<td style="text-align:center;border:1px solid var(--border);background:#D9D9D9;color:var(--red);">'+fmt(d.T)+'</td>' +
        '<td style="text-align:center;border:1px solid var(--border);font-weight:bold;color:var(--red);background:var(--card3);">'+fmt(rowTotal)+'</td>' +
        '</tr>';
    }).join('');
  }
  const MS_ORDER = ['Closed','Submitted','To be completed','Other'];
  const MS_COLORS = {'Closed':'#1a8a4a','Submitted':'#2563eb','To be completed':'#c8940a','Other':'#9a8d7c'};
  function syncMilestoneSummary(update){
    const list = Array.isArray(update.milestone_summary) ? update.milestone_summary : [];
    const cards = document.getElementById('msCards');
    if(cards){
      cards.innerHTML = list.map(s =>
        '<div class="kpi">' +
          '<div class="val">' + fmt(s.closed) + ' / ' + fmt(s.total) + '</div>' +
          '<div class="lbl">🎯 ' + esc(s.label) + ' &mdash; ' + s.pct + '%</div>' +
          '<div class="progress-bar"><div class="progress-fill" style="width:' + s.pct + '%"></div></div>' +
        '</div>'
      ).join('') ||
        '<div class="kpi"><div class="val">0 / 0</div><div class="lbl">🎯 &mdash;</div><div class="progress-bar"><div class="progress-fill" style="width:0%"></div></div></div>';
    }
    const cv = document.getElementById('chartMilestone');
    if(!cv || !list.length) return;
    let ch = null;
    try{ ch = window.Chart && Chart.getChart ? Chart.getChart(cv) : null; }catch(e){}
    if(!ch || !ch.data || !ch.data.datasets) return;
    ch.data.labels = list.map(s=>s.label);
    MS_ORDER.forEach((st,i)=>{
      let ds = ch.data.datasets[i];
      if(!ds || ds.label !== st) ds = ch.data.datasets.find(d=>d.label===st);
      if(ds) ds.data = list.map(s=> (s.status&&s.status[st]) || 0);
    });
    ch.update();
  }
  function fetchItr(){
    fetch('itr_live_state.json?v=' + Math.floor(Date.now()/120000), {cache:'no-store'})
      .then(r => r.json())
      .then(u => { try{ localStorage.setItem(LS, JSON.stringify(u)); }catch(e){} build(u); syncCards(u); syncRecent(u); syncChart(u); syncMilestone(u); syncMilestoneSummary(u); })
      .catch(() => { try{ const o=localStorage.getItem(LS); if(o){ const u=JSON.parse(o); build(u); syncCards(u); syncRecent(u); syncChart(u); syncMilestone(u); syncMilestoneSummary(u); } }catch(e){} });
  }
  if(!document.querySelector('#itr-live-css')) {
    const st = document.createElement('style'); st.id='itr-live-css';
    st.textContent = '#itr-live-badge{position:fixed;right:16px;bottom:16px;z-index:99999;background:#0b2f56;color:#fff;border:1px solid #38bdf8;border-radius:12px;padding:10px 14px;font-family:Segoe UI,Arial,sans-serif;box-shadow:0 6px 18px rgba(0,0,0,.35);min-width:210px}'
      + '.itr-badge-title{font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#7dd3fc}'
      + '.itr-badge-big{font-size:30px;font-weight:800;line-height:1.1}'
      + '.itr-badge-sub{font-size:12px;color:#cbd5e1;margin-top:2px}'
      + '.itr-badge-time{font-size:10px;color:#94a3b8;margin-top:4px}';
    document.head.appendChild(st);
  }
  if(!document.querySelector('#itr-live-badge')){
    const d = document.createElement('div'); d.id='itr-live-badge';
    d.innerHTML = '<div class="itr-badge-title">ITR Closed &middot; Live</div><div class="itr-badge-big">&hellip;</div>';
    document.body.appendChild(d);
  }
  fetchItr();
  setInterval(fetchItr, 120000);
})();