(() => {
  // Live overlay for the PUBLIC PS5 platform page.
  // The platform ships a static snapshot; this script pulls the fresh
  // itr_live_state.json every minute, flips each ITR row to its live
  // Closed/Open state, re-renders the KPIs, and shows a per-subsystem table.
  const LS = 'live_itr_state_v2';
  const REFRESH = 60000;
  let lastSig = '';
  let subState = null;

  const fmt = n => (Number(n) || 0).toLocaleString('en-US');
  const esc = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  function ensureBadge() {
    let el = document.getElementById('live-status-pill');
    if (el) return el;
    el = document.createElement('div');
    el.id = 'live-status-pill';
    el.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:99999;background:#0b2f56;color:#fff;' +
      'border:1px solid #38bdf8;border-radius:12px;padding:10px 14px;font-family:Segoe UI,Arial,sans-serif;' +
      'box-shadow:0 6px 18px rgba(0,0,0,.35);min-width:230px';
    document.body.appendChild(el);
    return el;
  }

  function paintBadge(u, ok, err) {
    const el = ensureBadge();
    if (!ok) {
      el.innerHTML = '<div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#7dd3fc">ITR LIVE</div>' +
        '<div style="font-size:13px;color:#fca5a5;margin-top:4px">تعذر التحديث' + (err ? ': ' + esc(err) : '') + '</div>' +
        '<div style="font-size:10px;color:#94a3b8;margin-top:4px">يعيد المحاولة تلقائياً…</div>';
      return;
    }
    const e = (u.closed_by_discipline || {}).E || 0;
    const i = (u.closed_by_discipline || {}).I || 0;
    const t = (u.closed_by_discipline || {}).T || 0;
    el.innerHTML =
      '<div style="font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#7dd3fc">ITR LIVE &middot; Smart Completions</div>' +
      '<div style="font-size:26px;font-weight:800;line-height:1.1;margin-top:2px">' + fmt(u.closed) +
        ' <span style="font-size:13px;color:#94a3b8;font-weight:400">closed</span></div>' +
      '<div style="font-size:12px;color:#fca5a5">' + fmt(u.open) + ' open &middot; ' + fmt(u.eit_total) + ' total E&amp;I&amp;T</div>' +
      '<div style="font-size:11px;color:#cbd5e1;margin-top:2px">E ' + fmt(e) + ' &middot; I ' + fmt(i) + ' &middot; T ' + fmt(t) + '</div>' +
      '<div style="font-size:10px;color:#94a3b8;margin-top:4px">Updated ' + esc(u.updated) + '</div>';
  }

  // The platform keeps every ITR in the global ITR array: [id, tag, disc, type,
  // desc, state, close, scope, subname, sid]. The live Smart Completions pull is
  // exactly the CPP AGI scope, so only those rows are flipped; EACOP / EITS / KENT
  // keep their own static values.
  function applyTaskStates(u) {
    const ts = u.task_states || {};
    let changed = 0, matched = 0;
    try {
      for (let k = 0; k < ITR.length; k++) {
        const row = ITR[k];
        if (String(row[7] == null ? '' : row[7]).trim().toUpperCase() !== 'CPP AGI') continue;
        const st = ts[row[0]];
        if (!st) continue;
        matched++;
        if (row[5] !== st) { row[5] = st; changed++; }
      }
    } catch (e) { /* platform not ready */ }
    return { changed: changed, matched: matched };
  }

  function ensureSubPanel() {
    let p = document.getElementById('live-sub-panel');
    if (p) return p;
    const host = document.getElementById('v-dash');
    if (!host) return null;
    p = document.createElement('div');
    p.className = 'panel';
    p.id = 'live-sub-panel';
    p.innerHTML =
      '<h2>ITR LIVE &mdash; Closed / Open by Subsystem <span id="live-sub-meta" style="font-size:11px;color:#6b7280;font-weight:400"></span></h2>' +
      '<div class="tools"><input type="search" id="live-sub-q" placeholder="Search subsystem...">' +
      '<span class="hint" id="live-sub-hint" style="margin:0"></span></div>' +
      '<div class="tblwrap" style="max-height:520px"><table id="live-sub-tbl"></table></div>';
    const kp = document.getElementById('dash-kpis');
    if (kp && kp.parentElement === host) host.insertBefore(p, kp.nextSibling);
    else host.insertBefore(p, host.firstChild);
    p.querySelector('#live-sub-q').addEventListener('input', () => {
      if (subState) renderSubTable(subState);
    });
    p.querySelector('#live-sub-tbl').addEventListener('click', ev => {
      const tr = ev.target.closest('tr.live-sub-row');
      if (tr && tr.dataset.sid && typeof goSub === 'function') goSub(tr.dataset.sid);
    });
    return p;
  }

  function renderSubTable(u) {
    subState = u;
    const p = ensureSubPanel();
    if (!p) return;
    const list = Array.isArray(u.subsystem_summary) ? u.subsystem_summary.slice() : [];
    const q = (p.querySelector('#live-sub-q').value || '').trim().toLowerCase();
    const f = q ? list.filter(s => ((s.name || '') + ' ' + (s.sid || '')).toLowerCase().indexOf(q) !== -1) : list;
    f.sort((a, b) => String(a.sid).localeCompare(String(b.sid), 'en', { numeric: true }));
    p.querySelector('#live-sub-meta').textContent =
      '  ' + fmt(u.closed) + ' closed / ' + fmt(u.open) + ' open  •  updated ' + (u.updated || '—');
    let h = '<thead><tr><th>SUB SYSTEM</th>' +
      '<th style="text-align:center">TOTAL</th><th style="text-align:center">CLOSED</th>' +
      '<th style="text-align:center">OPEN</th><th style="text-align:center">%</th><th>STATUS</th></tr></thead><tbody>';
    if (!f.length) {
      h += '<tr><td colspan="6"><div class="empty">No subsystems match</div></td></tr>';
    } else {
      for (const s of f) {
        const st = s.status || '';
        const col = st === 'COMPLETE' ? '#1a8a4a' : st === 'OPEN' ? '#c0392b' : '#b7791f';
        const bg = st === 'COMPLETE' ? '#e6f6ec' : st === 'OPEN' ? '#fdeceb' : '#fff7e6';
        const bd = s.by_discipline || {};
        const cd = s.closed_by_discipline || {};
        const nm = String(s.name || '').replace(/^[^-]+ - /, '');
        h += '<tr class="live-sub-row" data-sid="' + esc(s.sid) + '" style="cursor:pointer">' +
          '<td><b>' + esc(s.sid) + '</b> &middot; ' + esc(nm) +
            '<div style="font-size:10px;color:#6b7280">E ' + fmt(cd.E) + '/' + fmt(bd.E) +
            ' &nbsp; I ' + fmt(cd.I) + '/' + fmt(bd.I) + ' &nbsp; T ' + fmt(cd.T) + '/' + fmt(bd.T) + '</div></td>' +
          '<td style="text-align:center">' + fmt(s.total) + '</td>' +
          '<td style="text-align:center;color:#1a8a4a;font-weight:700">' + fmt(s.closed) + '</td>' +
          '<td style="text-align:center;color:#c0392b;font-weight:700">' + fmt(s.open) + '</td>' +
          '<td style="text-align:center">' + (s.pct == null ? '' : s.pct + '%') + '</td>' +
          '<td><span style="display:inline-block;padding:2px 9px;border-radius:10px;font-size:10px;' +
            'font-weight:700;color:' + col + ';background:' + bg + ';border:1px solid ' + col + '55">' +
            esc(st) + '</span></td></tr>';
      }
    }
    h += '</tbody>';
    p.querySelector('#live-sub-tbl').innerHTML = h;
    p.querySelector('#live-sub-hint').textContent = f.length + ' / ' + list.length + ' SUBSYSTEMS';
  }

  function apply(u) {
    const res = applyTaskStates(u);
    const sig = [u.updated, u.closed, u.open, res.changed].join('|');
    if (sig !== lastSig) {
      lastSig = sig;
      try { if (typeof renderDash === 'function') renderDash(); } catch (e) {}
      try {
        if (typeof V !== 'undefined') {
          if (V === 'it' && typeof renderList === 'function') renderList('it');
          else if (V === 'sub' && typeof renderSub === 'function') renderSub();
          else if (V === 'blk' && typeof renderBlk === 'function') renderBlk();
        }
      } catch (e) {}
    }
    renderSubTable(u);
    paintBadge(u, true);
    try { localStorage.setItem(LS, JSON.stringify({ updated: u.updated, closed: u.closed, open: u.open })); } catch (e) {}
  }

  function load() {
    fetch('itr_live_state.json?v=' + Math.floor(Date.now() / 30000), { cache: 'no-store' })
      .then(r => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
      .then(u => { if (u && u.updated) apply(u); else throw new Error('empty state'); })
      .catch(e => paintBadge(null, false, String((e && e.message) || e)));
  }

  load();
  setInterval(load, REFRESH);
})();