/* MyMoney phone app – reads data.json that Excel (AppLink macro) pushes to a private
   GitHub repo, and writes phone entries to inbox/<id>.json for Excel to import. */
'use strict';

/* ---------------- storage ---------------- */
const LS = {
  get(k, d) { try { const v = localStorage.getItem('mm.' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
  set(k, v) { try { localStorage.setItem('mm.' + k, JSON.stringify(v)); } catch (e) { /* storage full or blocked */ } },
  del(k) { try { localStorage.removeItem('mm.' + k); } catch (e) {} },
};
const state = {
  cfg: LS.get('cfg', null),          // {owner, repo, token, branch}
  data: LS.get('data', null),        // snapshot from Excel
  fetchedAt: LS.get('fetchedAt', null),
  sheets: LS.get('sheets', null),   // every sheet as Excel shows it
  sheet: null, sheetQ: '', choice: {},   // choice[sheet] = {label, v}
  sheetMode: LS.get('sheetMode', 'cards'),
  pending: LS.get('pending', []),    // [{id, sha, entry}] in GitHub inbox, not yet in Excel data
  outbox: LS.get('outbox', []),      // entries saved offline, not uploaded yet
  tab: 'home', lastError: null, syncing: false,
  tx: { q: '', type: 'All', year: 'All', month: 'All', cat: 'All', limit: 100 },
  add: null,
};

/* ---------------- GitHub ---------------- */
const GH = {
  url(path, withRef) {
    const c = state.cfg;
    return `https://api.github.com/repos/${encodeURIComponent(c.owner)}/${encodeURIComponent(c.repo)}/contents/${path}` +
      (withRef ? `?ref=${encodeURIComponent(c.branch || 'main')}` : '');
  },
  async call(method, path, { body, raw } = {}) {
    const res = await fetch(this.url(path, method === 'GET'), {
      method, cache: 'no-store',
      headers: {
        Authorization: 'Bearer ' + state.cfg.token,
        Accept: raw ? 'application/vnd.github.raw+json' : 'application/vnd.github+json',
        ...(body ? { 'Content-Type': 'application/json' } : {}),
      },
      body: body ? JSON.stringify(body) : undefined,
    });
    return res;
  },
  async getData() {
    const r = await this.call('GET', 'data.json', { raw: true });
    if (r.status === 404) throw new AppError('No data yet. In Excel, press Alt+F8 and run AppLinkNow once.');
    if (!r.ok) throw httpError(r.status);
    return r.json();
  },
  async getSheets() {
    const r = await this.call('GET', 'sheets.json', { raw: true });
    if (r.status === 404) return null;
    if (!r.ok) throw httpError(r.status);
    return r.json();
  },
  async listInbox() {
    const r = await this.call('GET', 'inbox');
    if (r.status === 404) return [];
    if (!r.ok) throw httpError(r.status);
    const arr = await r.json();
    return Array.isArray(arr) ? arr.filter(f => f.type === 'file' && /\.json$/.test(f.name)) : [];
  },
  async getEntry(path) {
    const r = await this.call('GET', path, { raw: true });
    if (!r.ok) throw httpError(r.status);
    return r.json();
  },
  async putEntry(entry) {
    const r = await this.call('PUT', `inbox/${entry.id}.json`, { body: {
      message: `Phone entry ${entry.id}`, branch: state.cfg.branch || 'main',
      content: b64(JSON.stringify(entry, null, 1)) } });
    if (!r.ok) throw httpError(r.status);
    const j = await r.json();
    return j.content && j.content.sha;
  },
  async deleteEntry(id, sha) {
    const r = await this.call('DELETE', `inbox/${id}.json`, { body: {
      message: `Phone entry removed ${id}`, branch: state.cfg.branch || 'main', sha } });
    if (!r.ok && r.status !== 404) throw httpError(r.status);
  },
  async testRepo() {
    const c = state.cfg;
    const r = await fetch(`https://api.github.com/repos/${encodeURIComponent(c.owner)}/${encodeURIComponent(c.repo)}`,
      { cache: 'no-store', headers: { Authorization: 'Bearer ' + c.token, Accept: 'application/vnd.github+json' } });
    if (!r.ok) throw httpError(r.status);
    return r.json();
  },
};
class AppError extends Error {}
function httpError(st) {
  if (st === 401) return new AppError('GitHub rejected the token (401). Open ⚙ and paste a new token.');
  if (st === 403) return new AppError('GitHub refused (403). The token needs “Contents: Read and write” on the data repo.');
  if (st === 404) return new AppError('GitHub can’t find the repo (404). Check the names in ⚙ and that the token can see it.');
  if (st === 409 || st === 422) return new AppError(`GitHub reported a conflict (${st}). Try again.`);
  return new AppError(`GitHub returned an error (${st}).`);
}
function b64(str) {
  const bytes = new TextEncoder().encode(str);
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

/* ---------------- data helpers ---------------- */
function tbl(name) {
  const t = state.data && state.data.tables && state.data.tables[name];
  if (!t) return [];
  return t.rows.map(r => Object.fromEntries(t.cols.map((c, i) => [c, r[i]])));
}
function kpis() {
  const m = {};
  tbl('AppKPIs').forEach(r => { m[r.Key] = r.Value; });
  return m;
}
const n = v => (typeof v === 'number' && isFinite(v)) ? v : (v === '' || v == null ? NaN : Number(v));
const inrFmt = new Intl.NumberFormat('en-IN', { maximumFractionDigits: 0 });
function inr(v, sign) {
  v = n(v); if (!isFinite(v)) return '–';
  const s = '₹' + inrFmt.format(Math.abs(Math.round(v)));
  if (v < 0) return '-' + s;
  return (sign && v > 0 ? '+' : '') + s;
}
function cr(v) {
  v = n(v); if (!isFinite(v)) return '–';
  if (Math.abs(v) >= 1e7) return '₹' + (v / 1e7).toFixed(1) + ' Cr';
  if (Math.abs(v) >= 1e5) return '₹' + (v / 1e5).toFixed(1) + ' L';
  return inr(v);
}
const pct = v => isFinite(n(v)) ? (n(v) * 100).toFixed(1).replace(/\.0$/, '') + '%' : '–';
function pDate(s) { if (!s) return null; const [y, m, d] = String(s).slice(0, 10).split('-').map(Number); return new Date(y, m - 1, d || 1); }
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const fDay = d => d ? `${d.getDate()} ${MON[d.getMonth()]} ${d.getFullYear()}` : '';
const fMon = d => d ? `${MON[d.getMonth()]} ${d.getFullYear()}` : '';
function isoToday() { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function tone(s) { s = String(s || ''); return /^[✓✅]/.test(s) ? 'good' : /^⚠/.test(s) ? (/Fix|Short|Runs out|off by|overdue/i.test(s) ? 'bad' : 'warn') : ''; }

/* Make bank details readable: "bank: {SBI} UPI/DR/123/SWIGGY/ICIC/... [key]" -> "Swiggy" */
function describe(t) {
  let d = String(t.Details || '');
  const bank = (d.match(/^bank:\s*\{([^}]*)\}/i) || [])[1] || '';
  d = d.replace(/^bank:\s*(\{[^}]*\})?\s*/i, '').replace(/\s*\[[^\]]*\]\s*$/, '');
  const app = (d.match(/\s*app#([A-Za-z0-9-]+)\s*$/) || [])[1] || '';
  d = d.replace(/\s*app#[A-Za-z0-9-]+\s*$/, '');
  let title = d;
  const upi = d.match(/UPI\/(?:DR|CR)\/\d+\/([^/]+)/i) || d.match(/UPI-([A-Za-z][A-Za-z .]{2,})/);
  if (upi) title = upi[1];
  title = title.replace(/\s+/g, ' ').trim();
  if (/^[A-Z0-9 .&\-]+$/.test(title) && title.length > 3) title = title.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
  if (!title || /^from payslip/i.test(title)) title = t.Category || 'Entry';
  if (title.length > 42) title = title.slice(0, 40) + '…';
  let src = String(t.Source || '');
  src = bank || (src.startsWith('Bank Statement - ') ? src.slice(17) : src);
  if (app) src = 'Phone entry';
  return { title, src, app };
}
const TYPE_ICON = { Income: '↓', Expenses: '↑', Savings: '◆', Transfer: '⇄' };

/* ---------------- sync ---------------- */
async function sync(opts = {}) {
  if (!state.cfg || state.syncing) return;
  state.syncing = true; document.getElementById('btn-sync').classList.add('spin');
  try {
    await flushOutbox();
    const data = await GH.getData();
    state.data = data; state.fetchedAt = new Date().toISOString(); state.lastError = null;
    LS.set('data', data); LS.set('fetchedAt', state.fetchedAt);
    try { const sh = await GH.getSheets(); if (sh) { state.sheets = sh; LS.set('sheets', sh); } } catch (e) { /* sheets are optional */ }
    await refreshPending();
    if (opts.toast) toast('Up to date');
  } catch (e) {
    state.lastError = (e instanceof AppError) ? e.message
      : (navigator.onLine === false ? 'You’re offline – showing the last data loaded.' : 'Couldn’t reach GitHub – showing the last data loaded.');
  } finally {
    state.syncing = false; document.getElementById('btn-sync').classList.remove('spin');
    render();
  }
}
function importedIds() {
  const s = new Set();
  tbl('AppRecent').forEach(t => { const m = String(t.Details || '').match(/app#([A-Za-z0-9-]+)/); if (m) s.add(m[1]); });
  return s;
}
async function refreshPending() {
  const files = await GH.listInbox();
  const done = importedIds();
  const known = Object.fromEntries(state.pending.map(p => [p.id, p]));
  const out = [];
  for (const f of files) {
    const id = f.name.replace(/\.json$/, '');
    if (done.has(id)) continue;                         // already in Excel, waiting for a save to clear
    let p = known[id];
    if (!p || p.sha !== f.sha) {
      try { p = { id, sha: f.sha, entry: await GH.getEntry(f.path) }; } catch (e) { continue; }
    }
    out.push(p);
  }
  state.pending = out; LS.set('pending', out);
}
async function flushOutbox() {
  if (!state.outbox.length) return;
  const left = [];
  for (const e of state.outbox) {
    try { const sha = await GH.putEntry(e); state.pending.push({ id: e.id, sha, entry: e }); }
    catch (err) { left.push(e); }
  }
  state.outbox = left; LS.set('outbox', left); LS.set('pending', state.pending);
  if (left.length) throw new AppError(`${left.length} entr${left.length === 1 ? 'y is' : 'ies are'} still waiting to upload.`);
}

/* ---------------- render ---------------- */
const TITLES = { home: 'MyMoney', tx: 'Activity', add: 'Add entry', budget: 'Budget', future: 'Retirement', settings: 'Settings', more: 'More', sheet: 'Sheet' };
const NAV_OF = { future: 'more', sheet: 'more' };
function render() {
  const v = document.getElementById('view');
  document.getElementById('screen-title').textContent = state.tab === 'sheet' ? state.sheet : TITLES[state.tab];
  document.querySelectorAll('.tab').forEach(b => b.classList.toggle('active', b.dataset.tab === (NAV_OF[state.tab] || state.tab)));
  renderFreshness();
  if (!state.cfg && state.tab !== 'settings') state.tab = 'settings';
  if (state.tab === 'settings') { v.innerHTML = viewSettings(); bindSettings(); return; }
  if (!state.data) {
    v.innerHTML = `<div class="empty">${state.syncing ? 'Loading your data…' : esc(state.lastError || 'No data loaded yet.')}<br><br>
      <button class="secondary" onclick="sync({toast:true})">Try again</button></div>`;
    return;
  }
  const views = { home: viewHome, tx: viewTx, add: viewAdd, budget: viewBudget, future: viewFuture, more: viewMore, sheet: viewSheet };
  v.innerHTML = views[state.tab]();
  if (state.tab === 'more' || state.tab === 'sheet') bindSheets();
  if (state.tab === 'tx') bindTx();
  if (state.tab === 'add') bindAdd();
}
function renderFreshness() {
  const f = document.getElementById('freshness'), b = document.getElementById('banner');
  const k = state.data ? kpis() : {};
  const pushed = state.data && state.data.pushedAt ? new Date(state.data.pushedAt) : null;
  f.textContent = state.data
    ? `Data to ${fDay(pDate(k.as_of)).replace(/ \d{4}$/, '')} · sent ${pushed ? fDay(pushed).replace(/ \d{4}$/, '') + ', ' + pushed.toTimeString().slice(0, 5) : '–'}`
    : '';
  let msg = state.lastError, cls = 'banner err';
  if (!msg && pushed && (Date.now() - pushed) > 3 * 864e5) {
    msg = `Excel last sent figures ${Math.floor((Date.now() - pushed) / 864e5)} days ago. Open the workbook on your PC to sync and update.`;
    cls = 'banner';
  }
  if (!msg && state.outbox.length) { msg = `${state.outbox.length} entr${state.outbox.length === 1 ? 'y' : 'ies'} saved on this phone, waiting to upload.`; cls = 'banner'; }
  b.hidden = !msg; b.className = cls; b.textContent = msg || '';
}

function ring(score) {
  const s = Math.max(0, Math.min(100, n(score) || 0)), r = 30, c = 2 * Math.PI * r;
  const col = s >= 75 ? 'var(--good)' : s >= 50 ? 'var(--warn)' : 'var(--bad)';
  return `<svg class="ring" width="76" height="76" viewBox="0 0 76 76" role="img" aria-label="Health score ${s} of 100">
    <circle cx="38" cy="38" r="${r}" fill="none" stroke="#2c2944" stroke-width="8"/>
    <circle cx="38" cy="38" r="${r}" fill="none" stroke="${col}" stroke-width="8" stroke-linecap="round"
      stroke-dasharray="${(c * s / 100).toFixed(1)} ${c.toFixed(1)}" transform="rotate(-90 38 38)"/>
    <text x="38" y="43" text-anchor="middle" font-size="19" font-weight="700" fill="var(--text)">${s}</text></svg>`;
}

function monthlyChart(rows) {
  rows = rows.slice(-12);
  if (!rows.length) return '';
  const W = 340, H = 170, pl = 8, pr = 8, pt = 10, pb = 22, iw = W - pl - pr, ih = H - pt - pb;
  const max = Math.max(1, ...rows.map(r => Math.max(n(r.Income) || 0, n(r.Spending) || 0)));
  const bw = iw / rows.length, y = v => pt + ih - (Math.max(0, v) / max) * ih;
  let g = '';
  rows.forEach((r, i) => {
    const x = pl + i * bw, w = (bw - 6) / 2;
    g += `<rect x="${(x + 2).toFixed(1)}" y="${y(n(r.Income)).toFixed(1)}" width="${w.toFixed(1)}" height="${(pt + ih - y(n(r.Income))).toFixed(1)}" rx="2.5" fill="var(--teal)"/>`;
    g += `<rect x="${(x + 3 + w).toFixed(1)}" y="${y(n(r.Spending)).toFixed(1)}" width="${w.toFixed(1)}" height="${(pt + ih - y(n(r.Spending))).toFixed(1)}" rx="2.5" fill="var(--amber)"/>`;
    const d = pDate(r.Month);
    if (i % 2 === rows.length % 2 || rows.length < 7) g += `<text x="${(x + bw / 2).toFixed(1)}" y="${H - 6}" text-anchor="middle" font-size="10" fill="var(--muted)">${MON[d.getMonth()]}</text>`;
  });
  return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Income and spending, last 12 months">${g}</svg>
    <div class="legend"><span><i style="background:var(--teal)"></i>Income</span><span><i style="background:var(--amber)"></i>Spending</span></div>`;
}

function viewHome() {
  const k = kpis(), month = pDate(k.month);
  const spent = n(k.spending_month), budget = n(k.budget_month), used = budget > 0 ? spent / budget : 0;
  const alerts = tbl('AppAlerts').map(a => a.Alert).filter(a => /^[⚠ℹ]/.test(a || ''));
  const months = tbl('AppMonthly');
  const lastLeft = months.slice(-6);
  return `
  <section class="card hero">
    <div class="label">Net worth</div>
    <div class="big num">${inr(k.net_worth)}</div>
    <div class="small" style="color:#cfc8ff">Incl. investments & PF, minus loans</div>
  </section>
  <div class="grid2">
    <div class="stat"><div class="l">Money available</div><div class="v num">${inr(k.money_available)}</div><div class="s">Cash & bank, all accounts</div></div>
    <div class="stat"><div class="l">Left over · ${month ? MON[month.getMonth()] : ''}</div><div class="v num ${n(k.left_over) < 0 ? 'bad' : 'good'}">${inr(k.left_over)}</div><div class="s">${esc(k.over_budget_count)} categories over budget</div></div>
    <div class="stat"><div class="l">Loans still owed</div><div class="v num">${inr(k.loans_owed)}</div><div class="s">Across all loans</div></div>
    <div class="stat"><div class="l">Savings rate this year</div><div class="v num ${n(k.savings_rate) >= n(k.savings_target) ? 'good' : 'warn'}">${pct(k.savings_rate)}</div><div class="s">Target ${pct(k.savings_target)}</div></div>
  </div>
  <section class="card" style="margin-top:12px">
    <h2>${esc(fMon(month))} so far</h2>
    <div class="kv"><span>Income</span><span class="num good">${inr(k.income_month)}</span></div>
    <div class="kv"><span>Spending</span><span class="num">${inr(spent)}</span></div>
    <div class="kv"><span>Saved & invested</span><span class="num">${inr(k.saved_month)}</span></div>
    <div style="margin:12px 0 6px" class="row small"><span class="muted">Spending vs budget ${inr(budget)}</span><span class="num">${Math.round(used * 100)}%</span></div>
    <div class="bar ${used > 1 ? 'over' : ''}"><i style="width:${Math.min(100, used * 100).toFixed(1)}%"></i></div>
    <p class="small ${tone(k.pace_text)}" style="margin:10px 0 0">${esc(k.pace_text)}</p>
    ${k.biggest_overspend ? `<p class="small muted" style="margin:4px 0 0">${esc(k.biggest_overspend)}</p>` : ''}
  </section>
  <section class="card">
    <h2>Last 12 months</h2>
    ${monthlyChart(months)}
    <div class="row small" style="margin-top:10px;flex-wrap:wrap;gap:6px 12px">${lastLeft.map(r => `<span class="muted">${MON[pDate(r.Month).getMonth()]} <b class="num ${n(r.LeftOver) < 0 ? 'bad' : 'good'}">${inr(r.LeftOver)}</b></span>`).join('')}</div>
    <p class="small muted" style="margin:6px 0 0">Left over per month (income − spending − saved)</p>
  </section>
  <section class="card">
    <h2>Financial health</h2>
    <div class="score">${ring(k.health_score)}<div><div style="font-size:17px;font-weight:650">${esc(k.health_label)}</div><div class="small muted">${esc(k.health_score)} / 100 · ✓ = 1 pt, OK = ½ pt</div></div></div>
    <ul class="checks">${tbl('AppHealth').map(h => `<li><span>${esc(h.Check)}<br><span class="small muted">${esc(h.Value)}</span></span><span class="st small ${tone(h.Status)}">${esc(h.Status)}</span></li>`).join('')}</ul>
  </section>
  ${alerts.length ? `<section class="card"><h2>Needs attention · ${esc(k.alerts_summary)}</h2><ul class="alerts">${alerts.map(a => `<li class="${/^⚠/.test(a) ? 'warn' : ''}">${esc(a)}</li>`).join('')}</ul></section>` : ''}
  <section class="card"><h2>Salary & tax</h2>
    <div class="kv"><span>Latest salary</span><span class="num small">${esc(k.salary_line)}</span></div>
    <p class="small" style="margin:8px 0 0">${esc(k.tax_better)}</p></section>`;
}

function allTx() {
  const pend = state.outbox.map(e => ({ e, state: 'outbox' })).concat(state.pending.map(p => ({ e: p.entry, state: 'pending', sha: p.sha, id: p.id })));
  const p = pend.map(x => ({ Date: x.e.date, Type: x.e.type, Category: x.e.category, Amount: Number(x.e.amount),
    Details: (x.e.note || x.e.category) + (x.e.paidWith ? ` (${x.e.paidWith})` : '') + ` app#${x.e.id}`, Source: 'Manual entry',
    _pending: x.state, _id: x.e.id, _sha: x.sha }));
  return p.concat(tbl('AppRecent'));
}
function viewTx() {
  const f = state.tx, all = allTx();
  const years = [...new Set(all.map(t => String(t.Date).slice(0, 4)))].filter(y => /^\d{4}$/.test(y)).sort().reverse();
  if (f.year !== 'All' && !years.includes(f.year)) f.year = 'All';
  const inYear = all.filter(t => f.year === 'All' || String(t.Date).startsWith(f.year));
  const monthsHere = new Set(inYear.map(t => String(t.Date).slice(5, 7)));
  const cats = [...new Set(all.filter(t => f.type === 'All' || t.Type === f.type).map(t => t.Category).filter(Boolean))].sort();
  if (f.cat !== 'All' && !cats.includes(f.cat)) f.cat = 'All';
  const q = f.q.trim().toLowerCase();
  const list = all.filter(t => (f.type === 'All' || t.Type === f.type) &&
    (f.year === 'All' || String(t.Date).startsWith(f.year)) &&
    (f.month === 'All' || String(t.Date).slice(5, 7) === f.month) &&
    (f.cat === 'All' || t.Category === f.cat) &&
    (!q || `${t.Details} ${t.Category} ${t.Source} ${Math.round(Math.abs(n(t.Amount)))}`.toLowerCase().includes(q)));
  let tot = { Income: 0, Expenses: 0 };
  list.forEach(t => { if (t.Type in tot) tot[t.Type] += n(t.Amount) || 0; });
  let html = `<input class="search" id="tx-q" type="search" placeholder="Search merchant, category or amount" value="${esc(f.q)}" autocomplete="off">
    <div class="chips" id="tx-types">${['All', 'Expenses', 'Income', 'Savings', 'Transfer'].map(t => `<button class="chip ${f.type === t ? 'on' : ''}" data-type="${t}">${t}</button>`).join('')}</div>
    <div class="filters">
      <select class="inp sel" id="tx-year" aria-label="Year"><option value="All">All years</option>${years.map(y => `<option ${f.year === y ? 'selected' : ''}>${y}</option>`).join('')}</select>
      <select class="inp sel" id="tx-month" aria-label="Month"><option value="All">All months</option>${MON.map((m, i) => { const v = String(i + 1).padStart(2, '0');
        return `<option value="${v}" ${f.month === v ? 'selected' : ''} ${monthsHere.has(v) ? '' : 'disabled'}>${f.year === 'All' ? m + ' (every year)' : m + ' ' + f.year}</option>`; }).join('')}</select>
      <select class="inp sel" id="tx-cat" aria-label="Category"><option value="All">All categories</option>${cats.map(c => `<option ${f.cat === c ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></div>
    <div class="small muted num" style="margin:8px 2px 0">${list.length.toLocaleString('en-IN')} transactions · in ${inr(tot.Income)} · out ${inr(tot.Expenses)}</div>`;
  if (!list.length) return html + `<div class="empty">Nothing matches.</div>`;
  let lastDay = '';
  list.slice(0, f.limit).forEach(t => {
    const day = String(t.Date).slice(0, 10);
    if (day !== lastDay) { html += `<div class="day"><span>${fDay(pDate(day))}</span></div>`; lastDay = day; }
    const d = describe(t), a = n(t.Amount);
    const cls = t.Type === 'Income' ? 'good' : t.Type === 'Transfer' ? 'muted' : '';
    const sign = t.Type === 'Income' ? '+' : t.Type === 'Expenses' ? '−' : '';
    html += `<div class="tx ${t._pending ? 'pending' : ''}"><div class="ic">${TYPE_ICON[t.Type] || '•'}</div>
      <div class="mid"><div class="t1">${esc(d.title)}${t._pending ? `<span class="badge">${t._pending === 'outbox' ? 'Not uploaded' : 'Waiting for Excel'}</span>` : ''}</div>
      <div class="t2">${esc(t.Category)} · ${esc(d.src)}</div>
      <div class="t3">${esc(fullDetails(t))}<br>${esc(t.Type)} · ${esc(t.Source)} · ${fDay(pDate(t.Date))}</div>
      ${t._pending === 'pending' ? `<button class="link-btn" data-del="${esc(t._id)}" data-sha="${esc(t._sha)}">Remove</button>` : ''}
      ${t._pending === 'outbox' ? `<button class="link-btn" data-delout="${esc(t._id)}">Remove</button>` : ''}</div>
      <div class="amt num ${cls}">${sign}${inr(Math.abs(a))}</div></div>`;
  });
  if (list.length > f.limit) html += `<button class="more" id="tx-more">Show ${Math.min(200, list.length - f.limit)} more (${(list.length - f.limit).toLocaleString('en-IN')} left)</button>`;
  else html += `<p class="small muted" style="text-align:center;margin-top:14px">End of list · all ${tbl('AppRecent').length.toLocaleString('en-IN')} transactions from Excel${oldest()}</p>`;
  return html;
}
function fullDetails(t) {
  return String(t.Details || '').replace(/^bank:\s*\{[^}]*\}\s*/i, '').replace(/\s*app#[A-Za-z0-9-]+\s*$/, '') || t.Category;
}
function oldest() {
  const r = tbl('AppRecent'); const d = r.length ? pDate(r[r.length - 1].Date) : null;
  return d ? ` (since ${fMon(d)})` : '';
}
function bindTx() {
  const q = document.getElementById('tx-q');
  q.addEventListener('input', () => { state.tx.q = q.value; state.tx.limit = 100; const pos = q.selectionStart; render(); const nq = document.getElementById('tx-q'); nq.focus(); nq.setSelectionRange(pos, pos); });
  document.querySelectorAll('#tx-types .chip').forEach(b => b.onclick = () => { state.tx.type = b.dataset.type; state.tx.limit = 100; render(); });
  document.getElementById('tx-year').onchange = e => { state.tx.year = e.target.value; state.tx.limit = 100; render(); };
  document.getElementById('tx-month').onchange = e => { state.tx.month = e.target.value; state.tx.limit = 100; render(); };
  document.getElementById('tx-cat').onchange = e => { state.tx.cat = e.target.value; state.tx.limit = 100; render(); };
  const more = document.getElementById('tx-more'); if (more) more.onclick = () => { const y = window.scrollY; state.tx.limit += 200; render(); window.scrollTo(0, y); };
  document.querySelectorAll('[data-del]').forEach(b => b.onclick = async () => {
    if (!confirmInline(b)) return;
    try { await GH.deleteEntry(b.dataset.del, b.dataset.sha); state.pending = state.pending.filter(p => p.id !== b.dataset.del); LS.set('pending', state.pending); toast('Removed'); }
    catch (e) { toast(e.message || 'Couldn’t remove – try again'); }
    render();
  });
  document.querySelectorAll('.tx').forEach(el => el.onclick = e => { if (!e.target.closest('button')) el.classList.toggle('open'); });
  document.querySelectorAll('[data-delout]').forEach(b => b.onclick = () => {
    if (!confirmInline(b)) return;
    state.outbox = state.outbox.filter(e => e.id !== b.dataset.delout); LS.set('outbox', state.outbox); toast('Removed'); render();
  });
}
function confirmInline(b) {           // first tap arms, second tap confirms (no browser dialogs)
  if (b.dataset.armed) return true;
  b.dataset.armed = '1'; b.textContent = 'Tap again to remove'; setTimeout(() => { if (b.isConnected) { delete b.dataset.armed; b.textContent = 'Remove'; } }, 3000);
  return false;
}

function viewBudget() {
  const k = kpis(), rows = tbl('AppBudget').filter(r => r.Category);
  const tT = rows.reduce((s, r) => s + (n(r.Target) || 0), 0), tS = rows.reduce((s, r) => s + (n(r.Spent) || 0), 0);
  const over = rows.filter(r => n(r.Spent) > n(r.Target));
  return `<section class="card">
    <h2>${esc(fMon(pDate(k.month)))} · monthly targets</h2>
    <div class="row"><span style="font-size:26px;font-weight:700" class="num">${inr(tS)}</span><span class="muted small num">of ${inr(tT)}</span></div>
    <div class="bar ${tS > tT ? 'over' : ''}" style="margin-top:8px"><i style="width:${Math.min(100, tT ? tS / tT * 100 : 0).toFixed(1)}%"></i></div>
    <p class="small ${tS > tT ? 'bad' : 'good'}" style="margin:8px 0 0">${tS > tT ? `${inr(tS - tT)} over budget` : `${inr(tT - tS)} left`} · ${over.length} of ${rows.length} categories over</p>
    <p class="small ${tone(k.pace_text)}" style="margin:4px 0 0">${esc(k.pace_text)}</p></section>
  <section class="card"><h2>By category</h2>
    ${rows.map(r => { const t = n(r.Target), s = n(r.Spent), u = t > 0 ? s / t : (s > 0 ? 1.01 : 0);
      return `<div class="bud"><div class="row"><span>${esc(r.Category)}</span><span class="num ${s > t ? 'bad' : ''}">${inr(s)} <span class="muted small">/ ${inr(t)}</span></span></div>
      <div class="bar ${s > t ? 'over' : ''}"><i style="width:${Math.min(100, u * 100).toFixed(1)}%"></i></div>
      <div class="s ${tone(r.Status)}">${esc(r.Status)}${s <= t && t > 0 ? ` · ${inr(n(r.Left))} left` : ''}</div></div>`; }).join('')}
    <p class="small muted" style="margin:12px 0 0">Targets come from Budget Planning in Excel.</p></section>`;
}

function projChart(rows) {
  if (!rows.length) return '';
  const W = 340, H = 190, pl = 36, pr = 10, pt = 10, pb = 22, iw = W - pl - pr, ih = H - pt - pb;
  const max = Math.max(1, ...rows.map(r => Math.max(n(r.NetWorth) || 0, n(r.Needed) || 0)));
  const x = i => pl + (i / (rows.length - 1 || 1)) * iw, y = v => pt + ih - (Math.max(0, n(v) || 0) / max) * ih;
  const path = key => rows.map((r, i) => `${i ? 'L' : 'M'}${x(i).toFixed(1)},${y(r[key]).toFixed(1)}`).join('');
  let g = '';
  [0, .5, 1].forEach(f => { const yy = pt + ih - f * ih; g += `<line x1="${pl}" x2="${W - pr}" y1="${yy}" y2="${yy}" stroke="#2f2b47"/><text x="${pl - 4}" y="${yy + 3}" text-anchor="end" font-size="9.5" fill="var(--muted)">${f ? Math.round(max * f / 1e7) + ' Cr' : '0'}</text>`; });
  rows.forEach((r, i) => { if (i % 5 === 0 || i === rows.length - 1) g += `<text x="${x(i).toFixed(1)}" y="${H - 6}" text-anchor="middle" font-size="9.5" fill="var(--muted)">${r.Year}</text>`; });
  g += `<path d="${path('Needed')}" fill="none" stroke="var(--amber)" stroke-width="2" stroke-dasharray="5 4"/>`;
  g += `<path d="${path('NetWorth')}" fill="none" stroke="var(--violet)" stroke-width="2.5"/>`;
  return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Projected net worth vs amount needed to retire">${g}</svg>
  <div class="legend"><span><i style="background:var(--violet)"></i>Projected net worth</span><span><i style="background:var(--amber)"></i>Needed to retire</span></div>`;
}
function viewFuture() {
  const k = kpis(), rows = tbl('AppProjection').filter(r => r.Year !== '' && r.Year != null);
  const ageR = n(k.retire_age) || 60;
  const pick = rows.filter((r, i) => i === 0 || n(r.Age) % 5 === 0 || n(r.Age) === ageR);
  const ea = n(k.earliest_retire_age);
  return `<section class="card hero">
    <div class="label">Net worth at ${ageR}</div>
    <div class="big num">${cr(k.nw_at_60)}</div>
    <div class="small" style="color:#cfc8ff">${cr(k.nw_at_60_today)} in today’s money · ${esc(k.retire_status)}</div></section>
  <section class="card"><h2>Retirement checks</h2>
    <div class="kv"><span>Earliest retirement</span><span>${isFinite(ea) ? 'Age ' + ea : esc(k.earliest_retire_age)} <span class="small ${tone(k.earliest_retire_status)}">${esc(k.earliest_retire_status)}</span></span></div>
    <div class="kv"><span>Money lasts until</span><span>${esc(k.money_lasts)} <span class="small ${tone(k.money_lasts_status)}">${esc(k.money_lasts_status)}</span></span></div>
    <div class="kv"><span>Corpus needed at ${ageR}</span><span class="num">${cr(k.corpus_needed)}</span></div>
    <div class="kv"><span>Investing per month</span><span class="num">${inr(k.invest_actual_month)} <span class="muted small">actual · plan ${inr(k.invest_plan_month)}</span></span></div></section>
  <section class="card"><h2>At your actual SIP pace</h2>
    <div class="row"><span style="font-size:24px;font-weight:700" class="num">${cr(k.actual_pace_60)}</span><span class="small ${tone(k.actual_pace_status)}">${esc(k.actual_pace_status)}</span></div>
    <p class="small ${tone(k.actual_pace_note)}" style="margin:8px 0 0">${esc(k.actual_pace_note)}</p></section>
  <section class="card"><h2>Path to ${ageR}</h2>${projChart(rows)}
    <table class="proj" style="margin-top:10px"><tr><th>Year</th><th>Age</th><th>Net worth</th><th>Today’s ₹</th></tr>
    ${pick.map(r => `<tr><td>${r.Year}</td><td>${r.Age}</td><td class="num">${cr(r.NetWorth)}</td><td class="num">${cr(r.NetWorthToday)}</td></tr>`).join('')}</table>
    <p class="small muted" style="margin:10px 0 0">From NW Projection in Excel (car and house plans included).</p></section>`;
}

/* ---------------- More: every Excel sheet ---------------- */
const SHEET_ICON = { Home: '⌂', Insights: '📊', 'Monthly Compare': '📅', Goals: '🎯', Loans: '🏦', LIC: '🛡', 'Stocks & MF': '📈',
  'Yearly Calendar': '🗓', 'Side Income': '💼', 'Goal Planner': '🧭', 'NW Projection': '🚀', 'Savings Schemes': '🏛', 'Big Expenses': '💸',
  Trends: '〰', 'Yearly Summary': '📆', 'Budget Planning': '📝', Balances: '💰', Salary: '🧾', Payslips: '📄' };
function sheetList() { return (state.sheets && state.sheets.sheets) || []; }
function viewMore() {
  const list = sheetList(), at = state.sheets && state.sheets.exportedAt ? new Date(state.sheets.exportedAt) : null;
  return `<section class="card"><h2>App screens</h2>
      <button class="navrow" data-go="future"><span class="ni">↗</span><span class="nt">Retirement<br><span class="small muted">Net worth at 60, earliest retirement, path to 60</span></span><span class="chev">›</span></button></section>
    <section class="card"><h2>Your Excel sheets${at ? ` · sent ${fDay(at).replace(/ \d{4}$/, '')}, ${at.toTimeString().slice(0, 5)}` : ''}</h2>
    ${list.length ? list.map(s => `<button class="navrow" data-sheet="${esc(s.name)}"><span class="ni">${SHEET_ICON[s.name] || '▦'}</span>
        <span class="nt">${esc(s.name)}${s.hidden ? ' <span class="badge">hidden in Excel</span>' : ''}<br><span class="small muted">${sheetHint(s)}</span></span><span class="chev">›</span></button>`).join('')
      : `<p class="small muted" style="margin:0">No sheets yet. Update the AppLink macro in Excel (see the update notes), run <b>AppLinkNow</b>, then tap ⟳.</p>`}
    </section>
    <p class="small muted" style="text-align:center">Sheets show exactly what Excel shows. Charts aren’t included – the app draws its own.</p>`;
}
function sheetHint(s) {
  if (s.choices && s.choices.length) return esc('Pick ' + s.choices.map(c => c.label.toLowerCase()).join(', ') + ': ' + s.choices[0].options.map(o => o.v).join(', '));
  const first = (s.rows || []).slice(0, 3).map(r => r.c.find(x => x)).filter(Boolean);
  const t = first.find((x, i) => i > 0 && x.length > 8) || first[1] || first[0] || '';
  return esc(t.length > 60 ? t.slice(0, 58) + '…' : t);
}
const NUMISH = /^[-−+▲▼(]?\s?₹?\s?-?[\d,]+(\.\d+)?\)?\s?(%|×|Cr|L|mo|yrs)?$/;
function viewSheet() {
  const sh = sheetList().find(s => s.name === state.sheet);
  if (!sh) { state.tab = 'more'; return viewMore(); }
  const names = new Set(sheetList().map(s => s.name));
  const q = state.sheetQ.trim().toLowerCase();
  const chs = sh.choices || [];
  let picked = state.choice[sh.name], src = sh.rows || [];
  if (picked) {
    const ch = chs.find(x => x.label === picked.label), o = ch && ch.options.find(x => x.v === picked.v);
    if (o) src = o.rows; else picked = null;
  }
  const chips = chs.map(ch => {
    const cur = picked && picked.label === ch.label ? picked.v : ch.current;
    return `<div class="small muted" style="margin:2px 2px 0">${esc(ch.label)}</div><div class="chips">${ch.options.map(o =>
      `<button class="chip ${o.v === cur ? 'on' : ''}" data-chl="${esc(ch.label)}" data-chv="${esc(o.v)}">${esc(o.v)}</button>`).join('')}</div>`;
  }).join('');
  const rows = src.filter(r => !q || r.c.join(' ').toLowerCase().includes(q));
  const width = Math.max(1, ...src.map(r => r.c.length));
  const cell = (t, bold) => {
    if (!t) return '<td></td>';
    const m = t.match(/^→\s*(.+?)(\s+for details)?$/);
    if (m && names.has(m[1].trim())) return `<td><button class="xl-link" data-sheet="${esc(m[1].trim())}">→ ${esc(m[1].trim())}</button></td>`;
    const cls = [NUMISH.test(t) ? 'n' : '', tone(t), bold ? 'b' : '', t.length > 38 ? 'long' : ''].filter(Boolean).join(' ');
    return `<td${cls ? ` class="${cls}"` : ''}>${esc(t)}</td>`;
  };
  const top = `<div class="row" style="gap:8px;margin-bottom:10px"><button class="secondary back" data-go="more">‹ Sheets</button>
      <input class="search" id="sheet-q" type="search" placeholder="Find in ${esc(sh.name)}" value="${esc(state.sheetQ)}" autocomplete="off" style="flex:1"></div>
    <div class="seg seg2" id="sheet-mode"><button data-m="cards" class="${state.sheetMode === 'cards' ? 'on' : ''}">Cards</button><button data-m="table" class="${state.sheetMode === 'table' ? 'on' : ''}">Table</button></div>`;
  if (state.sheetMode === 'cards') return top + chips + sheetCards(sh, rows, names);
  return top + chips + `${sh.hidden ? '<p class="small muted" style="margin:0 0 8px">This sheet is hidden in Excel.</p>' : ''}
    <div class="xl-wrap"><table class="xl"><tbody>${rows.map(r => {
      const b = new Set(r.b || []);
      const c = r.c.concat(Array(width - r.c.length).fill(''));
      let run = 1; while (run < width && !c[run]) run++;           // empty cells right of the first one
      if (c[0] && c[0].length > 24 && run > 1) {                   // titles & notes stretch across them
        const first = `<td colspan="${run}" class="span ${b.has(0) ? 'b' : ''} ${tone(c[0])}"><div>${esc(c[0])}</div></td>`;
        return `<tr>${first}${c.slice(run).map((t, i) => cell(t, b.has(i + run))).join('')}</tr>`;
      }
      return `<tr>${c.map((t, i) => cell(t, b.has(i))).join('')}</tr>`; }).join('')}</tbody></table></div>
    ${!rows.length ? '<div class="empty">Nothing matches.</div>' : ''}
    <p class="small muted" style="margin:10px 2px 0">Swipe sideways for more columns. Read-only – change things in Excel.</p>`;
}
/* Turn a sheet into app-style cards: bold one-cell rows start a card, bold multi-cell rows are column
   headers, label + value rows become tidy rows, table rows become items, notes become small text. */
const PRIMARY_HDR = /(used|now|owed|current|value|total|balance|net|amount|spent|left|corpus|worth)/i;
const isNum = t => NUMISH.test(String(t || '').trim());
const isStatus = t => /^[✓✅⚠ℹ🟡🔴]/.test(String(t || '').trim());
function barHtml(t) {
  const m = String(t).match(/^[█░]+\s+(-?\d+)%/);
  if (!m) return null;
  const p = Number(m[1]);
  return `<div class="row small" style="gap:8px;width:100%"><div class="bar ${p > 100 ? 'over' : ''}" style="flex:1"><i style="width:${Math.min(100, Math.max(0, p))}%"></i></div><span class="num">${p}%</span></div>`;
}
function valHtml(t, names) {
  const m = String(t).match(/^→\s*(.+?)(\s+for details)?$/);
  if (m && names.has(m[1].trim())) return `<button class="xl-link" data-sheet="${esc(m[1].trim())}">→ ${esc(m[1].trim())}</button>`;
  return `<span class="${tone(t)}">${esc(t)}</span>`;
}
function sheetCards(sh, rows, names) {
  const out = [];
  let card = null, headers = null, lastTiles = null;
  const open = (title, bold) => { card = { title, items: [] }; out.push(card); headers = null; lastTiles = null; };
  const add = html => { if (!card) open(''); card.items.push(html); };
  rows.forEach((r, ri) => {
    const c = r.c, b = new Set(r.b || []);
    const f = c.map((t, i) => t ? i : -1).filter(i => i >= 0);
    if (!f.length) return;
    const first = c[f[0]], allBold = f.every(i => b.has(i));
    // continuation line under stat tiles (e.g. "Cash & bank, all accounts")
    if (lastTiles && lastTiles.n < 2 && (!allBold || lastTiles.n === 0) && f.every(i => lastTiles.pos.includes(i)) &&
        !f.some(i => isNum(c[i]) || /^→/.test(c[i]))) {
      const subs = lastTiles.el.sub || (lastTiles.el.sub = lastTiles.pos.map(() => []));
      f.forEach(i => subs[lastTiles.pos.indexOf(i)].push(c[i])); lastTiles.n++; return;     // lines under each figure
    }
    lastTiles = null;
    if (ri === 0 && b.has(f[0])) {                                  // sheet title
      open(first); const rest = f.slice(1).map(i => c[i]);
      if (rest.length) add(`<p class="small muted sc-note">${esc(rest.join(' '))}</p>`); return;
    }
    if (f.length === 1) {
      if (b.has(f[0]) && first.length < 70 && !isNum(first)) { open(first); return; }
      if (first.length > 30 || !b.has(f[0])) { add(`<p class="small muted sc-note ${tone(first)}">${valHtml(first, names)}</p>`); return; }
    }
    const nonFirstNum = f.slice(1).filter(i => isNum(c[i])).length;
    if (allBold && f.length >= 2 && nonFirstNum === 0 && !isNum(first)) {   // header row
      if (!(card && card.items.length === 0 && card.title)) open(first);
      headers = {}; f.forEach(i => { headers[i] = c[i]; }); return;
    }
    if (headers && isNum(first) && f.length >= 2) {                 // a row of figures under headers -> tiles
      const tiles = { tiles: f.map(i => [headers[i] || '', c[i], b.has(i)]), sub: null };
      add(tiles); lastTiles = { el: tiles, pos: f, n: 0 }; return;
    }
    if (f.length === 2 && !headers) {                               // label : value
      add(`<div class="kv ${b.has(f[0]) ? 'kvb' : ''}"><span>${valHtml(first, names)}</span><span class="num sc-v">${barHtml(c[f[1]]) || valHtml(c[f[1]], names)}</span></div>`); return;
    }
    // item with several values
    let vals = f.slice(1).map(i => ({ h: headers ? (headers[i] || '') : '', v: c[i], i }));
    if (vals.some(x => barHtml(x.v))) vals = vals.filter(x => x.h || !isNum(x.v));   // drop helper numbers next to progress bars
    const nums = vals.filter(x => isNum(x.v));
    const prim = nums.find(x => /\b(used|now|owed|current|balance|spent)\b/i.test(x.h) && !/%/.test(x.h))
      || nums.find(x => PRIMARY_HDR.test(x.h) && !/%/.test(x.h)) || nums.find(x => !/%/.test(x.h)) || nums[0] || null;
    const rest = vals.filter(x => x !== prim);
    const bars = rest.filter(x => barHtml(x.v)), statuses = rest.filter(x => !barHtml(x.v) && isStatus(x.v)), others = rest.filter(x => !barHtml(x.v) && !isStatus(x.v));
    add(`<div class="sc-item ${b.has(f[0]) ? 'kvb' : ''}">
      <div class="row"><span class="sc-l">${valHtml(first, names)}</span>${prim ? `<span class="num sc-v">${esc(prim.v)}</span>` : ''}</div>
      ${bars.map(x => barHtml(x.v)).join('')}
      ${others.length ? `<div class="small muted sc-sub">${others.map(x => x.h ? `${esc(x.h)} <b class="num">${esc(x.v)}</b>` : `<span class="${tone(x.v)}">${esc(x.v)}</span>`).join(' · ')}</div>` : ''}
      ${statuses.map(x => `<div class="small ${tone(x.v)}">${esc(x.v)}</div>`).join('')}
    </div>`);
  });
  if (!out.length) return '<div class="empty">Nothing matches.</div>';
  return out.filter(cd => cd.items.length || cd.title).map(cd => `<section class="card">${cd.title ? `<h3 class="sc-h">${esc(cd.title)}</h3>` : ''}
    ${cd.items.map(it => typeof it === 'string' ? it : `<div class="grid2 sc-tiles">${it.tiles.map(([h, v, bb], k) =>
      `<div class="stat"><div class="l">${esc(h)}</div><div class="v num ${tone(v)} ${String(v).startsWith('-') ? 'bad' : ''}">${esc(v)}</div>${it.sub && it.sub[k] ? it.sub[k].map(x => `<div class="s">${esc(x)}</div>`).join('') : ''}</div>`).join('')}</div>`).join('')}</section>`).join('')
    + `<p class="small muted" style="text-align:center">Read-only – change things in Excel. Switch to Table to see the exact grid.</p>`;
}
function bindSheets() {
  document.querySelectorAll('[data-chl]').forEach(b => b.onclick = () => { state.choice[state.sheet] = { label: b.dataset.chl, v: b.dataset.chv }; render(); });
  document.querySelectorAll('#sheet-mode button').forEach(b => b.onclick = () => { state.sheetMode = b.dataset.m; LS.set('sheetMode', state.sheetMode); render(); });
  document.querySelectorAll('[data-sheet]').forEach(b => b.onclick = () => { state.sheet = b.dataset.sheet; state.sheetQ = ''; state.tab = 'sheet'; window.scrollTo(0, 0); render(); });
  document.querySelectorAll('[data-go]').forEach(b => b.onclick = () => { state.tab = b.dataset.go; window.scrollTo(0, 0); render(); });
  const q = document.getElementById('sheet-q');
  if (q) q.addEventListener('input', () => { state.sheetQ = q.value; const pos = q.selectionStart; render(); const nq = document.getElementById('sheet-q'); nq.focus(); nq.setSelectionRange(pos, pos); });
}

/* ---------------- add entry ---------------- */
function catsFor(type) {
  return tbl('AppLists').map(r => r[type]).filter(c => c && !/^Enter /i.test(c));
}
function newAdd() { return { type: 'Expenses', amount: '', category: '', date: isoToday(), paidWith: 'Cash', note: '' }; }
function viewAdd() {
  const a = state.add || (state.add = newAdd());
  const cats = catsFor(a.type);
  if (!cats.includes(a.category)) a.category = a.type === 'Expenses' && cats.includes('Cash') ? 'Cash' : (cats[0] || '');
  return `<section class="card">
    <div class="seg" id="add-type">${['Expenses', 'Income', 'Savings', 'Transfer'].map(t => `<button data-t="${t}" class="${a.type === t ? 'on' : ''}">${t === 'Expenses' ? 'Spent' : t}</button>`).join('')}</div>
    <label class="f" for="add-amt">Amount (₹)</label>
    <input class="inp amount num" id="add-amt" inputmode="decimal" placeholder="0" value="${esc(a.amount)}" autocomplete="off">
    <label class="f" for="add-cat">Category</label>
    <select class="inp" id="add-cat">${cats.map(c => `<option ${c === a.category ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select>
    <label class="f" for="add-date">Date</label>
    <input class="inp" id="add-date" type="date" value="${esc(a.date)}" max="${isoToday()}">
    <label class="f">Paid with</label>
    <div class="pay" id="add-pay">${['Cash', 'UPI', 'Card', 'Other'].map(p => `<button class="chip ${a.paidWith === p ? 'on' : ''}" data-p="${p}">${p}</button>`).join('')}</div>
    <label class="f" for="add-note">Note</label>
    <input class="inp" id="add-note" maxlength="80" placeholder="e.g. Vegetables at market" value="${esc(a.note)}">
    <button class="primary" id="add-save">Save entry</button>
    <p class="note">Use this for cash, or anything that won’t show up in a bank statement. Bank transactions arrive through your Excel sync, so adding them here too would count them twice.</p>
    <p class="note">Your entry waits on GitHub until Excel next opens or saves. Then it goes into Tracking as a “Manual entry” row, and totals update after the next send.</p>
  </section>`;
}
function bindAdd() {
  const a = state.add;
  const sync1 = () => { a.amount = document.getElementById('add-amt').value; a.category = document.getElementById('add-cat').value;
    a.date = document.getElementById('add-date').value; a.note = document.getElementById('add-note').value; };
  document.querySelectorAll('#add-type button').forEach(b => b.onclick = () => { sync1(); a.type = b.dataset.t; a.category = ''; render(); });
  document.querySelectorAll('#add-pay .chip').forEach(b => b.onclick = () => { sync1(); a.paidWith = b.dataset.p; render(); });
  document.getElementById('add-save').onclick = async () => {
    sync1();
    const amt = Number(String(a.amount).replace(/[,₹\s]/g, ''));
    if (!(amt > 0)) { toast('Enter an amount'); document.getElementById('add-amt').focus(); return; }
    if (!a.category) { toast('Pick a category'); return; }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(a.date)) { toast('Pick a date'); return; }
    const now = new Date(), stamp = now.toISOString().replace(/[-:]/g, '').slice(0, 15);
    const entry = { id: `${stamp}-${Math.random().toString(36).slice(2, 7)}`, date: a.date, type: a.type, category: a.category,
      amount: Math.round(amt * 100) / 100, paidWith: a.paidWith, note: a.note.trim().slice(0, 80), createdAt: now.toISOString() };
    const btn = document.getElementById('add-save'); btn.disabled = true; btn.textContent = 'Saving…';
    try {
      const sha = await GH.putEntry(entry);
      state.pending.unshift({ id: entry.id, sha, entry }); LS.set('pending', state.pending);
      toast(`Saved ${inr(entry.amount)} · Excel adds it next time it opens`);
    } catch (e) {
      state.outbox.push(entry); LS.set('outbox', state.outbox);
      toast(e instanceof AppError && !/offline/i.test(e.message) && navigator.onLine !== false ? e.message + ' Kept on this phone.' : 'Saved on this phone · uploads when you’re online');
    }
    state.add = newAdd(); state.tab = 'tx'; state.tx = { q: '', type: 'All', year: 'All', month: 'All', cat: 'All', limit: 100 }; render();
  };
}

/* ---------------- settings ---------------- */
function viewSettings() {
  const c = state.cfg || { owner: '', repo: 'mymoney-data', token: '', branch: 'main' };
  return `<section class="card">
    <h2>${state.cfg ? 'GitHub link' : 'Connect to your data'}</h2>
    <p class="small muted" style="margin:0 0 4px">Same details you put in the AppLink macro. The token is kept only on this phone.</p>
    <label class="f" for="s-owner">GitHub username</label><input class="inp" id="s-owner" autocapitalize="off" autocorrect="off" spellcheck="false" value="${esc(c.owner)}">
    <label class="f" for="s-repo">Private data repo</label><input class="inp" id="s-repo" autocapitalize="off" autocorrect="off" spellcheck="false" value="${esc(c.repo)}">
    <label class="f" for="s-token">Token</label><input class="inp" id="s-token" type="password" autocomplete="off" placeholder="github_pat_…" value="${esc(c.token)}">
    <label class="f" for="s-branch">Branch</label><input class="inp" id="s-branch" autocapitalize="off" value="${esc(c.branch || 'main')}">
    <button class="primary" id="s-save">Save & load data</button>
    <p class="small" id="s-msg" style="margin:10px 0 0"></p>
    ${state.cfg ? `<button class="secondary" id="s-back">Back</button><button class="secondary" id="s-forget" style="color:var(--bad)">Forget this phone</button>` : ''}
  </section>
  <p class="small muted" style="text-align:center">MyMoney · reads figures your Excel workbook sends to GitHub.</p>`;
}
function bindSettings() {
  const msg = document.getElementById('s-msg');
  document.getElementById('s-save').onclick = async () => {
    const cfg = { owner: document.getElementById('s-owner').value.trim(), repo: document.getElementById('s-repo').value.trim(),
      token: document.getElementById('s-token').value.trim(), branch: document.getElementById('s-branch').value.trim() || 'main' };
    if (!cfg.owner || !cfg.repo || !cfg.token) { msg.className = 'small bad'; msg.textContent = 'Fill in username, repo and token.'; return; }
    const prev = state.cfg; state.cfg = cfg; msg.className = 'small muted'; msg.textContent = 'Checking…';
    try {
      const repo = await GH.testRepo();
      if (!repo.private) { msg.className = 'small bad'; msg.textContent = 'That repo is PUBLIC – anyone could read your figures. Make it private on GitHub first.'; state.cfg = prev; return; }
      LS.set('cfg', cfg); state.tab = 'home'; state.lastError = null; render(); sync({ toast: true });
    } catch (e) { state.cfg = prev; msg.className = 'small bad'; msg.textContent = e.message || 'Couldn’t reach GitHub.'; }
  };
  const back = document.getElementById('s-back'); if (back) back.onclick = () => { state.tab = 'home'; render(); };
  const forget = document.getElementById('s-forget');
  if (forget) forget.onclick = () => { if (!confirmInline(forget)) { forget.textContent = 'Tap again to forget this phone'; return; }
    ['cfg', 'data', 'fetchedAt', 'pending', 'outbox', 'sheets'].forEach(LS.del); location.reload(); };
}

/* ---------------- shell ---------------- */
let toastTimer;
function toast(t) { const el = document.getElementById('toast'); el.textContent = t; el.hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => el.hidden = true, 3200); }
document.querySelectorAll('.tab').forEach(b => b.onclick = () => { state.tab = b.dataset.tab; window.scrollTo(0, 0); render(); });
document.getElementById('btn-sync').onclick = () => sync({ toast: true });
document.getElementById('btn-settings').onclick = () => { state.tab = 'settings'; render(); };
window.addEventListener('online', () => sync());
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible' && state.cfg && (!state.fetchedAt || Date.now() - new Date(state.fetchedAt) > 5 * 60e3)) sync();
});
if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
render();
if (state.cfg) sync();
