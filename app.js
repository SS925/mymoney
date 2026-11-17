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
  sheets: LS.get('sheets', null),
  plan: LS.get('plan', null),          // net-pay plan {savings, limits:{cat:amt}}
  budgetMode: LS.get('budgetMode', 'plan'), planEdit: null,   // every sheet as Excel shows it
  sheet: null, sheetQ: '', choice: {},   // choice[sheet] = {label, v}
  sheetMode: ['fold', 'tiles', 'table'].includes(LS.get('sheetMode2', 'fold')) ? LS.get('sheetMode2', 'fold') : 'fold', fold: {},
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
  async sha(path) {
    const r = await this.call('GET', path);
    if (r.status === 404) return null;
    if (!r.ok) throw httpError(r.status);
    const j = await r.json(); return j.sha;
  },
  async putJson(path, obj, message) {
    const sha = await this.sha(path);
    const r = await this.call('PUT', path, { body: { message, branch: state.cfg.branch || 'main', content: b64(JSON.stringify(obj, null, 1)), ...(sha ? { sha } : {}) } });
    if (!r.ok) throw httpError(r.status);
  },
  async remove(path, message) {
    const sha = await this.sha(path); if (!sha) return;
    const r = await this.call('DELETE', path, { body: { message, branch: state.cfg.branch || 'main', sha } });
    if (!r.ok && r.status !== 404) throw httpError(r.status);
  },
  async exists(path) { const r = await this.call('GET', path); return r.ok; },
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
const MONF = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
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
    try { const r = await GH.call('GET', 'plan.json', { raw: true }); if (r.ok) { state.plan = await r.json(); LS.set('plan', state.plan); } } catch (e) { /* plan is optional */ }
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
  const t = document.getElementById('screen-title'), av = document.getElementById('avatar');
  if (state.tab === 'home' && state.data) {
    const nm = userName();
    t.innerHTML = `<span class="hi">Welcome,</span>${esc(nm)}`;
    av.hidden = false; av.textContent = nm.split(/\s+/).map(w => w[0]).join('').slice(0, 2).toUpperCase();
  } else { t.textContent = state.tab === 'sheet' ? state.sheet : TITLES[state.tab]; av.hidden = true; }
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
  if (state.tab === 'budget') bindBudget();
  const ha = document.querySelector('[data-homeall]'); if (ha) ha.onclick = () => { state.homeAlertsAll = !state.homeAlertsAll; render(); };
  const hh = document.querySelector('[data-health]'); if (hh) hh.onclick = () => { state.healthOpen = !state.healthOpen; render(); };
  document.querySelectorAll('[data-gotab]').forEach(b => b.onclick = () => { state.tab = b.dataset.gotab; window.scrollTo(0, 0); render(); });
}
function userName() {
  try { const me = JSON.parse(localStorage.getItem('crm_me_v1') || 'null'); if (me && me.name) return me.name; } catch (e) {}
  return (state.cfg && state.cfg.owner) || 'MyMoney';
}
function renderFreshness() {
  const f = document.getElementById('freshness'), b = document.getElementById('banner');
  const k = state.data ? kpis() : {};
  const pushed = state.data && state.data.pushedAt ? new Date(state.data.pushedAt) : null;
  f.textContent = state.data
    ? `Updated ${pushed ? fDay(pushed).replace(/ \d{4}$/, '') + ', ' + pushed.toTimeString().slice(0, 5) : '–'}`
    : '';
  let msg = state.lastError, cls = 'banner err';
  if (!msg && pushed && (Date.now() - pushed) > 3 * 864e5) {
    msg = `Figures are ${Math.floor((Date.now() - pushed) / 864e5)} days old. Open the workbook on your PC to update.`;
    cls = 'banner';
  }
  if (!msg && state.outbox.length) { msg = `${state.outbox.length} entr${state.outbox.length === 1 ? 'y' : 'ies'} saved on this phone, waiting to upload.`; cls = 'banner'; }
  b.hidden = !msg; b.className = cls; b.textContent = msg || '';
}

function ring(score) {
  const s = Math.max(0, Math.min(100, n(score) || 0)), r = 30, c = 2 * Math.PI * r;
  const col = s >= 75 ? 'var(--good)' : s >= 50 ? 'var(--warn)' : 'var(--bad)';
  return `<svg class="ring" width="48" height="48" viewBox="0 0 76 76" role="img" aria-label="Health score ${s} of 100">
    <circle cx="38" cy="38" r="${r}" fill="none" stroke="#2b373b" stroke-width="9"/>
    <circle cx="38" cy="38" r="${r}" fill="none" stroke="${col}" stroke-width="9" stroke-linecap="round"
      stroke-dasharray="${(c * s / 100).toFixed(1)} ${c.toFixed(1)}" transform="rotate(-90 38 38)"/>
    <text x="38" y="46" text-anchor="middle" font-size="24" font-weight="700" fill="var(--text)">${s}</text></svg>`;
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
    g += `<rect x="${(x + 2).toFixed(1)}" y="${y(n(r.Income)).toFixed(1)}" width="${w.toFixed(1)}" height="${(pt + ih - y(n(r.Income))).toFixed(1)}" rx="2.5" fill="var(--good)" opacity=".85"/>`;
    g += `<rect x="${(x + 3 + w).toFixed(1)}" y="${y(n(r.Spending)).toFixed(1)}" width="${w.toFixed(1)}" height="${(pt + ih - y(n(r.Spending))).toFixed(1)}" rx="2.5" fill="var(--amber)"/>`;
    const d = pDate(r.Month);
    if (i % 2 === rows.length % 2 || rows.length < 7) g += `<text x="${(x + bw / 2).toFixed(1)}" y="${H - 6}" text-anchor="middle" font-size="10" fill="var(--muted)">${MON[d.getMonth()]}</text>`;
  });
  return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Income and spending, last 12 months">${g}</svg>
    <div class="legend"><span><i style="background:var(--good)"></i>Income</span><span><i style="background:var(--amber)"></i>Spending</span></div>`;
}

/* ---------------- net-pay limit ---------------- */
const SUGGEST = { 'Housing': 30000, 'Loans/EMI': 15000, 'Credit Card Bills': 24000, 'UPI – Other': 9000, 'Family & Transfers': 8000,
  'Insurance': 4000, 'Subscriptions': 3000, 'Groceries': 2500, 'Food Online': 1000, 'Paytm': 1000, 'Utilities': 600, 'Medical': 1000,
  'Shopping': 1000, 'Travel': 500, 'Donation': 500, 'Amazon': 300, 'Bank Charges': 0, 'Cash': 0 };
const FIXED = new Set(['Housing', 'Loans/EMI', 'Insurance']);
function netPay() {
  const m = String(kpis().salary_line || '').match(/\/\s*₹\s*([\d,]+)/);
  return m ? Number(m[1].replace(/,/g, '')) : NaN;
}
function budgetCats() { return tbl('AppBudget').map(r => r.Category).filter(Boolean); }
function suggestedPlan() {
  const net = netPay(), savings = 15000, avail = Math.max(0, net - savings);
  const limits = {}; budgetCats().forEach(c => { limits[c] = SUGGEST[c] != null ? SUGGEST[c] : 0; });
  const total = Object.values(limits).reduce((s, v) => s + v, 0);
  if (isFinite(avail) && total > avail) {                    // shrink flexible categories to fit a smaller net pay
    const fixed = Object.entries(limits).filter(([c]) => FIXED.has(c)).reduce((s, [, v]) => s + v, 0);
    const f = Math.max(0, (avail - fixed) / (total - fixed));
    Object.keys(limits).forEach(c => { if (!FIXED.has(c)) limits[c] = Math.floor(limits[c] * f / 100) * 100; });
  }
  return { savings, limits, suggested: true };
}
function activePlan() { return state.plan || suggestedPlan(); }
function planTotal(p) { return Object.values(p.limits || {}).reduce((s, v) => s + (Number(v) || 0), 0); }
function monthProgress() {                                   // how far through the data month we are
  const k = kpis(), m = pDate(k.month), asOf = pDate(k.as_of);
  if (!m) return { done: 1, left: 0, dim: 30 };
  const dim = new Date(m.getFullYear(), m.getMonth() + 1, 0).getDate();
  const today = new Date(), sameMonth = today.getFullYear() === m.getFullYear() && today.getMonth() === m.getMonth();
  const day = sameMonth ? today.getDate() : (asOf && asOf.getMonth() === m.getMonth() ? asOf.getDate() : dim);
  return { done: Math.min(1, day / dim), left: sameMonth ? dim - day + 1 : 0, dim, day, complete: !sameMonth && day >= dim };
}
function netPayCard(compact) {
  const k = kpis(), net = netPay(); if (!isFinite(net)) return '';
  const p = activePlan(), limit = net - (Number(p.savings) || 0), spent = n(k.spending_month) || 0;
  const used = limit > 0 ? spent / limit : 0, left = limit - spent, mp = monthProgress(), m = pDate(k.month);
  const pace = !mp.complete && mp.done > 0 ? spent / mp.done : spent;
  const head = left < 0 ? `${inr(-left)}` : `${inr(left)}`;
  let sub = `<b>${inr(spent)}</b> spent of your <b>${inr(limit)}</b> limit`;
  if (left >= 0 && mp.left > 0) sub += `, about <b>${inr(left / mp.left)}</b> a day for ${mp.left} days`;
  if (!mp.complete && mp.done < 1 && pace > limit && left >= 0) sub += `. At this pace you’ll reach ${inr(pace)}`;
  return `<section class="hero">
    <div class="hero-top"><span>${m ? MONF[m.getMonth()] + ' ' + m.getFullYear() : ''}</span><span>Net pay ${inr(net)}</span></div>
    <p class="hero-k">${left < 0 ? 'Over your spending limit' : 'Left to spend this month'}</p>
    <h2 class="hero-v num">${head}</h2>
    <div class="meter ${used > 1 ? 'over' : ''}"><i style="width:${Math.min(100, used * 100).toFixed(1)}%"></i>${used > 1 ? `<span class="tick" style="left:${(100 / used).toFixed(1)}%"></span>` : ''}</div>
    <p class="hero-s">${sub}.</p>
  </section>`;
}

const IC = {
  add: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14M5 12h14"/></svg>',
  receipt: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3h12v18l-3-2-3 2-3-2-3 2z"/><path d="M9 8h6M9 12h6"/></svg>',
  grid: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="7.5" cy="7.5" r="3"/><circle cx="16.5" cy="7.5" r="3"/><circle cx="7.5" cy="16.5" r="3"/><circle cx="16.5" cy="16.5" r="3"/></svg>',
};
function viewHome() {
  const k = kpis();
  const alerts = tbl('AppAlerts').map(x => String(x.Alert || '')).filter(x => /^[⚠ℹ]/.test(x));
  const shown = state.homeAlertsAll ? alerts : alerts.slice(0, 3);
  const months = tbl('AppMonthly');
  const recent = allTx().slice(0, 4);
  return `
  ${netPayCard()}
  <div class="actions">
    <button class="act" data-gotab="add" aria-label="Add entry">${IC.add}</button>
    <a class="act" href="receipts.html" aria-label="New receipt">${IC.receipt}</a>
    <button class="act wide" data-gotab="budget">${IC.grid}<span>Net-pay plan</span></button>
  </div>
  <section class="sec box">
    <div class="sec-h"><h3>Transactions</h3><button data-gotab="tx">View all</button></div>
    ${recent.map(t => txRow(t, true)).join('')}
  </section>
  <section class="sec"><div class="figs">
    <div class="fig"><span>Net worth</span><b class="num">${inr(k.net_worth)}</b></div>
    <div class="fig"><span>Money available</span><b class="num">${inr(k.money_available)}</b></div>
    <div class="fig"><span>Loans owed</span><b class="num">${inr(k.loans_owed)}</b></div>
    <div class="fig"><span>Savings rate</span><b class="num ${n(k.savings_rate) >= n(k.savings_target) ? 'good' : 'warn'}">${pct(k.savings_rate)}</b><em>target ${pct(k.savings_target)}</em></div>
  </div></section>
  ${alerts.length ? `<section class="sec box"><div class="sec-h"><h3>Needs attention</h3><span class="small muted">${alerts.length}</span></div>
    ${shown.map(x => `<div class="alert-row"><span class="dot ${/^⚠/.test(x) ? 'warn' : 'info'}"></span><span>${esc(x.replace(/^[⚠ℹ]\s*/, ''))}</span></div>`).join('')}
    ${alerts.length > 3 ? `<button class="more-btn" data-homeall>${state.homeAlertsAll ? 'Show fewer' : `Show all ${alerts.length}`}</button>` : ''}</section>` : ''}
  <section class="sec box"><div class="sec-h"><h3>Last 12 months</h3></div>
    ${monthlyChart(months)}</section>
  <section class="sec">
    <button class="health-row" data-health>${ring(k.health_score)}<span class="ht"><b>${esc(String(k.health_label || '').replace(/^[^A-Za-z]+/, ''))}</b><span>Financial health, ${esc(k.health_score)} out of 100</span></span><span class="chev">${state.healthOpen ? '⌃' : '›'}</span></button>
    ${state.healthOpen ? `<ul class="checks">${tbl('AppHealth').map(h => `<li><span>${esc(h.Check)}<br><span class="small muted">${esc(h.Value)}</span></span><span class="st small ${tone(h.Status)}">${esc(String(h.Status).replace(/\s*–\s*/, ', '))}</span></li>`).join('')}</ul>` : ''}
  </section>`;
}

function allTx() {
  const pend = state.outbox.map(e => ({ e, state: 'outbox' })).concat(state.pending.map(p => ({ e: p.entry, state: 'pending', sha: p.sha, id: p.id })));
  const p = pend.map(x => ({ Date: x.e.date, Type: x.e.type, Category: x.e.category, Amount: Number(x.e.amount),
    Details: (x.e.note || x.e.category) + (x.e.paidWith ? ` (${x.e.paidWith})` : '') + ` app#${x.e.id}`, Source: 'Manual entry',
    _pending: x.state, _id: x.e.id, _sha: x.sha }));
  return p.concat(tbl('AppRecent'));
}
function txRow(t, compact) {
  const d = describe(t), a = n(t.Amount);
  const cls = t.Type === 'Income' ? 'good' : t.Type === 'Transfer' ? 'muted' : '';
  const sign = t.Type === 'Income' ? '+' : t.Type === 'Expenses' ? '−' : '';
  const tile = { Income: 'in', Savings: 'sv', Transfer: 'tr' }[t.Type] || '';
  const letter = (String(d.title).match(/[A-Za-z0-9]/) || ['•'])[0].toUpperCase();
  return `<div class="tx ${t._pending ? 'pending' : ''}"><div class="ic ${tile}">${letter}</div>
    <div class="mid"><div class="t1">${esc(d.title)}${t._pending ? `<span class="badge">${t._pending === 'outbox' ? 'Not uploaded' : 'Waiting for Excel'}</span>` : ''}</div>
    <div class="t2">${compact ? fDay(pDate(t.Date)).replace(/ \d{4}$/, '') + ', ' : ''}${esc(t.Category)}</div>
    <div class="t3">${esc(fullDetails(t))}<br>${esc(t.Type)}, ${esc(d.src)}, ${fDay(pDate(t.Date))}</div>
    ${t._pending === 'pending' && !compact ? `<button class="link-btn" data-del="${esc(t._id)}" data-sha="${esc(t._sha)}">Remove</button>` : ''}
    ${t._pending === 'outbox' && !compact ? `<button class="link-btn" data-delout="${esc(t._id)}">Remove</button>` : ''}</div>
    <div class="amt num ${cls}">${sign}${inr(Math.abs(a))}</div></div>`;
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
      <select class="inp sel" id="tx-year" aria-label="Year"><option value="All">Year</option>${years.map(y => `<option ${f.year === y ? 'selected' : ''}>${y}</option>`).join('')}</select>
      <select class="inp sel" id="tx-month" aria-label="Month"><option value="All">Month</option>${MON.map((m, i) => { const v = String(i + 1).padStart(2, '0');
        return `<option value="${v}" ${f.month === v ? 'selected' : ''} ${monthsHere.has(v) ? '' : 'disabled'}>${f.year === 'All' ? m + ', any year' : m + ' ' + f.year}</option>`; }).join('')}</select>
      <select class="inp sel" id="tx-cat" aria-label="Category"><option value="All">Category</option>${cats.map(c => `<option ${f.cat === c ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></div>
    <div class="small muted num" style="margin:8px 2px 0">${list.length.toLocaleString('en-IN')} transactions, ${inr(tot.Income)} in, ${inr(tot.Expenses)} out</div>`;
  if (!list.length) return html + `<div class="empty">Nothing matches.</div>`;
  let lastDay = '', open = false;
  list.slice(0, f.limit).forEach(t => {
    const day = String(t.Date).slice(0, 10);
    if (day !== lastDay) { html += `${open ? '</div>' : ''}<div class="day"><span>${fDay(pDate(day))}</span></div><div class="txcard">`; lastDay = day; open = true; }
    html += txRow(t);
  });
  if (open) html += '</div>';
  if (list.length > f.limit) html += `<button class="more" id="tx-more">Show ${Math.min(200, list.length - f.limit)} more (${(list.length - f.limit).toLocaleString('en-IN')} left)</button>`;
  else html += `<p class="small muted" style="text-align:center;margin-top:14px">That’s all ${tbl('AppRecent').length.toLocaleString('en-IN')} transactions from Excel${oldest()}</p>`;
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

function viewBudgetExcel() {
  const k = kpis(), rows = tbl('AppBudget').filter(r => r.Category);
  const tT = rows.reduce((s, r) => s + (n(r.Target) || 0), 0), tS = rows.reduce((s, r) => s + (n(r.Spent) || 0), 0);
  const over = rows.filter(r => n(r.Spent) > n(r.Target));
  return `<section class="card">
    <h2>${esc(fMon(pDate(k.month)))} targets from Excel</h2>
    <div class="row"><span style="font-size:26px;font-weight:700" class="num">${inr(tS)}</span><span class="muted small num">of ${inr(tT)}</span></div>
    <div class="bar ${tS > tT ? 'over' : ''}" style="margin-top:8px"><i style="width:${Math.min(100, tT ? tS / tT * 100 : 0).toFixed(1)}%"></i></div>
    <p class="small ${tS > tT ? 'bad' : 'good'}" style="margin:8px 0 0">${tS > tT ? `${inr(tS - tT)} over budget` : `${inr(tT - tS)} left`}, ${over.length} of ${rows.length} categories over</p>
    <p class="small ${tone(k.pace_text)}" style="margin:4px 0 0">${esc(k.pace_text)}</p></section>
  <section class="sec"><div class="sec-h"><h3>By category</h3></div>
    ${rows.map(r => { const t = n(r.Target), s = n(r.Spent), u = t > 0 ? s / t : (s > 0 ? 1.01 : 0);
      return `<div class="brow"><div class="row"><span>${esc(r.Category)}</span><span class="num"><b class="${s > t ? 'bad' : ''}">${inr(s)}</b> <span class="lim">/ ${inr(t)}</span></span></div>
      <div class="meter thin ${s > t ? 'over' : ''}"><i style="width:${Math.min(100, u * 100).toFixed(1)}%"></i></div></div>`; }).join('')}
    <p class="small muted" style="margin:12px 0 0">Targets come from Budget Planning in Excel – they add up to ${inr(tT)}${isFinite(netPay()) ? `, ${Math.round(tT / netPay() * 100)}% of your net pay` : ''}.</p></section>`;
}

const DONUT_COLORS = ['#a9cdf5', '#6fa4e0', '#3f74b3', '#2a4d78', '#8c9aab', '#4d5866'];
function donut(parts, centerLabel, centerValue) {
  const tot = parts.reduce((s, p) => s + p.v, 0); if (!(tot > 0)) return '';
  const R = 80, C = 2 * Math.PI * R, gap = parts.length > 1 ? 3 : 0;
  let off = 0, arcs = '';
  parts.forEach((p, i) => {
    const len = Math.max(0, p.v / tot * C - gap);
    arcs += `<circle cx="110" cy="110" r="${R}" fill="none" stroke="${DONUT_COLORS[i % DONUT_COLORS.length]}" stroke-width="26"
      stroke-dasharray="${len.toFixed(2)} ${(C - len).toFixed(2)}" stroke-dashoffset="${(-off).toFixed(2)}" transform="rotate(-90 110 110)"><title>${esc(p.name)}: ${inr(p.v)}</title></circle>`;
    off += p.v / tot * C;
  });
  return `<div class="donut-wrap"><svg viewBox="0 0 220 220" role="img" aria-label="${esc(centerLabel)} ${inr(tot)}">${arcs}
    <text x="110" y="100" text-anchor="middle" font-size="15" fill="var(--muted)">${esc(centerLabel)}</text>
    <text x="110" y="128" text-anchor="middle" font-size="26" font-weight="600" fill="var(--text)">${esc(centerValue || cr(tot))}</text></svg></div>
    <div class="dlegend">${parts.map((p, i) => `<span><i style="background:${DONUT_COLORS[i % DONUT_COLORS.length]}"></i>${esc(p.name)} <em>${Math.round(p.v / tot * 100)}%</em></span>`).join('')}</div>`;
}
function spendParts() {
  const rows = tbl('AppBudget').filter(r => r.Category && n(r.Spent) > 0).map(r => ({ name: r.Category, v: n(r.Spent) })).sort((x, y) => y.v - x.v);
  const top = rows.slice(0, 5), rest = rows.slice(5).reduce((s, r) => s + r.v, 0);
  return rest > 0 ? top.concat([{ name: 'Other', v: rest }]) : top;
}
function viewBudget() {
  const seg = `<div class="seg seg2" id="bud-mode"><button data-b="plan" class="${state.budgetMode === 'plan' ? 'on' : ''}">Net-pay plan</button><button data-b="excel" class="${state.budgetMode === 'excel' ? 'on' : ''}">Excel targets</button></div>`;
  if (state.budgetMode === 'excel') return seg + viewBudgetExcel();
  if (state.planEdit) return seg + viewPlanEdit();
  const k = kpis(), net = netPay(), p = activePlan(), rows = tbl('AppBudget').filter(r => r.Category);
  const limit = net - (Number(p.savings) || 0);
  const spentBy = Object.fromEntries(rows.map(r => [r.Category, n(r.Spent) || 0]));
  const cats = Object.keys(p.limits || {}).filter(c => p.limits[c] > 0 || spentBy[c] > 0);
  rows.forEach(r => { if (!cats.includes(r.Category) && spentBy[r.Category] > 0) cats.push(r.Category); });
  const over = cats.filter(c => spentBy[c] > (Number(p.limits[c]) || 0));
  const m = pDate(kpis().month);
  return seg + netPayCard(true) + `
  <section class="sec box"><div class="sec-h"><h3>Where it went</h3><span class="small muted">${m ? MONF[m.getMonth()] : ''}</span></div>
    ${donut(spendParts(), 'Spent', cr(n(kpis().spending_month)))}</section>
  <section class="sec box"><div class="sec-h"><h3>Limits</h3><span class="small">${over.length} of ${cats.length} over</span></div>
    ${state.plan ? '' : `<p class="small muted" style="margin:0 0 6px">Suggested limits that fit your net pay. Edit them to make them yours.</p>`}
    <div class="rows">${cats.map(c => { const t = Number(p.limits[c]) || 0, s = spentBy[c] || 0, u = t > 0 ? s / t : (s > 0 ? 1.01 : 0);
      return `<div class="brow"><div class="row"><span>${esc(c)}</span><span class="num"><b class="${s > t ? 'bad' : ''}">${inr(s)}</b> <span class="lim">/ ${inr(t)}</span></span></div>
      <div class="meter thin ${s > t ? 'over' : ''}"><i style="width:${Math.min(100, u * 100).toFixed(1)}%"></i></div></div>`; }).join('')}</div>
    <div class="kv" style="margin-top:6px"><span class="muted">Saved first</span><span class="num">${inr(p.savings || 0)}</span></div>
    <div class="kv"><span class="muted">Limits add up to</span><span class="num">${inr(planTotal(p))} of ${inr(limit)}</span></div>
    <button class="secondary" id="plan-edit">Edit plan</button>
  </section>`;
}
function viewPlanEdit() {
  const e = state.planEdit, net = netPay(), avail = net - (Number(e.savings) || 0), tot = planTotal(e), ok = tot <= avail;
  const cats = budgetCats();
  return `<section class="card"><h2>Your net pay: ${inr(net)}</h2>
    <p class="small muted" style="margin:0 0 6px">Your latest net pay from payslips. Save first, then split the rest. Limits can’t add up to more than what’s left.</p>
    <label class="f" for="pe-sav">Save / invest every month</label>
    <input class="inp num" id="pe-sav" inputmode="numeric" value="${esc(e.savings)}">
    ${cats.map(c => `<div class="pe-row"><label for="pe-${esc(c)}">${esc(c)}${FIXED.has(c) ? ' <span class="muted small">fixed</span>' : ''}</label>
      <input class="inp num pe-inp" data-cat="${esc(c)}" id="pe-${esc(c)}" inputmode="numeric" value="${esc(e.limits[c] || 0)}"></div>`).join('')}
    <div class="kv" style="margin-top:10px"><span>Allocated</span><span class="num ${ok ? 'good' : 'bad'}" id="pe-tot">${inr(tot)} of ${inr(avail)}</span></div>
    <p class="small ${ok ? 'muted' : 'bad'}" id="pe-msg" style="margin:4px 0 0">${ok ? `${inr(avail - tot)} not allocated – a buffer for surprises.` : `${inr(tot - avail)} over your net pay – lower some limits or the saving amount.`}</p>
    <button class="primary" id="pe-save" ${ok ? '' : 'disabled'}>Save plan</button>
    <button class="secondary" id="pe-reset">Reset to suggested</button>
    <button class="secondary" id="pe-cancel">Cancel</button></section>`;
}
function bindBudget() {
  document.querySelectorAll('#bud-mode button').forEach(b => b.onclick = () => { state.budgetMode = b.dataset.b; LS.set('budgetMode', state.budgetMode); state.planEdit = null; render(); });
  const ed = document.getElementById('plan-edit');
  if (ed) ed.onclick = () => { const p = activePlan(); state.planEdit = { savings: p.savings || 0, limits: { ...p.limits } }; render(); window.scrollTo(0, 0); };
  if (!state.planEdit) return;
  const e = state.planEdit, num = v => Math.max(0, Math.round(Number(String(v).replace(/[^\d.]/g, '')) || 0));
  const refresh = () => {
    const avail = netPay() - e.savings, tot = planTotal(e), ok = tot <= avail;
    const t = document.getElementById('pe-tot'); t.textContent = `${inr(tot)} of ${inr(avail)}`; t.className = 'num ' + (ok ? 'good' : 'bad');
    const m = document.getElementById('pe-msg'); m.className = 'small ' + (ok ? 'muted' : 'bad');
    m.textContent = ok ? `${inr(avail - tot)} not allocated – a buffer for surprises.` : `${inr(tot - avail)} over your net pay – lower some limits or the saving amount.`;
    document.getElementById('pe-save').disabled = !ok;
  };
  document.getElementById('pe-sav').oninput = ev => { e.savings = num(ev.target.value); refresh(); };
  document.querySelectorAll('.pe-inp').forEach(inp => inp.oninput = () => { e.limits[inp.dataset.cat] = num(inp.value); refresh(); });
  document.getElementById('pe-reset').onclick = () => { const s = suggestedPlan(); state.planEdit = { savings: s.savings, limits: s.limits }; render(); };
  document.getElementById('pe-cancel').onclick = () => { state.planEdit = null; render(); };
  document.getElementById('pe-save').onclick = async () => {
    const b = document.getElementById('pe-save'); b.disabled = true; b.textContent = 'Saving…';
    const plan = { savings: e.savings, limits: e.limits, netPay: netPay(), updatedAt: new Date().toISOString() };
    state.plan = plan; LS.set('plan', plan); state.planEdit = null;
    try { await GH.putJson('plan.json', plan, 'Net-pay plan'); toast('Plan saved'); }
    catch (err) { toast('Saved on this phone – couldn’t reach GitHub, alerts use it after the next save'); }
    render();
  };
}

function projChart(rows) {
  if (!rows.length) return '';
  const W = 340, H = 190, pl = 36, pr = 10, pt = 10, pb = 22, iw = W - pl - pr, ih = H - pt - pb;
  const max = Math.max(1, ...rows.map(r => Math.max(n(r.NetWorth) || 0, n(r.Needed) || 0)));
  const x = i => pl + (i / (rows.length - 1 || 1)) * iw, y = v => pt + ih - (Math.max(0, n(v) || 0) / max) * ih;
  const path = key => rows.map((r, i) => `${i ? 'L' : 'M'}${x(i).toFixed(1)},${y(r[key]).toFixed(1)}`).join('');
  let g = '';
  [0, .5, 1].forEach(f => { const yy = pt + ih - f * ih; g += `<line x1="${pl}" x2="${W - pr}" y1="${yy}" y2="${yy}" stroke="#2b373b"/><text x="${pl - 4}" y="${yy + 3}" text-anchor="end" font-size="9.5" fill="var(--muted)">${f ? Math.round(max * f / 1e7) + ' Cr' : '0'}</text>`; });
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
  return `<section class="hero">
    <p class="hero-k">Net worth at ${ageR}</p>
    <h2 class="hero-v num">${cr(k.nw_at_60)}</h2>
    <p class="hero-s"><b>${cr(k.nw_at_60_today)}</b> in today’s money. <span class="${tone(k.retire_status)}">${esc(k.retire_status)}</span></p></section>
  <section class="sec"><div class="sec-h"><h3>Retirement checks</h3></div>
    <div class="kv"><span>Earliest retirement</span><span>${isFinite(ea) ? 'Age ' + ea : esc(k.earliest_retire_age)} <span class="small ${tone(k.earliest_retire_status)}">${esc(k.earliest_retire_status)}</span></span></div>
    <div class="kv"><span>Money lasts until</span><span>${esc(k.money_lasts)} <span class="small ${tone(k.money_lasts_status)}">${esc(k.money_lasts_status)}</span></span></div>
    <div class="kv"><span>Corpus needed at ${ageR}</span><span class="num">${cr(k.corpus_needed)}</span></div>
    <div class="kv"><span>Investing per month</span><span class="num">${inr(k.invest_actual_month)} <span class="muted small">actual, plan ${inr(k.invest_plan_month)}</span></span></div></section>
  <section class="sec"><div class="sec-h"><h3>At your actual SIP pace</h3></div>
    <div class="row"><span style="font-size:24px;font-weight:700" class="num">${cr(k.actual_pace_60)}</span><span class="small ${tone(k.actual_pace_status)}">${esc(k.actual_pace_status)}</span></div>
    <p class="small ${tone(k.actual_pace_note)}" style="margin:8px 0 0">${esc(k.actual_pace_note)}</p></section>
  <section class="sec"><div class="sec-h"><h3>Path to ${ageR}</h3></div>${projChart(rows)}
    <table class="proj" style="margin-top:10px"><tr><th>Year</th><th>Age</th><th>Net worth</th><th>Today’s ₹</th></tr>
    ${pick.map(r => `<tr><td>${r.Year}</td><td>${r.Age}</td><td class="num">${cr(r.NetWorth)}</td><td class="num">${cr(r.NetWorthToday)}</td></tr>`).join('')}</table>
    <p class="small muted" style="margin:10px 0 0">From NW Projection in Excel (car and house plans included).</p></section>`;
}

/* ---------------- More: every Excel sheet ---------------- */
const SHEET_ICON = { Home: '⌂', Insights: '📊', 'Monthly Compare': '📅', Goals: '🎯', Loans: '🏦', LIC: '🛡', 'Stocks & MF': '📈',
  'Yearly Calendar': '🗓', 'Side Income': '💼', 'Goal Planner': '🧭', 'NW Projection': '🚀', 'Savings Schemes': '🏛', 'Big Expenses': '💸',
  Trends: '〰', 'Yearly Summary': '📆', 'Budget Planning': '📝', Balances: '💰', Salary: '🧾', Payslips: '📄' };
function sheetList() { return (state.sheets && state.sheets.sheets) || []; }
/* Coaching receipts live in receipts.html on the same site, so their saved copy is readable here */
function receiptsSummary() {
  let list = [];
  try { list = JSON.parse(localStorage.getItem('crm_receipts_v1') || '[]') || []; } catch (e) { list = []; }
  const inrOf = r => r.cur === '₹' ? Number(r.total) || 0 : Number(r.inr) || 0;
  const d = new Date(), ym = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  const monthOf = r => r.month || String(r.date || '').slice(0, 7);
  const thisMonth = list.filter(r => monthOf(r) === ym);
  const last = list.slice().sort((x, y) => String(y.date || '').localeCompare(String(x.date || '')))[0];
  return { count: list.length, month: thisMonth.reduce((s, r) => s + inrOf(r), 0), monthCount: thisMonth.length,
    all: list.reduce((s, r) => s + inrOf(r), 0), last };
}
function viewMore() {
  const list = sheetList(), at = state.sheets && state.sheets.exportedAt ? new Date(state.sheets.exportedAt) : null;
  const rc = receiptsSummary();
  const row = (attrs, title, sub) => `<${attrs.startsWith('href') ? 'a' : 'button'} class="navrow" ${attrs}><span class="nt">${title}${sub ? `<br><span class="small muted">${sub}</span>` : ''}</span><span class="chev">›</span></${attrs.startsWith('href') ? 'a' : 'button'}>`;
  return `<section class="sec"><div class="sec-h"><h3>Coaching</h3>${rc.count ? `<span class="small">${inr(rc.month)} this month</span>` : ''}</div>
      <div class="group">
        ${row('href="receipts.html"', 'New receipt', '')}
        ${row('href="receipts.html#history"', 'Past receipts', rc.count ? `${rc.count} receipts, ${inr(rc.all)} in all${rc.last ? `. Last: ${esc(rc.last.client)}, ${fDay(pDate(rc.last.date))}` : ''}` : '')}
      </div></section>
    <section class="sec"><div class="sec-h"><h3>Plans</h3></div>
      <div class="group">
        ${row('data-go="future"', 'Retirement', '')}
        ${row('data-go="budget"', 'Net-pay plan', '')}
        ${row('data-go="settings"', 'Notifications and settings', pushStatus() === 'on' ? 'Notifications on' : '')}
      </div></section>
    <section class="sec"><div class="sec-h"><h3>Excel sheets</h3>${at ? `<span class="small">sent ${fDay(at).replace(/ \d{4}$/, '')}</span>` : ''}</div>
      ${list.length ? `<div class="group">${list.map(s => row(`data-sheet="${esc(s.name)}"`, esc(s.name) + (s.hidden ? ' <span class="small muted">(hidden in Excel)</span>' : ''), '')).join('')}</div>`
        : `<p class="small muted">No sheets yet. Run AppLinkNow in Excel, then tap ⟳.</p>`}
    </section>`;
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
    <div class="chips" id="sheet-mode" style="padding-top:0">${Object.entries(SHEET_VIEWS).map(([k, v]) => `<button class="chip ${state.sheetMode === k ? 'on' : ''}" data-m="${k}">${v[0]}</button>`).join('')}<button class="chip ${state.sheetMode === 'table' ? 'on' : ''}" data-m="table">Excel grid</button></div>`;
  if (state.sheetMode !== 'table') return top + chips + sheetBody(sh, rows, names);
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
const PRIMARY_HDR = /(used|now|owed|current|value|total|balance|net|amount|spent|left|corpus|worth)/i;
const isNum = t => NUMISH.test(String(t || '').trim());
const isStatus = t => /^[✓✅⚠ℹ🟡🔴]/.test(String(t || '').trim());
/* Parse a sheet into sections, then draw it in one of several views. */
function parseSheet(rows) {
  const out = [];
  let sec = null, headers = null, lastTiles = null;
  const open = title => { sec = { title, headers: null, items: [] }; out.push(sec); headers = null; lastTiles = null; };
  const add = it => { if (!sec) open(''); sec.items.push(it); };
  rows.forEach((r, ri) => {
    const c = r.c, b = new Set(r.b || []);
    const f = c.map((t, i) => t ? i : -1).filter(i => i >= 0);
    if (!f.length) return;
    const first = c[f[0]], allBold = f.every(i => b.has(i));
    if (lastTiles && lastTiles.n < 2 && (!allBold || lastTiles.n === 0) && f.every(i => lastTiles.pos.includes(i)) &&
        !f.some(i => isNum(c[i]) || /^→/.test(c[i]))) {
      f.forEach(i => lastTiles.el.tiles[lastTiles.pos.indexOf(i)].sub.push(c[i])); lastTiles.n++; return;
    }
    lastTiles = null;
    if (ri === 0 && b.has(f[0])) {
      open(first); const rest = f.slice(1).map(i => c[i]);
      if (rest.length) add({ type: 'note', t: rest.join(' ') }); return;
    }
    if (f.length === 1) {
      if (b.has(f[0]) && first.length < 70 && !isNum(first)) { open(first); return; }
      if (first.length > 30 || !b.has(f[0])) { add({ type: 'note', t: first }); return; }
    }
    const nonFirstNum = f.slice(1).filter(i => isNum(c[i])).length;
    if (allBold && f.length >= 2 && nonFirstNum === 0 && !isNum(first)) {
      if (!(sec && sec.items.length === 0 && sec.title)) open(first);
      headers = {}; f.forEach(i => { headers[i] = c[i]; }); sec.headers = headers; sec.headCols = f; return;
    }
    if (headers && isNum(first) && f.length >= 2) {
      const el = { type: 'tiles', tiles: f.map(i => ({ h: headers[i] || '', v: c[i], sub: [] })) };
      add(el); lastTiles = { el, pos: f, n: 0 }; return;
    }
    let vals = f.slice(1).map(i => ({ h: headers ? (headers[i] || '') : '', v: c[i], i }));
    if (vals.some(x => barPct(x.v) != null)) vals = vals.filter(x => x.h || !isNum(x.v));
    const nums = vals.filter(x => isNum(x.v));
    const prim = nums.find(x => /\b(used|now|owed|current|balance|spent)\b/i.test(x.h) && !/%/.test(x.h))
      || nums.find(x => PRIMARY_HDR.test(x.h) && !/%/.test(x.h)) || nums.find(x => !/%/.test(x.h)) || nums[0]
      || (f.length === 2 && !headers ? vals[0] : null);
    add({ type: 'row', label: first, bold: b.has(f[0]), prim, cells: c, cols: f,
          vals: vals.filter(x => x !== prim) });
  });
  return out.filter(s => s.items.length || s.title);
}
function barPct(t) { const m = String(t).match(/^[█░]+\s+(-?\d+)%/); return m ? Number(m[1]) : null; }
function barEl(p) { return `<div class="row small" style="gap:8px;width:100%"><div class="bar ${p > 100 ? 'over' : ''}" style="flex:1"><i style="width:${Math.min(100, Math.max(0, p))}%"></i></div><span class="num">${p}%</span></div>`; }
function vText(t, names) {
  const m = String(t).match(/^→\s*(.+?)(\s+for details)?$/);
  if (m && names.has(m[1].trim())) return `<button class="xl-link" data-sheet="${esc(m[1].trim())}">→ ${esc(m[1].trim())}</button>`;
  const p = barPct(t); if (p != null) return `${p}%`;
  return `<span class="${tone(t)}">${esc(t)}</span>`;
}
function subLine(vals) {
  const plain = vals.filter(x => barPct(x.v) == null && !isStatus(x.v));
  return plain.map(x => x.h ? `${esc(x.h)} <b class="num">${esc(x.v)}</b>` : `<span class="${tone(x.v)}">${esc(x.v)}</span>`).join(', ');
}
function sectionTotal(s) {
  const rows = s.items.filter(it => it.type === 'row' && it.prim && isNum(it.prim.v));
  const tot = rows.filter(r => r.bold).pop() || (rows.length === 1 ? rows[0] : null);
  return tot ? tot.prim.v : '';
}
const tileHtml = t => `<div class="stat"><div class="l">${esc(t.h)}</div><div class="v num ${tone(t.v)} ${String(t.v).startsWith('-') ? 'bad' : ''}">${esc(t.v)}</div>${t.sub.map(x => `<div class="s">${esc(x)}</div>`).join('')}</div>`;


/* Rows inside an opened section: one line per row, tap for the rest */
function viewList(secs, names) {
  return `<div class="vlist">${secs.map(s => `${s.title ? `<div class="vl-h">${esc(s.title)}</div>` : ''}${s.items.map(it => {
    if (it.type === 'note') return `<div class="vl-note ${tone(it.t)}">${vText(it.t, names)}</div>`;
    if (it.type === 'tiles') return it.tiles.map(t => `<div class="vl-row"><span class="vl-l">${esc(t.h)}</span><span class="vl-v num ${tone(t.v)}">${esc(t.v)}</span></div>`).join('');
    const p = it.vals.map(x => barPct(x.v)).find(x => x != null);
    const st = it.vals.find(x => isStatus(x.v)), sub = subLine(it.vals);
    const more = sub || st;
    return `<div class="vl-row ${it.bold ? 'b' : ''} ${more ? 'tap' : ''}"><span class="vl-l">${vText(it.label, names)}</span>
      <span class="vl-v num">${it.prim ? vText(it.prim.v, names) : ''}${p != null ? ` <span class="muted">${p}%</span>` : ''}${st ? ` <span class="${tone(st.v)}">${st.v.trim().charAt(0)}</span>` : ''}${more ? '<span class="chev2">›</span>' : ''}</span>
      ${more ? `<div class="vl-more">${sub}${st ? `<div class="${tone(st.v)}">${esc(st.v)}</div>` : ''}</div>` : ''}</div>`;
  }).join('')}`).join('')}</div>`;
}


/* Collapsible sections: titles + totals, tap to open (opened ones use the compact list) */
function viewFold(secs, names, sheet) {
  const openSet = state.fold[sheet] || (state.fold[sheet] = new Set());
  const allOpen = secs.every((s, k) => openSet.has(k));
  return `<div class="row small" style="justify-content:flex-end;margin:-4px 2px 6px"><button class="link-btn"  data-foldall="${allOpen ? 0 : 1}">${allOpen ? 'Close all' : 'Open all'}</button></div><div class="vlist">${secs.map((s, k) => {
    const isOpen = openSet.has(k), tot = sectionTotal(s);
    return `<div class="fold ${isOpen ? 'open' : ''}"><button class="fold-h" data-fold="${k}"><span class="fold-c">${isOpen ? '▾' : '▸'}</span><span class="vl-l">${esc(s.title || 'Section')}</span><span class="vl-v num">${esc(tot)}</span></button>
      ${isOpen ? viewList([{ ...s, title: '' }], names).replace('<div class="vlist">', '<div>') : ''}</div>`;
  }).join('')}</div>`;
}

/* Dense tiles: every main figure as a small tile, 3 per row */
function viewTiles(secs, names) {
  return secs.map(s => {
    const tiles = [];
    s.items.forEach(it => {
      if (it.type === 'tiles') it.tiles.forEach(t => tiles.push({ h: t.h, v: t.v, sub: [] }));
      if (it.type === 'row' && it.prim) tiles.push({ h: it.label, v: it.prim.v, sub: [], bold: it.bold });
    });
    const texts = s.items.filter(it => it.type === 'row' && !it.prim);
    if (!tiles.length && !texts.length) return '';
    return `<div class="vl-h">${esc(s.title || '')}</div>
      ${tiles.length ? `<div class="grid3 dtiles">${tiles.map(t => `<div class="stat ${t.bold ? 'tb' : ''}"><div class="l">${esc(t.h)}</div><div class="v num ${tone(t.v)} ${String(t.v).startsWith('-') ? 'bad' : ''}">${esc(t.v)}</div></div>`).join('')}</div>` : ''}
      ${texts.map(it => `<div class="vl-row"><span class="vl-l">${vText(it.label, names)}</span><span class="vl-v small muted">${subLine(it.vals)}</span></div>`).join('')}`;
  }).join('');
}

const SHEET_VIEWS = { fold: ['Sections', viewFold], tiles: ['Tiles', viewTiles] };
function sheetBody(sh, rows, names) {
  const secs = parseSheet(rows);
  if (!secs.length) return '<div class="empty">Nothing matches.</div>';
  const v = SHEET_VIEWS[state.sheetMode] || SHEET_VIEWS.fold;
  return v[1](secs, names, sh.name) + `<p class="small muted" style="text-align:center;margin-top:12px">Read-only – change things in Excel.</p>`;
}

function bindSheets() {
  document.querySelectorAll('[data-chl]').forEach(b => b.onclick = () => { state.choice[state.sheet] = { label: b.dataset.chl, v: b.dataset.chv }; render(); });
  document.querySelectorAll('#sheet-mode button').forEach(b => b.onclick = () => { state.sheetMode = b.dataset.m; LS.set('sheetMode2', state.sheetMode); render(); });
  document.querySelectorAll('[data-fold]').forEach(b => b.onclick = () => { const s = state.fold[state.sheet], k = Number(b.dataset.fold); s.has(k) ? s.delete(k) : s.add(k); render(); });
  document.querySelectorAll('[data-foldall]').forEach(b => b.onclick = () => { const n = document.querySelectorAll('[data-fold]').length;
    state.fold[state.sheet] = new Set(b.dataset.foldall === '1' ? [...Array(n).keys()] : []); render(); });
  document.querySelectorAll('.vl-row.tap').forEach(el => el.onclick = e => { if (!e.target.closest('button')) el.classList.toggle('open'); });
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
      toast(`Saved ${inr(entry.amount)}. Excel adds it next time it opens`);
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
  return `${state.cfg ? viewNotif() : ''}<section class="card">
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
`;
}
/* ---------------- phone notifications ---------------- */
const b64url = bytes => btoa(String.fromCharCode.apply(null, bytes)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const fromB64url = s => Uint8Array.from(atob(s.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((s.length + 3) % 4)), ch => ch.charCodeAt(0));
const isIOS = () => /iPhone|iPad|iPod/i.test(navigator.userAgent);
const standalone = () => window.navigator.standalone === true || (window.matchMedia && matchMedia('(display-mode: standalone)').matches);
function pushStatus() {
  if (!('serviceWorker' in navigator) || !('PushManager' in window) || !('Notification' in window))
    return isIOS() && !standalone() ? 'homescreen' : 'unsupported';
  if (isIOS() && !standalone()) return 'homescreen';
  if (Notification.permission === 'denied') return 'blocked';
  if (LS.get('pushId', null) && Notification.permission === 'granted') return 'on';
  return 'off';
}
function deviceName() { const u = navigator.userAgent; return /iPhone/.test(u) ? 'iPhone' : /iPad/.test(u) ? 'iPad' : /Android/.test(u) ? 'Android phone' : 'Browser'; }
function viewNotif() {
  const st = pushStatus();
  const body = {
    on: `<p class="small good" style="margin:0 0 4px">✓ On for this ${esc(deviceName())}</p>
         <p class="small muted" style="margin:0">Checked every morning at 8:30 and when Excel sends new figures. You’re told about new alerts, ⚠ warnings still open after a week, and a summary when a month ends.</p>
         <button class="secondary" id="n-test">Send a test notification</button><button class="secondary" id="n-off">Turn off on this phone</button>`,
    off: `<p class="small muted" style="margin:0">Get alerts on this phone for premiums due, spending over budget, out-of-date data, FDs maturing and a monthly summary.</p>
         <button class="primary" id="n-on">Turn on notifications</button>`,
    blocked: `<p class="small warn" style="margin:0">Notifications are blocked for this app. Allow them in your phone’s settings (Settings → Notifications → MyMoney, or the site settings in Chrome), then come back here.</p>`,
    homescreen: `<p class="small warn" style="margin:0">On iPhone, notifications only work when MyMoney is opened from the Home Screen. In Safari tap Share → Add to Home Screen, open MyMoney from there, then come back to ⚙.</p>`,
    unsupported: `<p class="small muted" style="margin:0">This browser can’t receive notifications. Use Chrome on Android, or Safari on iPhone (iOS 16.4 or later) with MyMoney added to the Home Screen.</p>`,
  }[st];
  return `<section class="card"><h2>Phone notifications</h2>${body}<p class="small" id="n-msg" style="margin:10px 0 0"></p><p class="small warn" id="n-wf" style="margin:6px 0 0" hidden></p></section>`;
}
async function ensureVapid() {
  const r = await GH.call('GET', 'push/vapid.json', { raw: true });
  if (r.ok) return r.json();
  if (r.status !== 404) throw httpError(r.status);
  // first phone: create the signing keys for notifications and keep them in the private data repo
  const kp = await crypto.subtle.generateKey({ name: 'ECDSA', namedCurve: 'P-256' }, true, ['sign', 'verify']);
  const pub = new Uint8Array(await crypto.subtle.exportKey('raw', kp.publicKey));
  const jwk = await crypto.subtle.exportKey('jwk', kp.privateKey);
  const v = { publicKey: b64url(pub), privateKey: jwk.d, appUrl: location.origin + location.pathname.replace(/index\.html$/, ''), createdAt: new Date().toISOString() };
  await GH.putJson('push/vapid.json', v, 'Notification keys');
  return v;
}
async function pushOn() {
  const perm = await Notification.requestPermission();
  if (perm !== 'granted') throw new AppError('Notifications weren’t allowed. Tap Turn on again and choose Allow.');
  const reg = await navigator.serviceWorker.ready;
  const v = await ensureVapid();
  let sub = await reg.pushManager.getSubscription();
  if (sub && sub.options && sub.options.applicationServerKey &&
      b64url(new Uint8Array(sub.options.applicationServerKey)) !== v.publicKey) { await sub.unsubscribe(); sub = null; }
  if (!sub) sub = await reg.pushManager.subscribe({ userVisibleOnly: true, applicationServerKey: fromB64url(v.publicKey) });
  const id = LS.get('pushId', null) || ('phone-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6));
  await GH.putJson(`push/subs/${id}.json`, { device: deviceName(), subscription: sub.toJSON(), createdAt: new Date().toISOString() }, 'Phone notifications on');
  LS.set('pushId', id);
}
async function pushOff() {
  const id = LS.get('pushId', null);
  try { const reg = await navigator.serviceWorker.ready; const sub = await reg.pushManager.getSubscription(); if (sub) await sub.unsubscribe(); } catch (e) {}
  if (id) await GH.remove(`push/subs/${id}.json`, 'Phone notifications off');
  LS.del('pushId');
}
function bindNotif() {
  const msg = document.getElementById('n-msg'); if (!msg) return;
  const busy = (b, t) => { b.disabled = true; b.textContent = t; };
  const fail = e => { render(); const m = document.getElementById('n-msg'); if (m) { m.className = 'small bad'; m.textContent = (e && e.message) || 'Something went wrong – try again.'; } };
  const on = document.getElementById('n-on');
  if (on) on.onclick = async () => { busy(on, 'Turning on…'); try { await pushOn(); render(); const m = document.getElementById('n-msg');
    m.className = 'small good'; m.textContent = 'Done. A “notifications are on” message should arrive within a few minutes.'; } catch (e) { fail(e); } };
  const test = document.getElementById('n-test');
  if (test) test.onclick = async () => { busy(test, 'Sending…'); try { await GH.putJson('push/test-request.json', { at: new Date().toISOString(), from: LS.get('pushId', '') }, 'Test notification');
    test.textContent = 'Test requested ✓'; msg.className = 'small muted'; msg.textContent = 'It should arrive within 1–3 minutes (GitHub has to start the checker).'; } catch (e) { fail(e); } };
  const off = document.getElementById('n-off');
  if (off) off.onclick = async () => { busy(off, 'Turning off…'); try { await pushOff(); render(); toast('Notifications off on this phone'); } catch (e) { fail(e); } };
  const wf = document.getElementById('n-wf');
  if (wf && pushStatus() === 'on') GH.exists('.github/workflows/notify.yml').then(ok => {
    if (!ok && wf.isConnected) { wf.hidden = false; wf.textContent = '⚠ The notification checker isn’t in your data repo yet. Upload notify.yml to mymoney-data/.github/workflows/ (see the update notes) – until then nothing is sent.'; }
  }).catch(() => {});
}
function bindSettings() {
  bindNotif();
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
    ['cfg', 'data', 'fetchedAt', 'pending', 'outbox', 'sheets', 'plan', 'pushId'].forEach(LS.del); location.reload(); };
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
if ('serviceWorker' in navigator && (location.protocol === 'https:' || /^(localhost|127\.0\.0\.1)$/.test(location.hostname))) navigator.serviceWorker.register('sw.js').catch(() => {});
{ const h = location.hash.slice(1); if (['home', 'tx', 'budget', 'more', 'add', 'settings'].includes(h) && state.cfg) state.tab = h; }
render();
if (state.cfg) sync();
