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
  plan: LS.get('plan', null),
  retireSaved: LS.get('retireSaved', null), retireEdit: null,          // net-pay plan {savings, limits:{cat:amt}}
  budgetMode: LS.get('budgetMode', 'plan'), planEdit: null,   // every sheet as Excel shows it
  sheet: null, sheetQ: '', choice: {},   // choice[sheet] = {label, v}
  sheetMode: ['fold', 'tiles', 'table'].includes(LS.get('sheetMode2', 'fold')) ? LS.get('sheetMode2', 'fold') : 'fold', fold: {},
  pending: LS.get('pending', []),    // [{id, sha, entry}] in GitHub inbox, not yet in Excel data
  outbox: LS.get('outbox', []),      // entries saved offline, not uploaded yet
  tab: 'home', lastError: null, syncing: false, stack: [], lastTab: 'home', goingBack: false,
  tx: { q: '', type: 'All', year: 'All', month: 'All', cat: 'All', limit: 100 },
  add: null,
  trades: LS.get('trades', []),     // buys/sells saved on the phone (trades/*.json), not yet in Excel
  prices: LS.get('prices', null),   // prices.json from the daily GitHub price check
  trade: null,
  balSaved: LS.get('balSaved', null), pfEdit: null,
  docs: Object.assign({ sips: [], bills: [], goals: [] }, LS.get('docs', {})), docEdit: null,   // app/*.json in the data repo
  history: LS.get('history', null),  // history.json: daily value of funds and stocks
  locked: false,
  recats: LS.get('recats', []), recat: null,   // category changes waiting for Excel (recat/*.json)   // PF / PPF figures saved on the phone (settings/balances.json)
};

/* ---------------- GitHub ---------------- */
const GH = {
  url(path, withRef) {
    const c = state.cfg;
    return `https://api.github.com/repos/${encodeURIComponent(c.owner)}/${encodeURIComponent(c.repo)}/contents/${path}` +
      (withRef ? `?ref=${encodeURIComponent(c.branch || 'main')}` : '');
  },
  async call(method, path, { body, raw } = {}) {
    if (method !== 'GET' && state.cfg && state.cfg.viewOnly) throw new AppError('This phone is view-only – changes are made on the owner’s phone.');
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
    try { const r = await GH.call('GET', 'settings/retirement.json', { raw: true }); if (r.ok) { state.retireSaved = await r.json(); LS.set('retireSaved', state.retireSaved); } } catch (e) { /* optional */ }
    try { const r = await GH.call('GET', 'settings/balances.json', { raw: true }); if (r.ok) { state.balSaved = await r.json(); LS.set('balSaved', state.balSaved); } } catch (e) { /* optional */ }
    try { const r = await GH.call('GET', 'prices.json', { raw: true }); if (r.ok) { state.prices = await r.json(); LS.set('prices', state.prices); } } catch (e) { /* optional */ }
    try { await refreshTrades(); } catch (e) { /* optional */ }
    await loadDocs();
    try { await refreshRecats(); } catch (e) { /* optional */ }
    try { const r = await GH.call('GET', 'history.json', { raw: true }); if (r.ok) { state.history = await r.json(); LS.set('history', state.history); } } catch (e) { /* optional */ }
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
const TITLES = { home: 'MyMoney', salary: 'Salary', money: 'Money', invest: 'Investments', tx: 'All transactions', add: 'Add transaction', budget: 'Net-pay plan', future: 'Retirement plan', trade: 'Add buy / sell', pfedit: 'Update PF / PPF', sips: 'SIPs', bills: 'Bills and subscriptions', goals: 'Goals', tax: 'Tax: 80C and gains', report: 'Monthly report', coaching: 'Coaching earnings', import: 'Import holdings', lockset: 'Lock, look and this phone', recat: 'Change category', loans: 'Loans', cover: 'Insurance cover', review: 'Monthly review', sort: 'Sort UPI – Other', shared: 'Shared expenses', year: 'Year in review', family: 'Family view', settings: 'Settings', more: 'More', sheet: 'Sheet' };
const NAV_OF = { trade: 'invest', pfedit: 'invest', sips: 'invest', tax: 'invest', import: 'invest', bills: 'money', goals: 'future', report: 'home', coaching: 'home', lockset: 'home', recat: 'money', loans: 'money', cover: 'future', review: 'money', sort: 'money', shared: 'money', year: 'home', family: 'home', tx: 'money', budget: 'money', add: 'money', sheet: 'home', more: 'home', settings: 'home' };
function render() {
  const v = document.getElementById('view');
  const t = document.getElementById('screen-title'), av = document.getElementById('avatar');
  if (state.tab === 'home' && state.data) {
    const nm = userName();
    const hr = new Date().getHours(), greet = hr < 12 ? 'Good morning' : hr < 17 ? 'Good afternoon' : 'Good evening';
    t.innerHTML = `<span class="hi">${greet}</span>Hi, ${esc(nm.split(/\s+/)[0])}`;
    av.hidden = false; av.textContent = nm.split(/\s+/).map(w => w[0]).join('').slice(0, 2).toUpperCase();
  } else { t.textContent = state.tab === 'sheet' ? state.sheet : TITLES[state.tab]; av.hidden = true; }
  document.querySelectorAll('.tab').forEach(b => b.classList.toggle('active', b.dataset.tab === (NAV_OF[state.tab] || state.tab)));
  document.querySelectorAll('.tab[data-tab]').forEach(b => { b.hidden = isViewer() && !famAllowed().includes(b.dataset.tab); });
  if (state.lastTab !== state.tab) { if (!state.goingBack && !MAIN_TABS.includes(state.tab)) state.stack.push(state.lastTab); if (MAIN_TABS.includes(state.tab)) state.stack = []; state.stack = state.stack.slice(-20); state.lastTab = state.tab; }
  state.goingBack = false;
  const bb = document.getElementById('btn-back'); if (bb) bb.hidden = MAIN_TABS.includes(state.tab) || state.tab === 'settings' && !state.cfg;
  renderFreshness();
  if (!state.cfg && state.tab !== 'settings') state.tab = 'settings';
  if (state.tab === 'settings') { v.innerHTML = viewSettings(); bindSettings(); return; }
  if (!state.data) {
    v.innerHTML = `<div class="empty">${state.syncing ? 'Loading your data…' : esc(state.lastError || 'No data loaded yet.')}<br><br>
      <button class="secondary" onclick="sync({toast:true})">Try again</button></div>`;
    return;
  }
  if (state.docEdit && state.docEdit.kind !== state.tab) state.docEdit = null;
  document.body.classList.toggle('viewer', isViewer());
  if (isViewer() && !viewerMay(state.tab)) state.tab = famAllowed()[0] || 'home';
  document.body.dataset.tab = state.tab;
  const views = { home: viewHome, salary: viewSalary, money: viewMoney, invest: viewInvest, tx: viewTx, add: viewAdd, budget: viewBudget, future: viewFuture, more: viewMore, sheet: viewSheet, trade: viewTrade, pfedit: viewPfEdit, sips: viewSips, bills: viewBills, goals: viewGoals, tax: viewTax, report: viewReport, coaching: viewCoaching, import: viewImport, lockset: viewLockSettings, recat: viewRecat, loans: viewLoans, cover: viewCover, review: viewReview, sort: viewSort, shared: viewShared, year: viewYear, family: viewFamily };
  v.innerHTML = subnav() + views[state.tab]();
  if (state.tab === 'more' || state.tab === 'sheet') bindSheets();
  if (state.tab === 'tx') bindTx();
  if (state.tab === 'add') bindAdd();
  if (state.tab === 'budget') bindBudget();
  if (state.tab === 'salary') bindSalary();
  if (state.tab === 'money') bindMoney();
  if (state.tab === 'invest') bindInvest();
  if (state.tab === 'trade') bindTrade();
  if (state.tab === 'pfedit') bindPfEdit();
  if (['sips', 'bills', 'goals'].includes(state.tab)) { if (state.docEdit) bindDocForm(state.tab); else bindDocList(state.tab); if (state.tab === 'goals' && !state.docEdit) bindGoals(); }
  if (state.tab === 'report') bindReport();
  if (state.tab === 'coaching') bindCoaching();
  if (state.tab === 'import') bindImport();
  if (state.tab === 'lockset') bindLockSettings();
  if (state.tab === 'recat') bindRecat();
  if (state.tab === 'loans') bindLoans();
  if (state.tab === 'cover') bindCover();
  if (state.tab === 'sort') bindSort();
  if (state.tab === 'shared') bindShared();
  if (state.tab === 'year') bindYear();
  if (state.tab === 'family') bindFamily();
  if (state.tab === 'coaching') bindCoachingMoney();
  document.querySelectorAll('[data-sheet-open]').forEach(b => b.onclick = () => { state.sheet = b.dataset.sheetOpen; state.sheetQ = ''; state.tab = 'sheet'; window.scrollTo(0, 0); render(); });
  if (state.tab === 'future') {
    if (state.retireEdit) bindRetireEdit();
    document.querySelectorAll('[data-q]').forEach(inp => inp.oninput = () => {
      const t = state.retireTry; t.x[inp.dataset.q] = Number(inp.value);
      if (t.x.retire <= t.x.age) t.x.retire = t.x.age + 1;
      document.getElementById('rqv-retire').textContent = t.x.retire + ' years';
      document.getElementById('rqv-invest').textContent = inr(t.x.invest);
      document.getElementById('rt-card').innerHTML = retireCard(t.x, t.b);
      const base = { ...retireInputs().x, ...(retirePending() && state.retireSaved ? state.retireSaved.values : {}) };
      document.getElementById('rq-save').hidden = t.x.retire === base.retire;
    });
    const qs = document.getElementById('rq-save');
    if (qs) qs.onclick = async () => {
      qs.disabled = true; qs.textContent = 'Saving…';
      const t = state.retireTry, v = {}; RETIRE_FIELDS.filter(f => f.row).forEach(f => { v[f.k] = t.x[f.k]; });
      const saved = { id: 'r' + Date.now().toString(36), values: v, savedAt: new Date().toISOString() };
      try { await GH.putJson('settings/retirement.json', saved, 'Retirement plan from the app'); state.retireSaved = saved; LS.set('retireSaved', saved); toast('Saved. Excel updates Goal Planner next time it opens'); }
      catch (err) { toast(err.message || 'Couldn’t reach GitHub – try again'); }
      render();
    };
    const ed = document.getElementById('rp-edit');
    if (ed) ed.onclick = () => { const t = state.retireTry; state.retireEdit = { x: { ...t.x }, b: t.b }; render(); window.scrollTo(0, 0); };
  }
  const ha = document.querySelector('[data-homeall]'); if (ha) ha.onclick = () => { state.homeAlertsAll = !state.homeAlertsAll; render(); };
  const hh = document.querySelector('[data-health]'); if (hh) hh.onclick = () => { state.healthOpen = !state.healthOpen; render(); };
  document.querySelectorAll('[data-gotab]').forEach(b => b.onclick = () => { state.tab = b.dataset.gotab; window.scrollTo(0, 0); render(); });
}

/* ---------------- easier navigation: sub-tabs, shortcuts, back ---------------- */
const MAIN_TABS = ['home', 'salary', 'money', 'invest', 'future'];
const SUBNAV = {
  money: [['money', 'Spending'], ['tx', 'Transactions'], ['review', 'Review'], ['bills', 'Bills'], ['shared', 'Shared'], ['sort', 'Sort UPI'], ['loans', 'Loans'], ['budget', 'Plan'], ['add', '＋ Add']],
  invest: [['invest', 'Overview'], ['sips', 'SIPs'], ['trade', 'Buy / sell'], ['pfedit', 'PF / PPF'], ['tax', 'Tax'], ['import', 'Import']],
  future: [['future', 'Retirement'], ['goals', 'Goals'], ['cover', 'Insurance']],
};
const VIEW_SUBS = new Set(['tx', 'bills', 'loans', 'sips', 'tax', 'review', 'year', 'goals', 'cover', 'shared', 'report', 'family', 'more', 'lockset', 'settings', 'sheet']);
function viewerMay(t) { const a = famAllowed(); return a.includes(t) || (VIEW_SUBS.has(t) && (a.includes(groupOf(t)) || ['family', 'lockset', 'settings', 'more'].includes(t))); }
function groupOf(tab) { return MAIN_TABS.includes(tab) ? tab : (NAV_OF[tab] || 'home'); }
function subnav() {
  const g = groupOf(state.tab), list = SUBNAV[g]; if (!list) return '';
  return `<div class="subnav" role="tablist">${list.filter(([t]) => !isViewer() || viewerMay(t)).map(([t, l]) => `<button class="sn ${state.tab === t ? 'on' : ''}" data-gotab="${t}" role="tab" aria-selected="${state.tab === t}">${l}</button>`).join('')}</div>`;
}
const SHORTCUTS = [['sips', '🔁', 'SIPs'], ['bills', '🧾', 'Bills'], ['trade', '📈', 'Buy / sell'], ['goals', '🎯', 'Goals'],
  ['tx', '📋', 'Transactions'], ['loans', '🏦', 'Loans'], ['tax', '🧮', 'Tax'], ['pfedit', '🏛', 'PF / PPF'],
  ['review', '🗓', 'Review'], ['shared', '🤝', 'Shared'], ['year', '🏆', 'Year'], ['report', '📄', 'Report'],
  ['coaching', '🎓', 'Coaching'], ['cover', '🛡', 'Insurance'], ['family', '👪', 'Family'], ['more', '⋯', 'More']];
function viewShortcuts() {
  return `<section class="sec"><div class="sec-h"><h3>Shortcuts</h3></div><div class="shortcuts">
    ${SHORTCUTS.filter(([t]) => !isViewer() || viewerMay(t)).map(([t, i, l]) => `<button class="sc" data-gotab="${t}"><span class="sc-i" aria-hidden="true">${i}</span><span>${l}</span></button>`).join('')}</div></section>`;
}
function goBack() {
  let t = state.stack.pop();
  while (t && t === state.tab) t = state.stack.pop();
  state.goingBack = true; state.docEdit = null; state.tab = t || (NAV_OF[state.tab] || 'home'); window.scrollTo(0, 0); render();
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
  salary: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><rect x="3" y="6" width="18" height="13" rx="2"/><path d="M3 10h18M7 15h3"/></svg>',
  limit: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M4 18a8 8 0 1 1 16 0"/><path d="M12 18l4-6"/></svg>',
  invest: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 17l6-6 4 4 8-8"/><path d="M15 7h6v6"/></svg>',
  retire: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>',
  add: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14M5 12h14"/></svg>',
  receipt: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3h12v18l-3-2-3 2-3-2-3 2z"/><path d="M9 8h6M9 12h6"/></svg>',
  grid: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="7.5" cy="7.5" r="3"/><circle cx="16.5" cy="7.5" r="3"/><circle cx="7.5" cy="16.5" r="3"/><circle cx="16.5" cy="16.5" r="3"/></svg>',
};
function latestPayslip(idx) {
  const sh = sheetList().find(s => s.name === 'Payslips'); if (!sh || sh.rows.length < 2) return null;
  const head = sh.rows[0].c, rows = sh.rows.slice(1).filter(r => r.c[0]);
  const pick = rows[idx == null ? rows.length - 1 : idx]; if (!pick) return null;
  const o = {}; head.forEach((h, k) => { o[h] = pick.c[k]; });
  return { row: o, months: rows.map(r => r.c[0]), index: idx == null ? rows.length - 1 : idx };
}
const num = t => { const v = Number(String(t == null ? '' : t).replace(/[₹,\s]/g, '')); return isFinite(v) ? v : 0; };
function viewHome() {
  const k = kpis(), m = pDate(k.month);
  const alerts = tbl('AppAlerts').map(x => String(x.Alert || '')).filter(x => /^[⚠ℹ]/.test(x));
  const recent = allTx().slice(0, 3), ps = latestPayslip();
  const net = netPay(), p = activePlan(), limit = net - (Number(p.savings) || 0), spent = n(k.spending_month) || 0, left = limit - spent;
  const qa = (attrs, icon, label) => `<${attrs.startsWith('href') ? 'a' : 'button'} class="qa" ${attrs}><span class="qa-i">${icon}</span><span>${label}</span></${attrs.startsWith('href') ? 'a' : 'button'}>`;
  return `
  <section class="hero">
    <p class="hero-k" style="margin-top:0">Total net worth</p>
    <h2 class="hero-v num">${inr(k.net_worth)}</h2>
    <p class="hero-s" style="margin:-6px 0 0">${inr(k.money_available)} in cash and bank</p>
    <div class="hero-tiles">
      <div><span>Income${m ? ', ' + MON[m.getMonth()] : ''}</span><b class="num">${inr(k.income_month)}</b></div>
      <div><span>Spent${m ? ', ' + MON[m.getMonth()] : ''}</span><b class="num">${inr(spent)}</b></div>
    </div>
  </section>
  ${viewNetWorthChart()}
  ${ps ? `<button class="strip" data-gotab="salary"><span class="strip-i">${IC.salary}</span><span class="strip-t"><span>Salary credited, ${esc(ps.row.Month)}</span><b class="num">${inr(num(ps.row['Net Pay']))}</b></span><span class="strip-a">Payslip ›</span></button>` : ''}
  ${isFinite(net) ? `<button class="strip" data-gotab="money"><span class="strip-i ${left < 0 ? 'bad' : 'good'}">${IC.limit}</span><span class="strip-t"><span>${left < 0 ? 'Over your net-pay limit' : 'Left from net pay'}</span><b class="num ${left < 0 ? 'bad' : 'good'}">${inr(Math.abs(left))}</b></span><span class="strip-a">Money ›</span></button>` : ''}
  <div class="qas">
    ${qa('data-gotab="add"', IC.add, 'Add expense')}
    ${qa('href="receipts.html"', IC.receipt, 'Receipt')}
    ${qa('data-gotab="invest"', IC.invest, 'Invest')}
    ${qa('data-gotab="future"', IC.retire, 'Retirement')}
  </div>
  ${viewShortcuts()}
  ${(() => { const d = tokenDays(); return d != null && d <= 14 ? `<section class="sec box warnbox"><div class="sec-h"><h3>GitHub token ${d < 0 ? 'has expired' : `expires in ${d} day${d === 1 ? '' : 's'}`}</h3></div><p class="small" style="margin:0">Make a new token on GitHub (same settings), then paste it in More → Notifications and settings and in Excel’s AppLink module.</p></section>` : ''; })()}
  ${viewLowBalance()}
  <section class="sec box">
    <div class="sec-h"><h3>Recent activity</h3><button data-gotab="tx">See all</button></div>
    ${recent.map(t => txRow(t, true)).join('')}
  </section>
  ${alerts.length ? `<section class="sec box"><div class="sec-h"><h3>Needs attention</h3><span class="small muted">${alerts.length}</span></div>
    ${(state.homeAlertsAll ? alerts : alerts.slice(0, 3)).map(x => `<div class="alert-row"><span class="dot ${/^⚠/.test(x) ? 'warn' : 'info'}"></span><span>${esc(x.replace(/^[⚠ℹ]\s*/, ''))}</span></div>`).join('')}
    ${alerts.length > 3 ? `<button class="more-btn" data-homeall>${state.homeAlertsAll ? 'Show fewer' : `Show all ${alerts.length}`}</button>` : ''}</section>` : ''}
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
  return p.concat(applyRecatsView(tbl('AppRecent')));
}
function txRow(t, compact) {
  const d = describe(t), a = n(t.Amount);
  const cls = t.Type === 'Income' ? 'good' : t.Type === 'Transfer' ? 'muted' : '';
  const sign = t.Type === 'Income' ? '+' : t.Type === 'Expenses' ? '−' : '';
  const tile = { Income: 'in', Savings: 'sv', Transfer: 'tr' }[t.Type] || '';
  const letter = (String(d.title).match(/[A-Za-z0-9]/) || ['•'])[0].toUpperCase();
  return `<div class="tx ${t._pending ? 'pending' : ''}"><div class="ic ${tile}">${letter}</div>
    <div class="mid"><div class="t1">${esc(d.title)}${t._recat ? '<span class="badge">Category change waiting for Excel</span>' : ''}${t._pending ? `<span class="badge">${t._pending === 'outbox' ? 'Not uploaded' : 'Waiting for Excel'}</span>` : ''}</div>
    <div class="t2">${compact ? fDay(pDate(t.Date)).replace(/ \d{4}$/, '') + ', ' : ''}${esc(t.Category)}</div>
    <div class="t3">${esc(fullDetails(t))}<br>${esc(t.Type)}, ${esc(d.src)}, ${fDay(pDate(t.Date))}</div>
    ${t._pending === 'pending' && !compact ? `<button class="link-btn blue" data-edit="${esc(t._id)}">Edit</button> <button class="link-btn" data-del="${esc(t._id)}" data-sha="${esc(t._sha)}">Remove</button>` : ''}
    ${t._pending === 'outbox' && !compact ? `<button class="link-btn" data-delout="${esc(t._id)}">Remove</button>` : ''}
    ${!t._pending && !compact && /^bank:/i.test(String(t.Details || '')) ? `<button class="link-btn blue" data-recat="${esc(txKey(t))}">Change category</button>` : ''}</div>
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
  document.querySelectorAll('[data-edit]').forEach(b => b.onclick = () => editPending(b.dataset.edit));
  document.querySelectorAll('[data-recat]').forEach(b => b.onclick = () => { const t = allTx().find(x => !x._pending && txKey(x) === b.dataset.recat); if (t) openRecat(t); });
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
/* Same formulas as Goal Planner C16:C28 */
function retirePlan(x, b) {
  const Y = Math.max(0, x.retire - x.age), N = Math.max(0, x.until - x.retire);
  const costR = x.cost * Math.pow(1 + x.inf, Y);                               // C18
  const rr = (1 + x.ra) / (1 + x.inf) - 1;                                     // C19
  const corpus = costR * 12 * (rr ? (1 - Math.pow(1 + rr, -N)) / rr * (1 + rr) : N);   // C20 PV(...,1)
  const epfFV = b.epfToday * Math.pow(1 + b.epfRate, Y);                       // C24
  const ie = b.epfRate / 12;
  const epfFut = b.epfMonthly * (ie ? (Math.pow(1 + ie, Y * 12) - 1) / ie : Y * 12);   // C25 FV
  const need = Math.max(0, corpus - epfFV - epfFut);                           // C26
  const i = x.rb / 12, n = Y * 12;
  const sipFlat = n ? (i ? need * i / (Math.pow(1 + i, n) - 1) : need / n) : 0;       // C27 PMT
  const grow = Math.abs(x.rb - x.step) < 1e-9 ? Y * Math.pow(1 + x.rb, Y - 1)
    : (Math.pow(1 + x.rb, Y) - Math.pow(1 + x.step, Y)) / (x.rb - x.step);
  const sipStart = Y ? need / (12 * grow) : 0;                                  // C28
  const atPace = (x.invest || 0) * 12 * grow + epfFV + epfFut;                 // your pace, same step-up model
  return { Y, N, costR, rr, corpus, epfFV, epfFut, need, sipFlat, sipStart, atPace, gap: corpus - atPace };
}

/* Goal Planner inputs, read from the sheet Excel sends (so the editor starts from your real plan) */
const RETIRE_FIELDS = [
  { k: 'retire', label: 'Retire at age', row: 'Retirement age', min: 35, max: 75, step: 1, fmt: v => `${v}` },
  { k: 'until', label: 'Money should last until age', row: 'Plan money until age', min: 70, max: 100, step: 1, fmt: v => `${v}` },
  { k: 'cost', label: 'Monthly living cost today', row: 'Monthly living cost today', min: 10000, max: 400000, step: 5000, fmt: v => inr(v) },
  { k: 'invest', label: 'You invest each month', row: null, min: 0, max: 300000, step: 1000, fmt: v => inr(v), hint: 'Only for this what-if. Excel takes it from your SIPs.' },
  { k: 'step', label: 'Raise investing each year by', row: 'Yearly SIP step-up', min: 0, max: .25, step: .01, fmt: v => pct(v) },
  { k: 'rb', label: 'Return before retirement', row: 'Return before retirement', min: .04, max: .16, step: .005, fmt: v => pct(v) },
  { k: 'ra', label: 'Return after retirement', row: 'Return after retirement', min: .03, max: .12, step: .005, fmt: v => pct(v) },
  { k: 'inf', label: 'Inflation', row: 'Inflation (yearly)', min: .03, max: .10, step: .005, fmt: v => pct(v) },
  { k: 'age', label: 'Your age now', row: 'Current age', min: 18, max: 65, step: 1, fmt: v => `${v}` },
];
function sheetRowValue(sheetName, label) {
  const sh = sheetList().find(s => s.name === sheetName); if (!sh) return null;
  const r = sh.rows.find(x => (x.c.find(t => t) || '').trim().toLowerCase().startsWith(label.toLowerCase()));
  if (!r) return null;
  const vals = r.c.filter(t => t); const t = vals[1]; if (t == null) return null;
  const num = Number(String(t).replace(/[₹,\s]/g, '').replace(/%$/, ''));
  return isFinite(num) ? (/%\s*$/.test(t) ? num / 100 : num) : null;
}
function retireInputs() {
  const k = kpis(), g = l => sheetRowValue('Goal Planner', l);
  const x = { age: g('Current age') ?? n(k.current_age) ?? 31, retire: g('Retirement age') ?? n(k.retire_age) ?? 60, until: g('Plan money until age') ?? 85,
    inf: g('Inflation (yearly)') ?? .06, rb: g('Return before retirement') ?? .11, ra: g('Return after retirement') ?? .07,
    cost: g('Monthly living cost today') ?? 80000, step: g('Yearly SIP step-up') ?? .10, invest: Math.round(n(k.invest_actual_month) || 0) };
  const b = { epfToday: g('EPF + PPF today') ?? 0, epfRate: g('EPF interest rate') ?? .0825, epfMonthly: g('Monthly EPF going in') ?? 0 };
  return { x, b };
}
function retirePending() {
  const s = state.retireSaved; if (!s || !s.id) return false;
  return !(state.data && state.data.retireApplied === s.id);
}
function viewRetireEdit() {
  const e = state.retireEdit;
  return `<section class="card"><h2>Change your retirement plan</h2>
    <p class="small muted" style="margin:0">Slide to see the effect straight away. Save to Excel to make it your plan.</p>
    ${RETIRE_FIELDS.map(f => `<div class="rf"><div class="row"><label for="rf-${f.k}">${f.label}</label><b class="num" id="rv-${f.k}">${f.fmt(e.x[f.k])}</b></div>
      <input type="range" id="rf-${f.k}" data-k="${f.k}" min="${f.min}" max="${f.max}" step="${f.step}" value="${e.x[f.k]}">
      ${f.hint ? `<p class="small muted" style="margin:2px 0 0">${f.hint}</p>` : ''}</div>`).join('')}
  </section>
  <section class="card" id="rp-res">${retireResults(e.x, e.b)}</section>
  <button class="primary" id="rp-save">Save to Excel</button>
  <button class="secondary" id="rp-reset">Reset to Excel’s values</button>
  <button class="secondary" id="rp-cancel">Cancel</button>
  <p class="note">Excel writes these into Goal Planner the next time the workbook opens, or when you run AppLinkNow, then recalculates the whole plan. Your monthly investing isn’t saved; Excel works it out from your SIPs.</p>`;
}
function retireResults(x, b) {
  const r = retirePlan(x, b), ok = r.gap <= 0;
  return `<p class="hero-k" style="margin:0;color:var(--muted)">You’ll need at ${x.retire}</p>
    <h2 class="num" style="margin:2px 0 4px;font-size:34px;font-weight:600">${cr(r.corpus)}</h2>
    <p class="small muted" style="margin:0">To spend ${inr(r.costR)} a month from ${x.retire} until ${x.until}, rising with inflation.</p>
    <div class="kv" style="margin-top:12px"><span>Investing ${inr(x.invest)} a month, +${pct(x.step)} a year</span><b class="num ${ok ? 'good' : 'bad'}">${cr(r.atPace)}</b></div>
    <p class="small ${ok ? 'good' : 'bad'}" style="margin:6px 0 8px">${ok ? `On track, ${cr(-r.gap)} more than you need.` : `Short by ${cr(r.gap)}.`}</p>
    <div class="kv"><span>Start SIP at, then raise ${pct(x.step)} a year</span><b class="num">${inr(r.sipStart)}</b></div>
    <div class="kv"><span>Or the same SIP every month</span><b class="num">${inr(r.sipFlat)}</b></div>
    <div class="kv"><span class="muted">EPF and PPF at ${x.retire}</span><span class="num">${cr(r.epfFV + r.epfFut)}</span></div>
    <div class="kv"><span class="muted">Still needed from your investments</span><span class="num">${cr(r.need)}</span></div>`;
}
function bindRetireEdit() {
  const e = state.retireEdit;
  document.querySelectorAll('#main-rf input[type=range], input[type=range][data-k]').forEach(inp => inp.oninput = () => {
    e.x[inp.dataset.k] = Number(inp.value);
    const f = RETIRE_FIELDS.find(z => z.k === inp.dataset.k); document.getElementById('rv-' + f.k).textContent = f.fmt(e.x[f.k]);
    if (e.x.until <= e.x.retire) { e.x.until = e.x.retire + 1; document.getElementById('rf-until').value = e.x.until; document.getElementById('rv-until').textContent = e.x.until; }
    if (e.x.retire <= e.x.age) { e.x.retire = e.x.age + 1; document.getElementById('rf-retire').value = e.x.retire; document.getElementById('rv-retire').textContent = e.x.retire; }
    document.getElementById('rp-res').innerHTML = retireResults(e.x, e.b);
  });
  document.getElementById('rp-cancel').onclick = () => { state.retireEdit = null; render(); window.scrollTo(0, 0); };
  document.getElementById('rp-reset').onclick = () => { const s = retireInputs(); state.retireEdit = { x: s.x, b: s.b }; render(); };
  document.getElementById('rp-save').onclick = async () => {
    const btn = document.getElementById('rp-save'); btn.disabled = true; btn.textContent = 'Saving…';
    const v = {}; RETIRE_FIELDS.filter(f => f.row).forEach(f => { v[f.k] = e.x[f.k]; });
    const saved = { id: 'r' + Date.now().toString(36), values: v, savedAt: new Date().toISOString() };
    try {
      await GH.putJson('settings/retirement.json', saved, 'Retirement plan from the app');
      state.retireSaved = saved; LS.set('retireSaved', saved); state.retireTry = { x: { ...e.x }, b: e.b }; state.retireEdit = null;
      toast('Saved. Excel updates Goal Planner next time it opens'); render(); window.scrollTo(0, 0);
    } catch (err) { btn.disabled = false; btn.textContent = 'Save to Excel'; toast(err.message || 'Couldn’t reach GitHub – try again'); }
  };
}
function viewFuture() {
  if (state.retireEdit) return viewRetireEdit();
  const k = kpis(), rows = tbl('AppProjection').filter(r => r.Year !== '' && r.Year != null);
  const ageR = n(k.retire_age) || 60;
  const pick = rows.filter((r, i) => i === 0 || n(r.Age) % 5 === 0 || n(r.Age) === ageR);
  const ea = n(k.earliest_retire_age);
  const ri = retireInputs(), base = { ...ri.x, ...(retirePending() && state.retireSaved ? state.retireSaved.values : {}) };
  const tx = state.retireTry || (state.retireTry = { x: { ...base }, b: ri.b });
  const changed = tx.x.retire !== base.retire;
  const ep = sheetRowCells('Savings Schemes', 'EPF'), pp = sheetRowCells('Savings Schemes', 'PPF'), fd = sheetRowCells('Savings Schemes', 'Fixed deposits');
  const today = [ep && ['Employees’ Provident Fund (EPF)', num(ep[2])], pp && ['Public Provident Fund (PPF)', num(pp[2])], fd && num(fd[2]) > 0 && ['Fixed deposits', num(fd[2])]].filter(Boolean);
  return `<section class="hero" id="rt-card">${retireCard(tx.x, tx.b)}</section>
  ${retirePending() ? `<p class="banner" style="margin:14px 0 0">Your new plan is saved. Excel applies it next time the workbook opens; the Excel figures below update after that.</p>` : ''}
  <section class="sec box"><div class="sec-h"><h3>Adjust your plan</h3><button id="rp-edit">More settings</button></div>
    <div class="rf"><div class="row"><label for="rq-retire">Retirement age</label><b class="num" id="rqv-retire">${tx.x.retire} years</b></div>
      <input type="range" id="rq-retire" data-q="retire" min="35" max="75" step="1" value="${tx.x.retire}"></div>
    <div class="rf" style="border-bottom:none"><div class="row"><label for="rq-invest">Monthly investment</label><b class="num" id="rqv-invest">${inr(tx.x.invest)}</b></div>
      <input type="range" id="rq-invest" data-q="invest" min="0" max="300000" step="1000" value="${tx.x.invest}"></div>
    <p class="small muted" style="margin:0">Assumes ${pct(tx.x.rb)} return before retirement, ${pct(tx.x.step)} more invested each year, ${pct(tx.x.inf)} inflation.</p>
    <button class="primary" id="rq-save" ${changed ? '' : 'hidden'}>Save retirement age to Excel</button>
  </section>
  ${today.length ? `<section class="sec box"><div class="sec-h"><h3>Retirement savings today</h3><span class="small muted">${inr(today.reduce((s, t) => s + t[1], 0))}</span></div>
    ${today.map(t => `<div class="kv"><span>${t[0]}</span><span class="num">${inr(t[1])}</span></div>`).join('')}</section>` : ''}
  <section class="sec"><div class="sec-h"><h3>From your Excel plan</h3></div>
    <div class="figs"><div class="fig"><span>Net worth at ${ageR}</span><b class="num">${cr(k.nw_at_60)}</b><em>${cr(k.nw_at_60_today)} in today’s money</em></div>
    <div class="fig"><span>Status</span><b style="font-size:16px">${esc(String(k.retire_status || '').replace(/^[^A-Za-z]+/, ''))}</b></div></div></section>
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
    <p class="small muted" style="margin:10px 0 0">From NW Projection in Excel (car and house plans included).</p></section>
  <section class="sec"><div class="group"><button class="navrow" data-gotab="goals"><span class="nt">Goals<br><span class="small muted">${(state.docs.goals || []).length ? `${state.docs.goals.length} goal${state.docs.goals.length === 1 ? '' : 's'} · what-if` : 'Emergency fund, car, trip – with what-if'}</span></span><span class="chev">›</span></button><button class="navrow" data-gotab="cover"><span class="nt">Insurance cover check</span><span class="chev">›</span></button></div></section>`;
}

/* ---------------- Salary, Money, Investments (layout from the reference screens) ---------------- */
function sheetRowCells(sheetName, startsWith) {
  const sh = sheetList().find(s => s.name === sheetName); if (!sh) return null;
  const r = sh.rows.find(x => (x.c.find(t => t) || '').trim().toLowerCase().startsWith(startsWith.toLowerCase()));
  return r ? r.c.filter(t => t) : null;
}
const PAY_LABEL = { 'Basic': 'Basic pay', 'HRA': 'House rent allowance', 'Special Allowance': 'Special allowance', 'Statutory Bonus': 'Statutory bonus',
  'Leave Encashment': 'Leave encashment', 'Performance Incentive': 'Performance bonus', 'PF (Employee)': 'Provident fund (PF)',
  'Professional Tax': 'Professional tax', 'Income Tax': 'Income tax (TDS)', 'ESI': 'ESI', 'Labour Welfare Fund': 'Labour welfare fund', 'Meal Recovery': 'Meal recovery' };
function viewSalary() {
  const first = latestPayslip();
  if (!first) return `<div class="empty">No payslips yet. Excel sends them with your sheets; run AppLinkNow, then tap ⟳.</div>`;
  const idx = state.payIdx == null ? first.index : state.payIdx, ps = latestPayslip(idx), r = ps.row;
  const gross = num(r['Gross Earnings']), ded = num(r['Total Deductions']), net = num(r['Net Pay']);
  const earn = ['Basic', 'HRA', 'Special Allowance', 'Statutory Bonus', 'Leave Encashment', 'Performance Incentive'].filter(h => num(r[h]) > 0);
  const deds = ['PF (Employee)', 'Professional Tax', 'Income Tax', 'ESI', 'Labour Welfare Fund', 'Meal Recovery'].filter(h => num(r[h]) > 0);
  const line = (l, v, b) => `<div class="kv ${b ? 'kvb' : ''}"><span>${l}</span><span class="num">${inr(v)}</span></div>`;
  const fy = (sheetList().find(s => s.name === 'Salary') || { rows: [] }).rows.map(x => x.c.filter(t => t)).filter(c => /^FY \d{4}/.test(c[0] || '') && num(c[1]) > 0);
  return `<div class="pick-row"><span class="muted small">Payslip for</span>
      <select class="inp sel" id="pay-month">${ps.months.map((m, k) => `<option value="${k}" ${k === idx ? 'selected' : ''}>${esc(m)}</option>`).reverse().join('')}</select></div>
    <section class="hero">
      <p class="hero-k" style="margin-top:0">Net take-home pay</p>
      <h2 class="hero-v num">${inr(net)}</h2>
      <div class="meter"><i style="width:${gross ? (net / gross * 100).toFixed(1) : 0}%"></i></div>
      <div class="hero-tiles"><div><span>Gross</span><b class="num">${inr(gross)}</b></div><div><span>Deductions</span><b class="num">${inr(ded)}</b></div></div>
    </section>
    <section class="sec box"><div class="sec-h"><h3>Earnings</h3></div>
      ${earn.map(h => line(PAY_LABEL[h] || h, num(r[h]))).join('')}${line('Gross earnings', gross, true)}</section>
    <section class="sec box"><div class="sec-h"><h3>Deductions</h3></div>
      ${deds.map(h => line(PAY_LABEL[h] || h, num(r[h]))).join('')}${line('Total deductions', ded, true)}</section>
    <section class="sec box"><div class="sec-h"><h3>Details</h3></div>
      <div class="kv"><span class="muted">Employer</span><span>${esc(r.Employer || '')}</span></div>
      <div class="kv"><span class="muted">Role</span><span>${esc(r.Designation || '')}</span></div>
      <div class="kv"><span class="muted">Paid into</span><span>${esc(r.Bank || '')}</span></div>
      <div class="kv"><span class="muted">Employer PF</span><span class="num">${inr(num(r['PF (Employer)']))}</span></div></section>
    <button class="primary" id="pay-share">Share payslip summary</button>
    ${fy.length ? `<section class="sec box"><div class="sec-h"><h3>By financial year</h3></div>
      ${fy.slice().reverse().map(c => `<div class="kv"><span>${esc(c[0])}<br><span class="small muted">${esc(c[1])} payslips</span></span><span class="num" style="text-align:right">${esc(c[3])}<br><span class="small muted">net</span></span></div>`).join('')}</section>` : ''}`;
}
function bindSalary() {
  const sel = document.getElementById('pay-month'); if (!sel) return;
  sel.onchange = () => { state.payIdx = Number(sel.value); render(); };
  document.getElementById('pay-share').onclick = async () => {
    const r = latestPayslip(state.payIdx == null ? undefined : state.payIdx).row;
    const t = `Payslip ${r.Month} – ${r.Employer}\nGross ${inr(num(r['Gross Earnings']))}\nDeductions ${inr(num(r['Total Deductions']))}\nNet pay ${inr(num(r['Net Pay']))}`;
    try { if (navigator.share) await navigator.share({ title: `Payslip ${r.Month}`, text: t }); else { await navigator.clipboard.writeText(t); toast('Copied'); } } catch (e) {}
  };
}

function viewMoney() {
  const k = kpis(), m = pDate(k.month), mode = state.moneyMode || 'exp', ym = String(k.month || '').slice(0, 7);
  const seg = `<div class="seg seg2" id="money-mode"><button data-mm="inc" class="${mode === 'inc' ? 'on' : ''}">Income</button><button data-mm="exp" class="${mode === 'exp' ? 'on' : ''}">Expenses</button></div>`;
  const bars = (list, cls) => { const max = Math.max(1, ...list.map(x => x.v));
    return list.map(x => `<div class="brow"><div class="row"><span>${esc(x.name)}</span><b class="num">${inr(x.v)}</b></div>
      <div class="meter thin ${cls || ''}"><i style="width:${(x.v / max * 100).toFixed(1)}%"></i></div></div>`).join(''); };
  let body;
  if (mode === 'inc') {
    const rows = tbl('AppRecent').filter(t => t.Type === 'Income' && String(t.Date).startsWith(ym));
    const by = {}; rows.forEach(t => { by[t.Category] = (by[t.Category] || 0) + (n(t.Amount) || 0); });
    const list = Object.entries(by).map(([name, v]) => ({ name, v })).sort((x, y) => y.v - x.v);
    const tot = list.reduce((s, x) => s + x.v, 0);
    body = `<section class="hero"><p class="hero-k" style="margin-top:0">Income received${m ? ' in ' + MONF[m.getMonth()] : ''}</p>
        <h2 class="hero-v num">${inr(tot)}</h2><p class="hero-s" style="margin:0">${rows.length} credits, by the date they reached your bank</p></section>
      <section class="sec box"><div class="sec-h"><h3>Where it came from</h3></div>${list.length ? bars(list, 'inc') : '<p class="small muted">No income this month yet.</p>'}</section>`;
  } else {
    const net = netPay(), p = activePlan(), limit = net - (Number(p.savings) || 0), spent = n(k.spending_month) || 0, used = limit > 0 ? spent / limit : 0;
    const list = tbl('AppBudget').filter(r => r.Category && n(r.Spent) > 0).map(r => ({ name: r.Category, v: n(r.Spent) })).sort((x, y) => y.v - x.v);
    body = `<section class="hero"><div class="row"><p class="hero-k" style="margin:0">Total spent${m ? ' in ' + MONF[m.getMonth()] : ''}</p>
          ${isFinite(net) ? `<p class="hero-k" style="margin:0;text-align:right">Limit ${inr(limit)}</p>` : ''}</div>
        <h2 class="hero-v num">${inr(spent)}</h2>
        ${isFinite(net) ? `<div class="meter ${used > 1 ? 'over' : ''}"><i style="width:${Math.min(100, used * 100).toFixed(1)}%"></i></div>
        <p class="hero-s">${used > 1 ? `${inr(spent - limit)} over your net-pay limit (${Math.round(used * 100)}%)` : `${inr(limit - spent)} left, ${Math.round(used * 100)}% of your net-pay limit used`}</p>` : ''}</section>
      <section class="sec box"><div class="sec-h"><h3>Where your money went</h3><button data-gotab="budget">Plan</button></div>${bars(list)}</section>
      <section class="sec"><div class="group"><button class="navrow" data-gotab="review"><span class="nt">${(() => { const d = pDate(kpis().month); return d ? MONF[d.getMonth()] : 'Monthly'; })()} review<br><span class="small muted">What went over and three things to change</span></span><span class="chev">›</span></button></div></section>
      ${viewForecast()}
      ${viewTrends()}
      <section class="sec"><div class="group"><button class="navrow" data-gotab="loans"><span class="nt">Loans<br><span class="small muted">${inr(n(kpis().loans_owed))} still owed · prepay what-if</span></span><span class="chev">›</span></button></div></section>
      <section class="sec"><div class="group"><button class="navrow" data-gotab="bills"><span class="nt">Bills and subscriptions<br><span class="small muted">${(() => { const l = statuses('bills'), nx = l.filter(x => x.st.s === 'due').sort((a, b) => a.st.due - b.st.due)[0]; return l.length ? (nx ? `Next: ${esc(nx.it.name)}, ${esc(nx.st.txt.toLowerCase())}` : `${l.filter(x => x.st.s === 'paid').length} of ${l.length} paid`) : 'Rent, EMIs, LIC, subscriptions – with reminders'; })()}</span></span><span class="chev">›</span></button></div></section>
      <section class="sec box"><div class="sec-h"><h3>Last 12 months</h3></div>${monthlyChart(tbl('AppMonthly'))}</section>`;
  }
  return seg + body + `<button class="secondary" data-gotab="tx" style="margin-top:16px">View all transactions</button>
    <button class="fab" data-gotab="add"><span>＋</span> Add transaction</button>`;
}
function bindMoney() {
  document.querySelectorAll('[data-mm]').forEach(b => b.onclick = () => { state.moneyMode = b.dataset.mm; render(); });
}

function viewInvest() {
  const tot = sheetRowValue('Balances', 'Investments total');
  const sum = sheetRowCells('Savings Schemes', 'Total');          // Total | Invested | Est. value | Gain | Gain %
  const mf = sheetRowCells('Balances', 'Mutual funds');
  const mix = ['Safe', 'Pension', 'Market'].map(w => {
    const sh = sheetList().find(s => s.name === 'Balances'); if (!sh) return null;
    for (const r of sh.rows) { const i = r.c.findIndex(t => String(t).startsWith(w)); if (i >= 0) return { name: r.c[i], v: num(r.c[i + 1]), pc: r.c[i + 2] }; }
    return null; }).filter(Boolean);
  const mixTot = mix.reduce((s, x) => s + x.v, 0) || 1;
  const MIXC = ['#a9cdf5', '#3f74b3', '#6fd3a0'];
  const hold = ['EPF', 'PPF', 'Fixed deposits', 'Other schemes'].map(l => { const c = sheetRowCells('Savings Schemes', l); return c ? { name: c[0], inv: num(c[1]), v: num(c[2]), g: c[4] } : null; })
    .filter(x => x && (x.v > 0 || x.inv > 0));
  if (mf) hold.push({ name: 'Mutual funds and stocks', v: num(mf[mf.length > 3 ? 3 : 2]), g: '' });
  const k = kpis();
  return `<section class="hero">
      <p class="hero-k" style="margin-top:0">Portfolio value</p>
      <h2 class="hero-v num">${tot != null ? inr(tot) : '–'}</h2>
      ${sum ? `<p class="hero-s" style="margin:-6px 0 0">Schemes: invested ${inr(num(sum[1]))}, gain <b>${inr(num(sum[3]))}</b> (${esc(sum[4] || '')})</p>` : ''}
      <div class="hero-tiles"><div><span>Investing a month</span><b class="num">${inr(k.invest_actual_month)}</b></div><div><span>Plan says</span><b class="num">${inr(k.invest_plan_month)}</b></div></div>
    </section>
    ${mix.length ? `<section class="sec box"><div class="sec-h"><h3>Asset allocation</h3></div>
      <div class="stack">${mix.map((x, i) => `<i style="width:${(x.v / mixTot * 100).toFixed(1)}%;background:${MIXC[i]}"></i>`).join('')}</div>
      <div class="dlegend">${mix.map((x, i) => `<span><i style="background:${MIXC[i]}"></i>${esc(x.name.replace(/\s*\(.*$/, ''))} <em>${esc(x.pc || '')}</em></span>`).join('')}</div></section>` : ''}
    ${viewInvestFunds()}
    ${viewInvestChart()}
    ${viewInvestPf()}
    ${viewInvestFds()}
    <section class="sec"><div class="group">
      <button class="navrow" data-gotab="sips"><span class="nt">SIPs<br><span class="small muted">${(() => { const l = statuses('sips'); return l.length ? `${l.filter(x => x.st.s === 'paid').length} of ${l.length} paid this month${l.some(x => x.st.s === 'missed') ? ' · one missed' : ''}` : 'Track and get reminders'; })()}</span></span><span class="chev">›</span></button>
      <button class="navrow" data-gotab="tax"><span class="nt">Tax: 80C and capital gains</span><span class="chev">›</span></button>
      <button class="navrow" data-gotab="import"><span class="nt">Import holdings from Groww</span><span class="chev">›</span></button>
    </div></section>
    <section class="sec box"><div class="sec-h"><h3>All investments</h3></div>
      ${hold.map(h => `<div class="kv"><span>${esc(h.name)}${h.inv ? `<br><span class="small muted">Invested ${inr(h.inv)}</span>` : ''}</span>
        <span class="num" style="text-align:right">${inr(h.v)}${h.g ? `<br><span class="small ${num(h.g) >= 0 ? 'good' : 'bad'}">+${esc(h.g)}</span>` : ''}</span></div>`).join('')}</section>
    <button class="secondary" data-sheet-open="Stocks & MF">Funds and stocks in Excel</button>
    <button class="secondary" data-sheet-open="Savings Schemes">PF, PPF and FDs in Excel</button>`;
}

/* ---------------- Investments: holdings at latest prices, buy / sell from the phone ---------------- */
const INV_TYPES = ['Equity MF', 'Index fund', 'ELSS (tax saver)', 'Small / Mid cap MF', 'Hybrid MF', 'Debt MF', 'Stock', 'ETF', 'Gold'];
const INV_PLATFORMS = ['Groww', 'YES Bank SIP (auto-debit)', 'Paytm Money', 'Zerodha', 'INDmoney', 'Other'];
function inv() { return (state.data && state.data.invest) || null; }
function priceOf(code) {
  const c = String(code || '').trim().toUpperCase(); if (!c || !state.prices) return null;
  return (state.prices.prices || []).find(p => p.code === c) || null;
}
/* Holdings as Excel last sent them, plus buys/sells saved on the phone that Excel hasn't added yet,
   valued at the latest price from prices.json (units × price). */
function invHoldings() {
  const x = inv(), list = [];
  const find = nm => list.find(h => h.name.toLowerCase() === String(nm).trim().toLowerCase());
  for (const h of (x && x.holdings) || []) list.push({ name: h.name, code: h.code || '', type: h.type || '', platform: h.platform || '',
    units: Number(h.units) || 0, invested: Number(h.invested) || 0, xlValue: Number(h.value) || 0, pending: 0 });
  const inExcel = new Set(((x && x.log) || []).map(r => r[5]));
  for (const p of state.trades.filter(t => !inExcel.has(t.id)).slice().reverse()) {     // oldest first
    const t = p.entry; let h = find(t.name);
    if (!h) { h = { name: t.name, code: t.code || '', type: t.type || '', platform: t.platform || '', units: 0, invested: 0, xlValue: 0, pending: 0 }; list.push(h); }
    if (!h.code && t.code) h.code = t.code;
    if (t.side === 'Buy') { h.units += t.units; h.invested += t.amount; }
    else { h.invested -= h.units > 0 ? h.invested * Math.min(1, t.units / h.units) : 0; h.units -= t.units; }
    h.pending++;
  }
  for (const h of list) {
    const p = priceOf(h.code);
    h.price = p ? p.price : null; h.priceDate = p ? p.date : '';
    h.value = p && h.units > 0 ? h.units * p.price : (h.pending ? Math.max(0, h.invested) : h.xlValue);
    h.gain = h.value - h.invested;
  }
  return list.filter(h => h.units > 1e-6 || h.pending).sort((a, b) => b.value - a.value);
}
function tradeLog() {
  const x = inv(), inExcel = new Set(((x && x.log) || []).map(r => r[5]));
  const rows = state.trades.filter(t => !inExcel.has(t.id)).map(t => ({ date: t.entry.date, side: t.entry.side, name: t.entry.name, units: t.entry.units, amount: t.entry.amount, id: t.id, pending: true }));
  for (const r of (x && x.log) || []) rows.push({ date: r[0], side: r[1], name: r[2], units: r[3], amount: r[4], id: r[5], pending: false });
  return rows.sort((a, b) => String(b.date).localeCompare(String(a.date)));
}
function fUnits(u) { return (Math.round(u * 1000) / 1000).toLocaleString('en-IN', { maximumFractionDigits: 3 }); }

function viewInvestFunds() {
  const hs = invHoldings(), tot = hs.reduce((s, h) => s + h.value, 0), invd = hs.reduce((s, h) => s + h.invested, 0);
  const dates = hs.map(h => h.priceDate).filter(Boolean).sort();
  const log = tradeLog().slice(0, 6);
  return `<section class="sec box"><div class="sec-h"><h3>Funds and stocks</h3>${hs.length ? `<span class="small num">${inr(tot)}</span>` : ''}</div>
      ${hs.length ? hs.map(h => `<div class="kv hold"><span>${esc(h.name)}<br><span class="small muted">${fUnits(h.units)} units${h.price ? ` × ₹${h.price.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : ''}${h.pending ? ` · <span class="warn">${h.pending} waiting for Excel</span>` : ''}</span></span>
          <span class="num" style="text-align:right">${inr(h.value)}<br><span class="small ${h.gain >= 0 ? 'good' : 'bad'}">${h.gain >= 0 ? '+' : '−'}${inr(Math.abs(h.gain))}${h.invested > 0 ? ` (${(h.gain / h.invested * 100).toFixed(1)}%)` : ''}</span></span></div>`).join('')
        + `<p class="small muted" style="margin:8px 0 0">Invested ${inr(invd)}. ${dates.length ? `Prices from ${fDay(pDate(dates[dates.length - 1]))}.` : 'Add a price code to get the latest price.'}</p>`
        : `<p class="small muted" style="margin:4px 0 0">No buys or sells yet. Add your first one and it becomes a holding.</p>`}
      <button class="primary" data-gotab="trade" style="margin-top:12px">Add buy / sell</button>
    </section>
    ${log.length ? `<section class="sec box"><div class="sec-h"><h3>Recent buys and sells</h3></div>
      ${log.map(r => `<div class="kv"><span>${esc(r.name)}<br><span class="small muted">${esc(fDay(pDate(r.date)))} · ${esc(r.side)} ${fUnits(Number(r.units) || 0)} units${r.pending ? ' · waiting for Excel' : ''}</span></span>
        <span class="num" style="text-align:right">${r.side === 'Sell' ? '+' : ''}${inr(Number(r.amount) || 0)}${r.pending ? `<br><button class="linkbtn" data-trade-edit="${esc(r.id)}">Edit</button> · <button class="linkbtn" data-trade-del="${esc(r.id)}">Remove</button>` : ''}</span></div>`).join('')}</section>` : ''}`;
}
function viewInvestPf() {
  const x = inv() || {}, e = x.epf, p = x.ppf || {}, b = balPending(), pc = v => (Number(v) * 100).toFixed(2).replace(/\.?0+$/, '') + '%';
  const wait = b ? `<p class="small warn" style="margin:2px 0 8px">Waiting for Excel: EPF ${inr(b.values.epf)}, EPS ${inr(b.values.eps)} (as of ${esc(fDay(pDate(b.values.asOf)))}), PPF ${inr(b.values.ppf)}. Open the workbook to apply.</p>` : '';
  return `<section class="sec box"><div class="sec-h"><h3>PF and PPF</h3><span class="small">from your passbook</span></div>${wait}
      ${e ? `<div class="kv"><span>EPF + EPS<br><span class="small muted">Passbook as of ${e.asOf ? esc(fDay(pDate(e.asOf))) : '–'}</span></span><span class="num" style="text-align:right">${inr(e.value)}<br><span class="small muted">EPS ${inr(e.epsTotal)}</span></span></div>
      <div class="kv"><span>Into EPF each month<br><span class="small muted">You ${pc(e.yourRate)} = ${inr(e.youMonth)} · employer ${inr(e.employerMonth)} after EPS</span></span><span class="num" style="text-align:right">${inr(e.epfMonth)}<br><span class="small muted">EPS ${inr(e.epsMonth)}</span></span></div>
      <div class="kv"><span>PPF<br><span class="small muted">${p.updated ? 'Updated ' + esc(fDay(pDate(p.updated))) : 'Balance you typed'}</span></span><span class="num">${inr(p.balance)}</span></div>`
      : `<p class="small muted" style="margin:4px 0 0">Run the new AppLink macro once in Excel to see your PF and PPF here.</p>`}
      <button class="primary" data-gotab="pfedit" style="margin-top:12px">Update PF / PPF</button>
    </section>`;
}
/* Update PF / PPF from the phone: saved as settings/balances.json, Excel writes it into Savings Schemes */
function balPending() { const b = state.balSaved; return b && b.id && (!state.data || state.data.balancesApplied !== b.id) ? b : null; }
function pfMonthly(wage, you, er, cap, rate) {
  const emp = Math.round(wage * er), eps = Math.min(emp, Math.round(Math.min(wage, cap) * rate));
  return { you: Math.round(wage * you), eps, erEpf: emp - eps, epf: Math.round(wage * you) + emp - eps };
}
function viewPfEdit() {
  const x = inv() || {}, e = x.epf || {}, p = x.ppf || {}, s = (balPending() || {}).values || {};
  const f = state.pfEdit || (state.pfEdit = {
    epf: s.epf ?? e.epfPassbook ?? '', eps: s.eps ?? e.epsPassbook ?? '', asOf: s.asOf ?? e.asOf ?? isoToday(),
    ppf: s.ppf ?? p.balance ?? '', ppfDate: s.ppfDate ?? p.updated ?? isoToday(),
    wage: s.wage ?? e.wage ?? 25000, you: +(((s.yourRate ?? e.yourRate ?? 0.12) * 100).toFixed(2)), er: +(((s.employerRate ?? e.employerRate ?? 0.12) * 100).toFixed(2)) });
  const m = pfMonthly(Number(f.wage) || 0, (Number(f.you) || 0) / 100, (Number(f.er) || 0) / 100, Number(e.epsCap) || 25000, Number(e.epsRate) || 0.0833);
  const val = v => v === '' || v == null ? '' : String(Math.round(Number(v) * 100) / 100);
  return `<section class="card">
    <h2 style="margin-top:0">EPF passbook</h2>
    <p class="small muted" style="margin:0 0 4px">From UMANG or passbook.epfindia.gov.in. Overwrites the old figures – nothing is added as a new row.</p>
    <label class="f" for="pf-epf">EPF balance (₹)</label>
    <input class="inp num" id="pf-epf" inputmode="decimal" value="${esc(val(f.epf))}" placeholder="Employee + employer share, with interest">
    <label class="f" for="pf-eps">EPS pension balance (₹)</label>
    <input class="inp num" id="pf-eps" inputmode="decimal" value="${esc(val(f.eps))}" placeholder="Pension contribution total">
    <label class="f" for="pf-asof">Passbook balance as of</label>
    <input class="inp" id="pf-asof" type="date" value="${esc(f.asOf)}" max="${isoToday()}">
    <p class="small muted" style="margin:8px 0 0">Each full month after this date adds ${inr(m.epf)} to EPF and ${inr(m.eps)} to EPS.</p>
  </section>
  <section class="card">
    <h2 style="margin-top:0">PPF</h2>
    <label class="f" for="pf-ppf">PPF balance (₹)</label>
    <input class="inp num" id="pf-ppf" inputmode="decimal" value="${esc(val(f.ppf))}" placeholder="Balance in your passbook">
    <label class="f" for="pf-ppfd">Balance date</label>
    <input class="inp" id="pf-ppfd" type="date" value="${esc(f.ppfDate || '')}" max="${isoToday()}">
  </section>
  <section class="card">
    <h2 style="margin-top:0">PF % and wage</h2>
    <p class="small muted" style="margin:0 0 4px">Change these only when your salary or PF % changes.</p>
    <label class="f" for="pf-wage">PF wage per month (Basic + DA, ₹)</label>
    <input class="inp num" id="pf-wage" inputmode="decimal" value="${esc(val(f.wage))}">
    <div class="row2">
      <div><label class="f" for="pf-you">Your PF %</label><input class="inp num" id="pf-you" inputmode="decimal" value="${esc(f.you)}"></div>
      <div><label class="f" for="pf-er">Employer PF %</label><input class="inp num" id="pf-er" inputmode="decimal" value="${esc(f.er)}"></div>
    </div>
    <div class="kv" style="margin-top:10px"><span>You → EPF</span><span class="num" id="pf-m-you">${inr(m.you)}</span></div>
    <div class="kv"><span>Employer → EPS (pension, max ₹2,083)</span><span class="num" id="pf-m-eps">${inr(m.eps)}</span></div>
    <div class="kv"><span>Employer → EPF</span><span class="num" id="pf-m-er">${inr(m.erEpf)}</span></div>
    <div class="kv"><span><b>Into EPF each month</b></span><span class="num"><b id="pf-m-epf">${inr(m.epf)}</b></span></div>
  </section>
  <button class="primary" id="pf-save">Save PF and PPF</button>
  <button class="secondary" id="pf-cancel">Cancel</button>
  <p class="note">Saved on GitHub. Excel writes these into Savings Schemes (blue cells) next time it opens – then save the workbook.</p>`;
}
function bindPfEdit() {
  const f = state.pfEdit, $ = id => document.getElementById(id), e = (inv() || {}).epf || {};
  const n = v => { const t = String(v).replace(/[,₹\s%]/g, ''); return t === '' ? null : Number(t); };
  const read = () => { f.epf = $('pf-epf').value; f.eps = $('pf-eps').value; f.asOf = $('pf-asof').value; f.ppf = $('pf-ppf').value;
    f.ppfDate = $('pf-ppfd').value; f.wage = $('pf-wage').value; f.you = $('pf-you').value; f.er = $('pf-er').value; };
  const live = () => {
    read();
    const m = pfMonthly(n(f.wage) || 0, (n(f.you) || 0) / 100, (n(f.er) || 0) / 100, Number(e.epsCap) || 25000, Number(e.epsRate) || 0.0833);
    $('pf-m-you').textContent = inr(m.you); $('pf-m-eps').textContent = inr(m.eps); $('pf-m-er').textContent = inr(m.erEpf); $('pf-m-epf').textContent = inr(m.epf);
  };
  ['pf-wage', 'pf-you', 'pf-er'].forEach(id => $(id).oninput = live);
  $('pf-cancel').onclick = () => { state.pfEdit = null; state.tab = 'invest'; window.scrollTo(0, 0); render(); };
  $('pf-save').onclick = async () => {
    read();
    const v = { epf: n(f.epf), eps: n(f.eps), asOf: f.asOf, ppf: n(f.ppf), ppfDate: f.ppfDate, wage: n(f.wage), yourRate: n(f.you), employerRate: n(f.er) };
    for (const [k, label] of [['epf', 'EPF balance'], ['eps', 'EPS balance'], ['ppf', 'PPF balance'], ['wage', 'PF wage']])
      if (v[k] == null || !(v[k] >= 0) || Number.isNaN(v[k])) { toast(`Enter the ${label}`); return; }
    if (!(v.yourRate > 0 && v.yourRate <= 50) || !(v.employerRate >= 0 && v.employerRate <= 50)) { toast('PF % should be between 0 and 50'); return; }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(v.asOf)) { toast('Pick the passbook date'); return; }
    v.yourRate = Math.round(v.yourRate * 100) / 10000; v.employerRate = Math.round(v.employerRate * 100) / 10000;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(v.ppfDate || '')) delete v.ppfDate;
    const saved = { id: 'b' + Date.now().toString(36), ...v, values: v, savedAt: new Date().toISOString() };
    const btn = $('pf-save'); btn.disabled = true; btn.textContent = 'Saving…';
    try {
      await GH.putJson('settings/balances.json', saved, 'PF and PPF balances from the app');
      state.balSaved = saved; LS.set('balSaved', saved); state.pfEdit = null; state.tab = 'invest';
      toast('Saved. Excel updates Savings Schemes next time it opens'); window.scrollTo(0, 0); render();
    } catch (err) { toast(err.message || 'Couldn’t reach GitHub – try again'); btn.disabled = false; btn.textContent = 'Save PF and PPF'; }
  };
}

/* Add buy / sell */
function newTrade() { return { side: 'Buy', kind: 'MF', name: '', code: '', type: 'Equity MF', units: '', price: '', amount: '', platform: 'Groww', date: isoToday(), note: '', results: [] }; }
function viewTrade() {
  const t = state.trade || (state.trade = newTrade());
  const mine = invHoldings();
  return `<section class="card">
    <div class="seg" id="tr-side">${['Buy', 'Sell'].map(s => `<button data-v="${s}" class="${t.side === s ? 'on' : ''}">${s}</button>`).join('')}</div>
    <div class="seg" id="tr-kind" style="margin-top:8px">${[['MF', 'Mutual fund'], ['Stock', 'Stock / ETF']].map(([k, l]) => `<button data-v="${k}" class="${t.kind === k ? 'on' : ''}">${l}</button>`).join('')}</div>
    ${mine.length ? `<label class="f">Your holdings</label><div class="pay" id="tr-mine">${mine.slice(0, 12).map((h, i) => `<button class="chip ${t.name === h.name ? 'on' : ''}" data-i="${i}">${esc(h.name.length > 28 ? h.name.slice(0, 27) + '…' : h.name)}</button>`).join('')}</div>` : ''}
    <label class="f" for="tr-name">${t.kind === 'MF' ? 'Fund name' : 'Stock name'}</label>
    <input class="inp" id="tr-name" autocomplete="off" placeholder="${t.kind === 'MF' ? 'Type 3+ letters to search, e.g. Parag Parikh' : 'e.g. Tata Motors'}" value="${esc(t.name)}">
    <div id="tr-results" class="group" ${t.results.length ? '' : 'hidden'}>${t.results.map((r, i) => `<button class="navrow" data-r="${i}"><span class="nt">${esc(r.schemeName)}<br><span class="small muted">Code ${esc(r.schemeCode)}</span></span></button>`).join('')}</div>
    <label class="f" for="tr-code">Price code</label>
    <input class="inp" id="tr-code" autocapitalize="characters" autocorrect="off" spellcheck="false" placeholder="${t.kind === 'MF' ? 'AMFI code – filled in when you pick a fund' : 'NSE symbol, e.g. TATAMOTORS'}" value="${esc(t.code)}">
    <div class="row2">
      <div><label class="f" for="tr-units">Units</label><input class="inp num" id="tr-units" inputmode="decimal" placeholder="0" value="${esc(t.units)}"></div>
      <div><label class="f" for="tr-price">Price per unit (₹)</label><input class="inp num" id="tr-price" inputmode="decimal" placeholder="NAV / price" value="${esc(t.price)}"></div>
    </div>
    <label class="f" for="tr-amt">Amount (₹)</label>
    <input class="inp amount num" id="tr-amt" inputmode="decimal" placeholder="0" value="${esc(t.amount)}">
    <div class="row2">
      <div><label class="f" for="tr-type">Type</label><select class="inp" id="tr-type">${INV_TYPES.map(x => `<option ${x === t.type ? 'selected' : ''}>${esc(x)}</option>`).join('')}</select></div>
      <div><label class="f" for="tr-plat">Platform</label><select class="inp" id="tr-plat">${INV_PLATFORMS.map(x => `<option ${x === t.platform ? 'selected' : ''}>${esc(x)}</option>`).join('')}</select></div>
    </div>
    <label class="f" for="tr-date">Date</label>
    <input class="inp" id="tr-date" type="date" value="${esc(t.date)}" max="${isoToday()}">
    <label class="f" for="tr-note">Note</label>
    <input class="inp" id="tr-note" maxlength="80" placeholder="e.g. Monthly SIP" value="${esc(t.note)}">
    <button class="primary" id="tr-save">${t.editId ? 'Save changes' : `Save ${t.side === 'Buy' ? 'buy' : 'sell'}`}</button>
    <p class="note">Amount is what you paid (buy) or got back (sell). Each one is added as a new row in Excel’s buy/sell log on Stocks &amp; MF the next time the workbook opens.</p>
  </section>`;
}
function bindTrade() {
  const t = state.trade, $ = id => document.getElementById(id);
  const read = () => { t.name = $('tr-name').value; t.code = $('tr-code').value.trim().toUpperCase(); t.units = $('tr-units').value; t.price = $('tr-price').value;
    t.amount = $('tr-amt').value; t.type = $('tr-type').value; t.platform = $('tr-plat').value; t.date = $('tr-date').value; t.note = $('tr-note').value; };
  const n = v => Number(String(v).replace(/[,₹\s]/g, ''));
  document.querySelectorAll('#tr-side button').forEach(b => b.onclick = () => { read(); t.side = b.dataset.v; render(); });
  document.querySelectorAll('#tr-kind button').forEach(b => b.onclick = () => { read(); t.kind = b.dataset.v; t.type = t.kind === 'MF' ? 'Equity MF' : 'Stock'; t.results = []; render(); });
  const mine = invHoldings();
  document.querySelectorAll('#tr-mine .chip').forEach(b => b.onclick = () => {
    read(); const h = mine[Number(b.dataset.i)];
    Object.assign(t, { name: h.name, code: h.code, type: h.type || t.type, platform: INV_PLATFORMS.includes(h.platform) ? h.platform : t.platform, results: [] });
    t.kind = /^\d+$/.test(h.code) || /MF|fund|ELSS/i.test(h.type) ? 'MF' : 'Stock';
    const p = priceOf(h.code); if (p && !t.price) t.price = String(p.price);
    render();
  });
  // units × price = amount
  $('tr-units').oninput = $('tr-price').oninput = () => { const u = n($('tr-units').value), p = n($('tr-price').value); if (u > 0 && p > 0) $('tr-amt').value = (Math.round(u * p * 100) / 100).toString(); };
  $('tr-amt').oninput = () => { const u = n($('tr-units').value), a = n($('tr-amt').value); if (u > 0 && a > 0) $('tr-price').value = (Math.round(a / u * 10000) / 10000).toString(); };
  // fund search (public mutual fund list)
  let timer;
  $('tr-name').oninput = () => {
    if (t.kind !== 'MF') return;
    clearTimeout(timer); const q = $('tr-name').value.trim();
    if (q.length < 3) { $('tr-results').hidden = true; return; }
    timer = setTimeout(async () => {
      try {
        const r = await fetch('https://api.mfapi.in/mf/search?q=' + encodeURIComponent(q));
        const arr = (await r.json()) || [];
        t.results = arr.filter(x => /direct/i.test(x.schemeName) === /direct/i.test(q) || !/direct|regular/i.test(q)).slice(0, 8);
        if (!t.results.length) t.results = arr.slice(0, 8);
        const box = $('tr-results'); if (!box) return;
        box.innerHTML = t.results.map((x, i) => `<button class="navrow" data-r="${i}"><span class="nt">${esc(x.schemeName)}<br><span class="small muted">Code ${esc(x.schemeCode)}</span></span></button>`).join('');
        box.hidden = !t.results.length; bindResults();
      } catch (e) { /* offline or search unavailable: type the code yourself */ }
    }, 350);
  };
  const bindResults = () => document.querySelectorAll('#tr-results [data-r]').forEach(b => b.onclick = () => {
    read(); const x = t.results[Number(b.dataset.r)]; t.name = x.schemeName; t.code = String(x.schemeCode); t.results = [];
    if (/elss|tax saver/i.test(x.schemeName)) t.type = 'ELSS (tax saver)'; else if (/index|nifty|sensex/i.test(x.schemeName)) t.type = 'Index fund';
    else if (/small|mid ?cap/i.test(x.schemeName)) t.type = 'Small / Mid cap MF'; else if (/hybrid|balanced/i.test(x.schemeName)) t.type = 'Hybrid MF';
    else if (/debt|liquid|gilt|bond/i.test(x.schemeName)) t.type = 'Debt MF';
    render();
  });
  bindResults();
  $('tr-save').onclick = async () => {
    read();
    const units = n(t.units), amount = n(t.amount);
    if (!t.name.trim()) { toast('Enter the fund or stock name'); $('tr-name').focus(); return; }
    if (!(units > 0)) { toast('Enter the units'); $('tr-units').focus(); return; }
    if (!(amount > 0)) { toast('Enter the amount'); $('tr-amt').focus(); return; }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(t.date)) { toast('Pick a date'); return; }
    if (t.side === 'Sell') {
      const h = invHoldings().find(x => x.name.toLowerCase() === t.name.trim().toLowerCase());
      if (h && t.editId) { const old = state.trades.find(x => x.id === t.editId); if (old && old.entry.side === 'Sell') h.units += old.entry.units; }
      if (!h) { toast('Sell: pick one of your holdings'); return; }
      if (units > h.units + 1e-6) { toast(`You have ${fUnits(h.units)} units – can’t sell more`); return; }
    }
    const now = new Date(), stamp = now.toISOString().replace(/[-:]/g, '').slice(0, 15);
    const entry = { id: t.editId || `${stamp}-${Math.random().toString(36).slice(2, 7)}`, date: t.date, side: t.side, name: t.name.trim().slice(0, 120),
      code: t.code.slice(0, 20), type: t.type, units: Math.round(units * 10000) / 10000, amount: Math.round(amount * 100) / 100,
      platform: t.platform, note: t.note.trim().slice(0, 80), createdAt: t.createdAt || now.toISOString() };
    const btn = $('tr-save'); btn.disabled = true; btn.textContent = 'Saving…';
    try {
      await GH.putJson(`trades/${entry.id}.json`, entry, `${entry.side} ${entry.name}`);
      state.trades = [{ id: entry.id, entry }, ...state.trades.filter(x => x.id !== entry.id)]; LS.set('trades', state.trades);
      toast(`Saved. Excel adds it to Stocks & MF next time it opens`);
      state.trade = null; state.tab = 'invest'; window.scrollTo(0, 0); render();
    } catch (e) {
      toast(e instanceof AppError ? e.message : 'Couldn’t reach GitHub – try again when you’re online');
      btn.disabled = false; btn.textContent = `Save ${t.side === 'Buy' ? 'buy' : 'sell'}`;
    }
  };
}
function bindInvest() {
  document.querySelectorAll('[data-trade-edit]').forEach(b => b.onclick = () => editTrade(b.dataset.tradeEdit));
  document.querySelectorAll('[data-trade-del]').forEach(b => b.onclick = async () => {
    const id = b.dataset.tradeDel; b.disabled = true;
    try {
      await GH.remove(`trades/${id}.json`, `Removed ${id}`);
      state.trades = state.trades.filter(t => t.id !== id); LS.set('trades', state.trades); toast('Removed'); render();
    } catch (e) { toast(e.message || 'Couldn’t reach GitHub'); b.disabled = false; }
  });
}
async function refreshTrades() {
  const r = await GH.call('GET', 'trades');
  if (r.status === 404) { state.trades = []; LS.set('trades', []); return; }
  if (!r.ok) return;
  const files = (await r.json()).filter(f => f.type === 'file' && /\.json$/.test(f.name));
  const known = Object.fromEntries(state.trades.map(t => [t.id, t])), out = [];
  for (const f of files) {
    const id = f.name.replace(/\.json$/, '');
    let t = known[id];
    if (!t) { try { t = { id, entry: await GH.getEntry(f.path) }; } catch (e) { continue; } }
    out.push(t);
  }
  out.sort((a, b) => String(b.entry.createdAt || '').localeCompare(String(a.entry.createdAt || '')));
  state.trades = out; LS.set('trades', out);
}

/* ======================================================================
   SIPs, bills & subscriptions, goals (app-owned files in the data repo: app/*.json)
   ====================================================================== */
const DOCS = { sips: 'app/sips.json', bills: 'app/bills.json', goals: 'app/goals.json', profile: 'app/profile.json', coaching: 'app/coaching.json', shared: 'app/shared.json', family: 'app/family.json' };
async function loadDocs() {
  for (const k of Object.keys(DOCS)) {
    try { const r = await GH.call('GET', DOCS[k], { raw: true }); if (r.ok) { const j = await r.json(); if (Array.isArray(j) || (j && typeof j === 'object')) state.docs[k] = j; } else if (r.status === 404 && state.docs[k] == null) state.docs[k] = /profile|coaching|family/.test(k) ? {} : []; }
    catch (e) { /* keep the copy on this phone */ }
  }
  LS.set('docs', state.docs);
}
async function saveDoc(k, msg) { await GH.putJson(DOCS[k], state.docs[k], msg || `Update ${k}`); LS.set('docs', state.docs); }
function uid(p) { return p + Date.now().toString(36) + Math.random().toString(36).slice(2, 5); }
const DAY = 864e5;
function dayOnly(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function dueIn(y, m, day) { const last = new Date(y, m + 1, 0).getDate(); return new Date(y, m, Math.min(Math.max(1, day), last)); }
function bankEnd() { const r = tbl('AppRecent'); return r.length ? pDate(r[0].Date) : null; }
/* Finds the bank transaction that paid an item: same amount (±tolerance), in the date window, keyword in the details */
function findPayment(item, from, to, types, tol, used) {
  const kw = String(item.keyword || '').trim().toLowerCase(), amt = Number(item.amount) || 0;
  return tbl('AppRecent').find(t => {
    if (used && used.has(t)) return false;
    const d = pDate(t.Date); if (!d || d < from || d > to) return false;
    if (!types.includes(t.Type)) return false;
    if (amt > 0 && Math.abs(Math.abs(n(t.Amount)) - amt) > Math.max(2, amt * tol)) return false;
    if (!(amt > 0) && !kw) return false;
    return !kw || `${t.Details} ${t.Category}`.toLowerCase().includes(kw);
  }) || null;
}
/* Status of one SIP or bill for the period that contains its next/current due date */
function itemStatus(item, kind) {
  const today = dayOnly(new Date()), end = bankEnd(), used = state._used || (state._used = new Set());
  let due;
  if (kind === 'bills' && item.every === 'Yearly') {
    const m = (Number(item.month) || 1) - 1;
    due = dueIn(today.getFullYear(), m, Number(item.day) || 1);
    if (due < new Date(today - 20 * DAY)) due = dueIn(today.getFullYear() + 1, m, Number(item.day) || 1);
  } else due = dueIn(today.getFullYear(), today.getMonth(), Number(item.day) || 1);
  const yearly = kind === 'bills' && item.every === 'Yearly';
  const from = yearly ? new Date(due - 20 * DAY) : new Date(due.getFullYear(), due.getMonth(), 1);
  const to = yearly ? new Date(+due + 20 * DAY) : new Date(due.getFullYear(), due.getMonth() + 1, 0);
  const types = kind === 'sips' ? ['Savings', 'Expenses', 'Transfer'] : ['Expenses', 'Savings', 'Transfer'];
  const hit = findPayment(item, from, to, types, kind === 'sips' ? 0.01 : 0.1, used);
  if (hit) { used.add(hit); return { s: 'paid', due, txt: `Paid ${fDay(pDate(hit.Date)).replace(/ \d{4}$/, '')}`, cls: 'good' }; }
  const days = Math.round((due - today) / DAY);
  if (days > 0) return { s: 'due', due, txt: days === 1 ? 'Due tomorrow' : `Due in ${days} days`, cls: days <= 3 ? 'warn' : 'muted' };
  if (days === 0) return { s: 'due', due, txt: 'Due today', cls: 'warn' };
  if (!end || end < due) return { s: 'wait', due, txt: `Due ${fDay(due).replace(/ \d{4}$/, '')} – bank data not synced yet`, cls: 'muted' };
  return { s: 'missed', due, txt: `Not found in bank data (due ${fDay(due).replace(/ \d{4}$/, '')})`, cls: 'bad' };
}
function statuses(kind) { state._used = new Set(); return (state.docs[kind] || []).map(it => ({ it, st: itemStatus(it, kind) })); }

const DOC_FIELDS = {
  sips: [['name', 'Fund name', 'text', 'e.g. Parag Parikh Flexi Cap'], ['amount', 'Amount each month (₹)', 'num'], ['day', 'Day of the month it runs', 'num', '1–28'],
    ['keyword', 'Text in the bank details', 'text', 'e.g. GROWW, ICCL or YESB00709 – leave empty to match by amount'], ['account', 'Paid from', 'acct']],
  bills: [['name', 'Name', 'text', 'e.g. Rent, Netflix, LIC, HDFC credit card'], ['amount', 'Amount (₹) – leave empty if it changes, like a card bill', 'num'], ['every', 'How often', 'sel', ['Monthly', 'Yearly']],
    ['day', 'Due day of the month', 'num', '1–31'], ['month', 'Due month (yearly bills)', 'selm'], ['keyword', 'Text in the bank details', 'text', 'e.g. NETFLIX, LIC, CRED, rent payee name'], ['account', 'Paid from', 'acct']],
  goals: [['name', 'Goal', 'text', 'e.g. Emergency fund, Car, Trip'], ['target', 'Target amount (₹)', 'num'], ['date', 'Target date', 'month'],
    ['saved', 'Saved so far (₹)', 'num'], ['monthly', 'Putting in each month (₹)', 'num'], ['ret', 'Expected return (% a year)', 'num', 'e.g. 7 for FD/debt, 12 for equity']],
};
const DOC_NAME = { sips: 'SIP', bills: 'bill', goals: 'goal' };
function viewDocForm(kind) {
  const e = state.docEdit, it = e.item;
  const fld = ([k, label, type, hint]) => {
    const v = it[k] ?? '';
    if (type === 'sel') return `<label class="f">${label}</label><div class="seg" data-sel="${k}">${hint.map(o => `<button data-v="${o}" class="${v === o ? 'on' : ''}">${o}</button>`).join('')}</div>`;
    if (type === 'acct') { const ac = bankAccounts().map(a => a.name); return ac.length ? `<label class="f" for="df-${k}">${label}</label><select class="inp" id="df-${k}"><option value="">Not set</option>${ac.map(a => `<option ${v === a ? 'selected' : ''}>${esc(a)}</option>`).join('')}</select>` : ''; }
    if (type === 'selm') return it.every === 'Yearly' ? `<label class="f" for="df-${k}">${label}</label><select class="inp" id="df-${k}">${MONF.map((m, i) => `<option value="${i + 1}" ${Number(v) === i + 1 ? 'selected' : ''}>${m}</option>`).join('')}</select>` : '';
    return `<label class="f" for="df-${k}">${label}</label><input class="inp ${type === 'num' ? 'num' : ''}" id="df-${k}" ${type === 'num' ? 'inputmode="decimal"' : ''} ${type === 'month' ? 'type="month"' : ''} value="${esc(v)}" placeholder="${esc(hint || '')}">`;
  };
  return `<section class="card"><h2 style="margin-top:0">${e.isNew ? 'Add' : 'Edit'} ${DOC_NAME[kind]}</h2>
    ${DOC_FIELDS[kind].map(fld).join('')}
    <button class="primary" id="df-save">Save</button>
    ${e.isNew ? '' : `<button class="secondary" id="df-del" style="color:var(--bad)">Delete</button>`}
    <button class="secondary" id="df-cancel">Cancel</button></section>`;
}
function bindDocForm(kind) {
  const e = state.docEdit, it = e.item, $ = id => document.getElementById(id);
  const read = () => DOC_FIELDS[kind].forEach(([k, , type]) => { const el = $('df-' + k); if (el) it[k] = type === 'num' ? el.value.replace(/[,₹\s%]/g, '') : el.value.trim(); });
  document.querySelectorAll('[data-sel] button').forEach(b => b.onclick = () => { read(); it[b.parentElement.dataset.sel] = b.dataset.v; render(); });
  $('df-cancel').onclick = () => { state.docEdit = null; render(); };
  const save = async (list, msg) => {
    const old = state.docs[kind]; state.docs[kind] = list;
    try { await saveDoc(kind, msg); state.docEdit = null; toast('Saved'); }
    catch (err) { state.docs[kind] = old; toast(err.message || 'Couldn’t reach GitHub – try again'); }
    render();
  };
  if ($('df-del')) $('df-del').onclick = () => { if (!confirmInline($('df-del'))) return; save(state.docs[kind].filter(x => x.id !== it.id), `Delete ${DOC_NAME[kind]} ${it.name}`); };
  $('df-save').onclick = () => {
    read();
    if (!it.name) { toast('Enter a name'); return; }
    for (const [k, label, type] of DOC_FIELDS[kind]) if (type === 'num' && it[k] !== '' && !(Number(it[k]) >= 0)) { toast(`Check ${label.toLowerCase()}`); return; }
    if (kind === 'sips' && !(Number(it.amount) > 0)) { toast('Enter the amount'); return; }
    if (kind === 'bills' && !(Number(it.amount) > 0) && !String(it.keyword || '').trim()) { toast('Enter the amount, or the bank text if the amount changes'); return; }
    if (kind !== 'goals' && !(Number(it.day) >= 1 && Number(it.day) <= 31)) { toast('Enter the day (1–31)'); return; }
    if (kind === 'goals' && !(Number(it.target) > 0)) { toast('Enter the target amount'); return; }
    if (kind === 'goals' && !/^\d{4}-\d{2}$/.test(it.date || '')) { toast('Pick the target date'); return; }
    const clean = { ...it }; DOC_FIELDS[kind].forEach(([k, , type]) => { if (type === 'num') clean[k] = clean[k] === '' ? null : Number(clean[k]); });
    if (kind === 'bills' && clean.every !== 'Yearly') delete clean.month;
    const list = e.isNew ? [...state.docs[kind], clean] : state.docs[kind].map(x => x.id === clean.id ? clean : x);
    save(list, `${e.isNew ? 'Add' : 'Update'} ${DOC_NAME[kind]} ${clean.name}`);
  };
}
function docAdd(kind, extra) { state.docEdit = { kind, isNew: true, item: { id: uid(kind[0]), ...(kind === 'bills' ? { every: 'Monthly' } : {}), ...(extra || {}) } }; render(); window.scrollTo(0, 0); }
function bindDocList(kind) {
  document.querySelectorAll('[data-doc-add]').forEach(b => b.onclick = () => docAdd(b.dataset.docAdd));
  document.querySelectorAll('[data-doc-edit]').forEach(b => b.onclick = () => {
    const [k, id] = b.dataset.docEdit.split(':'); const it = state.docs[k].find(x => x.id === id);
    if (it) { state.docEdit = { kind: k, isNew: false, item: { ...it } }; render(); window.scrollTo(0, 0); }
  });
}

/* SIPs */
function viewSips() {
  if (state.docEdit) return viewDocForm('sips');
  const list = statuses('sips'), tot = list.reduce((s, x) => s + (Number(x.it.amount) || 0), 0), paid = list.filter(x => x.st.s === 'paid').length;
  return `${viewLowBalance()}<section class="hero"><p class="hero-k" style="margin-top:0">SIPs each month</p><h2 class="hero-v num">${inr(tot)}</h2>
      <p class="hero-s" style="margin:0">${list.length ? `${paid} of ${list.length} paid this month` : 'Add your SIPs to track them'}${bankEnd() ? `. Bank data up to ${fDay(bankEnd())}` : ''}</p></section>
    <section class="sec box"><div class="sec-h"><h3>Your SIPs</h3><button data-doc-add="sips">＋ Add</button></div>
      ${list.length ? list.map(({ it, st }) => `<button class="kv rowbtn" data-doc-edit="sips:${esc(it.id)}"><span>${esc(it.name)}<br><span class="small ${st.cls}">${esc(st.txt)}</span></span><span class="num">${inr(it.amount)}<br><span class="small muted">on the ${ord(it.day)}</span></span></button>`).join('')
        : '<p class="small muted">Nothing yet. Tap ＋ Add for each SIP – you get a reminder the day before, and a warning if it doesn’t show up in your bank data.</p>'}
    </section><p class="note">Paid / missed is checked against your bank transactions: same amount, in the month, with the text you gave. Reminders need notifications turned on (More → Notifications).</p>`;
}
function ord(d) { d = Number(d) || 0; const s = ['th', 'st', 'nd', 'rd'], v = d % 100; return d + (s[(v - 20) % 10] || s[v] || s[0]); }

/* Bills & subscriptions */
function viewBills() {
  if (state.docEdit) return viewDocForm('bills');
  const list = statuses('bills').sort((a, b) => (a.st.s === 'paid') - (b.st.s === 'paid') || a.st.due - b.st.due);
  const monthly = (state.docs.bills || []).reduce((s, b) => s + (Number(b.amount) || 0) / (b.every === 'Yearly' ? 12 : 1), 0);
  const next = list.find(x => x.st.s === 'due');
  return `${viewLowBalance()}<section class="hero"><p class="hero-k" style="margin-top:0">Bills and subscriptions</p><h2 class="hero-v num">${inr(monthly)}<span style="font-size:16px;font-weight:500"> a month</span></h2>
      <p class="hero-s" style="margin:0">${next ? `Next: ${esc(next.it.name)} ${inr(next.it.amount)}, ${esc(next.st.txt.toLowerCase())}` : list.length ? 'Nothing due right now' : 'Add rent, EMIs, LIC, phone, OTT…'}</p></section>
    <section class="sec box"><div class="sec-h"><h3>All bills</h3><button data-doc-add="bills">＋ Add</button></div>
      ${list.length ? list.map(({ it, st }) => `<button class="kv rowbtn" data-doc-edit="bills:${esc(it.id)}"><span>${esc(it.name)}<br><span class="small ${st.cls}">${esc(st.txt)}</span></span><span class="num">${Number(it.amount) > 0 ? inr(it.amount) : 'varies'}<br><span class="small muted">${it.every === 'Yearly' ? `yearly, ${MON[(Number(it.month) || 1) - 1]} ${Number(it.day)}` : `monthly, ${ord(it.day)}`}</span></span></button>`).join('')
        : '<p class="small muted">Nothing yet. You get a reminder the day before each bill is due.</p>'}
    </section>`;
}

/* Goals with what-if */
function goalCalc(g, extra) {
  const r = (Number(g.ret) || 0) / 1200, saved = Number(g.saved) || 0, pm = (Number(g.monthly) || 0) + (extra || 0), target = Number(g.target) || 0;
  const [y, m] = String(g.date || '').split('-').map(Number), now = new Date();
  const months = Math.max(0, (y - now.getFullYear()) * 12 + (m - 1 - now.getMonth()));
  const fv = (n) => r ? saved * Math.pow(1 + r, n) + pm * (Math.pow(1 + r, n) - 1) / r : saved + pm * n;
  const atDate = fv(months);
  const need = months > 0 ? (r ? Math.max(0, (target - saved * Math.pow(1 + r, months)) * r / (Math.pow(1 + r, months) - 1)) : Math.max(0, (target - saved) / months)) : Math.max(0, target - saved);
  let reach = null;                                    // months until the target at this pace
  if (saved >= target) reach = 0; else if (pm > 0 || r > 0) { for (let k = 1; k <= 600; k++) if (fv(k) >= target) { reach = k; break; } }
  return { months, atDate, need, reach, pct: target ? Math.min(1, saved / target) : 0 };
}
function monthsLabel(k) { if (k == null) return 'never at this pace'; const d = new Date(); d.setMonth(d.getMonth() + k); return `${MON[d.getMonth()]} ${d.getFullYear()}`; }
function viewGoals() {
  if (state.docEdit) return viewDocForm('goals');
  const gs = state.docs.goals || [], extra = state.goalExtra || {};
  return `${gs.length ? '' : `<section class="hero"><p class="hero-k" style="margin-top:0">Goals</p><h2 class="hero-v" style="font-size:24px">Plan for what’s next</h2><p class="hero-s" style="margin:0">An emergency fund, a car, a trip – see how much to put in each month and when you’ll get there.</p></section>`}
    ${gs.map(g => { const c = goalCalc(g), x = extra[g.id] || 0, w = goalCalc(g, x), ok = c.atDate >= (Number(g.target) || 0);
      return `<section class="sec box"><div class="sec-h"><h3>${esc(g.name)}</h3><button data-doc-edit="goals:${esc(g.id)}">Edit</button></div>
        <div class="row"><span class="num"><b>${inr(g.saved)}</b> of ${inr(g.target)}</span><span class="small muted">by ${esc(fMon(pDate(g.date + '-01')))}</span></div>
        <div class="meter" style="margin:8px 0"><i style="width:${(c.pct * 100).toFixed(1)}%"></i></div>
        <p class="small ${ok ? 'good' : 'warn'}" style="margin:0">${ok ? `On track – about ${inr(c.atDate)} by then` : `Short by ${inr((Number(g.target) || 0) - c.atDate)} at ${inr(g.monthly || 0)} a month. Need ${inr(c.need)} a month`}</p>
        <label class="f" for="gx-${esc(g.id)}">What if I add more each month? <b class="num" id="gxv-${esc(g.id)}">+${inr(x)}</b></label>
        <input type="range" id="gx-${esc(g.id)}" data-gx="${esc(g.id)}" min="0" max="${Math.max(5000, Math.ceil((c.need || 10000) * 2 / 1000) * 1000)}" step="500" value="${x}">
        <p class="small muted" id="gxr-${esc(g.id)}" style="margin:4px 0 0">${goalWhatIf(g, x)}</p></section>`; }).join('')}
    <button class="primary" data-doc-add="goals">＋ Add a goal</button>
    <p class="note">Goals are kept in your data repo (app/goals.json). Update “saved so far” whenever you check.</p>`;
}
function goalWhatIf(g, x) { const w = goalCalc(g, x); return `At ${inr((Number(g.monthly) || 0) + x)} a month you reach ${inr(g.target)} in ${esc(monthsLabel(w.reach))}` + (x ? '' : ' – move the slider to try more'); }
function bindGoals() {
  document.querySelectorAll('[data-gx]').forEach(inp => inp.oninput = () => {
    const id = inp.dataset.gx, g = state.docs.goals.find(x => x.id === id), x = Number(inp.value);
    (state.goalExtra || (state.goalExtra = {}))[id] = x;
    document.getElementById('gxv-' + id).textContent = '+' + inr(x);
    document.getElementById('gxr-' + id).textContent = goalWhatIf(g, x);
  });
}

/* ======================================================================
   Edit pending phone entries
   ====================================================================== */
function editPending(id) {
  const p = state.pending.find(x => x.id === id); if (!p) return;
  const e = p.entry;
  state.add = { type: e.type, amount: String(e.amount), category: e.category, date: e.date, paidWith: e.paidWith || 'Cash', note: e.note || '', editId: id, createdAt: e.createdAt };
  state.tab = 'add'; window.scrollTo(0, 0); render();
}
function editTrade(id) {
  const t = state.trades.find(x => x.id === id); if (!t) return;
  const e = t.entry;
  state.trade = { side: e.side, kind: /^\d+$/.test(e.code || '') || /MF|fund|ELSS/i.test(e.type || '') ? 'MF' : 'Stock', name: e.name, code: e.code || '', type: e.type || 'Equity MF',
    units: String(e.units), price: String(Math.round(e.amount / e.units * 10000) / 10000), amount: String(e.amount), platform: e.platform || 'Groww', date: e.date, note: e.note || '', results: [], editId: id, createdAt: e.createdAt };
  state.tab = 'trade'; window.scrollTo(0, 0); render();
}
/* Quick add: the entries you add most often, newest first */
function quickPicks() {
  const seen = new Set(), out = [];
  const src = state.pending.map(p => p.entry).concat(allTx().filter(t => /app#/.test(t.Details || '')).map(t => ({ type: t.Type, category: t.Category, amount: Math.abs(n(t.Amount)),
    note: fullDetails(t).replace(/\s*\((Cash|UPI|Card|Other)\)\s*$/, ''), paidWith: (String(t.Details).match(/\((Cash|UPI|Card|Other)\)/) || [])[1] || 'Cash' })));
  for (const e of src) {
    const key = `${e.type}|${e.category}|${String(e.note || '').toLowerCase()}`;
    if (seen.has(key) || !e.category) continue; seen.add(key); out.push(e);
    if (out.length >= 6) break;
  }
  return out;
}

/* ======================================================================
   Investments extras: value history chart, FDs, tax view
   ====================================================================== */
function lineChart(pts) {
  if (pts.length < 2) return `<p class="small muted">The chart fills in day by day after the price check runs (about 10 pm).</p>`;
  const W = 320, H = 130, P = 6, xs = pts.map((_, i) => P + i * (W - 2 * P) / (pts.length - 1));
  const vals = pts.flatMap(p => [p.v, p.i]), lo = Math.min(...vals) * 0.98, hi = Math.max(...vals) * 1.02 || 1;
  const y = v => H - P - (v - lo) / (hi - lo || 1) * (H - 2 * P);
  const path = k => pts.map((p, i) => `${i ? 'L' : 'M'}${xs[i].toFixed(1)},${y(p[k]).toFixed(1)}`).join('');
  return `<svg viewBox="0 0 ${W} ${H}" width="100%" role="img" aria-label="Portfolio value over time">
      <path d="${path('i')}" fill="none" stroke="var(--slate)" stroke-width="1.6" stroke-dasharray="4 4"/>
      <path d="${path('v')}" fill="none" stroke="var(--blue1)" stroke-width="2.2"/></svg>
    <div class="dlegend"><span><i style="background:var(--blue1)"></i>Value</span><span><i style="background:var(--slate)"></i>Invested</span>
      <span class="muted">${esc(fDay(pDate(pts[0].d)))} – ${esc(fDay(pDate(pts[pts.length - 1].d)))}</span></div>`;
}
function viewInvestChart() {
  const h = (state.history && state.history.points) || [];
  if (!h.length && !invHoldings().length) return '';
  return `<section class="sec box"><div class="sec-h"><h3>Funds and stocks over time</h3></div>${lineChart(h.slice(-180))}</section>`;
}
function fdRows() {
  const sh = sheetList().find(s => s.name === 'Savings Schemes'); if (!sh) return [];
  const out = []; let on = false, hdr = null;
  for (const r of sh.rows) {
    const c0 = String(r.c[0] || '');
    if (/Fixed deposits \(bank/.test(c0)) { on = true; hdr = r.c; continue; }
    if (on && /^Total/.test(c0)) break;
    if (on && c0) { const o = {}; hdr.forEach((h, i) => { o[h] = r.c[i]; }); o.name = c0; out.push(o); }
  }
  return out;
}
function viewInvestFds() {
  const fds = fdRows().filter(f => f.Status === 'Active'); if (!fds.length) return '';
  return `<section class="sec box"><div class="sec-h"><h3>Fixed deposits</h3></div>
    ${fds.map(f => `<div class="kv"><span>${esc(f.name)}<br><span class="small ${Number(String(f['Days left']).replace(/,/g, '')) <= 30 ? 'warn' : 'muted'}">Matures ${esc(f['Maturity date'] || '')} · ${esc(f['Days left'] || '')} days left</span></span>
      <span class="num" style="text-align:right">${esc(f['Value now'] || '')}<br><span class="small muted">at maturity ${esc(f['Maturity value'] || '')}</span></span></div>`).join('')}</section>`;
}
function fyStart(d) { d = d || new Date(); return new Date(d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1, 3, 1); }
function capitalGains(from, to) {
  const types = Object.fromEntries(invHoldings().map(h => [h.name.toLowerCase(), h.type]).concat(((inv() || {}).holdings || []).map(h => [String(h.name).toLowerCase(), h.type])));
  const pools = {}, sells = [];
  for (const r of tradeLog().slice().reverse()) {                   // oldest first
    const k = String(r.name).toLowerCase(), p = pools[k] || (pools[k] = { units: 0, cost: 0, first: r.date });
    const u = Number(r.units) || 0, a = Number(r.amount) || 0, d = pDate(r.date);
    if (r.side === 'Buy') { if (p.units <= 1e-9) p.first = r.date; p.units += u; p.cost += a; }
    else if (p.units > 0) {
      const cost = p.cost * Math.min(1, u / p.units); p.cost -= cost; p.units -= u;
      if (d >= from && d <= to) {
        const held = (d - pDate(p.first)) / DAY, debt = /Debt/i.test(types[k] || '');
        sells.push({ name: r.name, date: r.date, gain: a - cost, kind: debt ? 'slab' : held > 365 ? 'LT' : 'ST' });
      }
    }
  }
  return sells;
}
function viewTax() {
  const from = fyStart(), to = new Date(from.getFullYear() + 1, 2, 31), fy = `${from.getFullYear()}–${String(from.getFullYear() + 1).slice(2)}`;
  const inFy = t => { const d = pDate(t.Date || t.date); return d && d >= from && d <= to; };
  const e = (inv() || {}).epf || {}, now = new Date();
  const monthsDone = Math.max(0, Math.min(12, (now.getFullYear() - from.getFullYear()) * 12 + now.getMonth() - 3));
  const epf = (Number(e.youMonth) || 0) * monthsDone;
  const ppfC = sheetRowCells('Savings Schemes', 'Deposited this financial year'), ppf = ppfC ? num(ppfC[1]) : 0;
  const elssNames = new Set(invHoldings().filter(h => /ELSS/i.test(h.type)).map(h => h.name.toLowerCase()));
  const elss = tradeLog().filter(r => r.side === 'Buy' && elssNames.has(String(r.name).toLowerCase()) && inFy(r)).reduce((s, r) => s + (Number(r.amount) || 0), 0);
  const ins = tbl('AppRecent').filter(t => t.Type === 'Expenses' && /insurance|lic/i.test(t.Category) && /\blic\b|life/i.test(t.Details) && inFy(t)).reduce((s, t) => s + Math.abs(n(t.Amount)), 0);
  const items = [['EPF – your share (estimated)', epf], ['PPF deposits', ppf], ['ELSS buys', elss], ['Life insurance premiums (LIC)', ins]];
  const used = items.reduce((s, x) => s + x[1], 0), left = Math.max(0, 150000 - used);
  const cg = capitalGains(from, to), st = cg.filter(x => x.kind === 'ST').reduce((s, x) => s + x.gain, 0), lt = cg.filter(x => x.kind === 'LT').reduce((s, x) => s + x.gain, 0), sl = cg.filter(x => x.kind === 'slab').reduce((s, x) => s + x.gain, 0);
  return `<section class="hero"><p class="hero-k" style="margin-top:0">Section 80C, FY ${fy}</p><h2 class="hero-v num">${inr(Math.min(used, 150000))}<span style="font-size:16px;font-weight:500"> of ₹1,50,000</span></h2>
      <div class="meter"><i style="width:${Math.min(100, used / 1500).toFixed(1)}%"></i></div>
      <p class="hero-s">${left ? `${inr(left)} left to invest by 31 Mar` : 'Limit reached'}</p></section>
    <section class="sec box"><div class="sec-h"><h3>What counts so far</h3></div>${items.map(([l, v]) => `<div class="kv"><span>${l}</span><span class="num">${inr(v)}</span></div>`).join('')}
      <p class="small muted" style="margin:8px 0 0">80C only helps in the old tax regime. Home-loan principal and tuition fees also count but aren’t tracked here.</p></section>
    <section class="sec box"><div class="sec-h"><h3>Capital gains from sells, FY ${fy}</h3></div>
      ${cg.length ? `<div class="kv"><span>Short-term (equity, held ≤ 1 year)<br><span class="small muted">taxed at 20%</span></span><span class="num ${st >= 0 ? '' : 'bad'}">${inr(st)}</span></div>
      <div class="kv"><span>Long-term (equity, held > 1 year)<br><span class="small muted">12.5% above ₹1.25 lakh a year</span></span><span class="num ${lt >= 0 ? '' : 'bad'}">${inr(lt)}</span></div>
      ${sl ? `<div class="kv"><span>Debt funds<br><span class="small muted">added to income, taxed at your slab</span></span><span class="num">${inr(sl)}</span></div>` : ''}
      ${cg.map(x => `<div class="kv small"><span class="muted">${esc(fDay(pDate(x.date)))} · ${esc(x.name)}</span><span class="num ${x.gain >= 0 ? 'good' : 'bad'}">${inr(x.gain)}</span></div>`).join('')}`
        : '<p class="small muted">No sells this financial year.</p>'}
      <p class="small muted" style="margin:8px 0 0">Estimate using average cost from your buy/sell log. Not tax advice – check with your tax filing or a CA.</p></section>
    ${(() => { const hl = tbl('AppRecent').filter(t => t.Type === 'Expenses' && /insurance|medical/i.test(t.Category) && /health|mediclaim|star|niva|care|hdfc ergo|icici lombard|bajaj allianz|tata aig/i.test(t.Details) && inFy(t)).reduce((s, t) => s + Math.abs(n(t.Amount)), 0);
      const intr = tbl('AppRecent').filter(t => t.Type === 'Income' && /interest/i.test(t.Category) && inFy(t)).reduce((s, t) => s + Math.abs(n(t.Amount)), 0);
      const div = tbl('AppRecent').filter(t => t.Type === 'Income' && /dividend/i.test(t.Category) && inFy(t)).reduce((s, t) => s + Math.abs(n(t.Amount)), 0);
      return `<section class="sec box"><div class="sec-h"><h3>Other items for your return</h3></div>
      <div class="kv"><span>80D – health insurance premiums<br><span class="small muted">up to ₹25,000 (₹50,000 for senior-citizen parents)</span></span><span class="num">${inr(hl)}</span></div>
      <div class="kv"><span>Savings bank interest<br><span class="small muted">80TTA: first ₹10,000 tax-free (old regime)</span></span><span class="num">${inr(intr)}</span></div>
      <div class="kv"><span>Dividends<br><span class="small muted">taxable at your slab</span></span><span class="num">${inr(div)}</span></div>
      <div class="kv"><span>EPF + EPS value<br><span class="small muted">for your records</span></span><span class="num">${inr(((inv() || {}).epf || {}).value)}</span></div>
      <div class="kv"><span>Which regime</span><span class="small">${esc(kpis().tax_better || '–')}</span></div></section>`; })()}
    <button class="secondary noprint" onclick="window.print()">Save tax summary as PDF</button>`;
}

/* ======================================================================
   Category trends (Money → Expenses)
   ====================================================================== */
function categoryTrends() {
  const k = kpis(), ym = String(k.month || '').slice(0, 7); if (!ym) return [];
  const [y, m] = ym.split('-').map(Number), prev = [1, 2, 3].map(i => { const d = new Date(y, m - 1 - i, 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; });
  const by = {};
  for (const t of tbl('AppRecent')) {
    if (t.Type !== 'Expenses') continue;
    const mo = String(t.Date).slice(0, 7), c = t.Category || 'Other'; if (mo !== ym && !prev.includes(mo)) continue;
    const o = by[c] || (by[c] = { now: 0, before: 0 }); if (mo === ym) o.now += Math.abs(n(t.Amount)); else o.before += Math.abs(n(t.Amount)) / 3;
  }
  const frac = monthProgress().done || 1;
  return Object.entries(by).map(([name, o]) => ({ name, now: o.now, avg: o.before, diff: o.now - o.before * frac }))
    .filter(x => x.avg > 200 || x.now > 500).sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff)).slice(0, 6);
}
function viewTrends() {
  const list = categoryTrends(); if (!list.length) return '';
  return `<section class="sec box"><div class="sec-h"><h3>Compared with your usual</h3><span class="small">3-month average</span></div>
    ${list.map(x => `<div class="kv"><span>${esc(x.name)}<br><span class="small muted">usually ${inr(x.avg)} a month</span></span>
      <span class="num" style="text-align:right">${inr(x.now)}<br><span class="small ${x.diff > 0 ? 'bad' : 'good'}">${x.diff > 0 ? '▲' : '▼'} ${inr(Math.abs(x.diff))} ${x.diff > 0 ? 'more' : 'less'} so far</span></span></div>`).join('')}</section>`;
}

/* ======================================================================
   Monthly report (print → Save as PDF)
   ====================================================================== */
function viewReport() {
  const months = tbl('AppMonthly').map(r => String(r.Month).slice(0, 7)).filter(Boolean);
  const mk = state.reportMonth && months.includes(state.reportMonth) ? state.reportMonth : months[months.length - 1] || String(kpis().month || '').slice(0, 7);
  const row = tbl('AppMonthly').find(r => String(r.Month).startsWith(mk)) || {};
  const tx = tbl('AppRecent').filter(t => String(t.Date).startsWith(mk));
  const cats = {}; tx.filter(t => t.Type === 'Expenses').forEach(t => { cats[t.Category] = (cats[t.Category] || 0) + Math.abs(n(t.Amount)); });
  const top = Object.entries(cats).sort((a, b) => b[1] - a[1]);
  const inc = {}; tx.filter(t => t.Type === 'Income').forEach(t => { inc[t.Category] = (inc[t.Category] || 0) + Math.abs(n(t.Amount)); });
  const k = kpis(), d = pDate(mk + '-01'), hs = invHoldings();
  return `<div class="noprint"><label class="f" for="rp-m">Month</label><select class="inp" id="rp-m">${months.slice().reverse().map(m => `<option value="${m}" ${m === mk ? 'selected' : ''}>${fMon(pDate(m + '-01'))}</option>`).join('')}</select>
      <button class="primary" id="rp-print" style="margin-top:12px">Save as PDF / print</button></div>
    <article class="report">
      <h1>MyMoney – ${d ? MONF[d.getMonth()] + ' ' + d.getFullYear() : ''}</h1><p class="small muted">${esc(userName())} · made ${fDay(new Date())}</p>
      <table class="rt"><tr><th>Income</th><th>Spent</th><th>Saved / invested</th><th>Left over</th></tr>
        <tr class="num"><td>${inr(row.Income)}</td><td>${inr(row.Spending)}</td><td>${inr(row.Saved)}</td><td>${inr(row.LeftOver)}</td></tr></table>
      <h2>Where the money went</h2><table class="rt">${top.map(([c, v]) => `<tr><td>${esc(c)}</td><td class="num">${inr(v)}</td></tr>`).join('') || '<tr><td>No spending</td></tr>'}</table>
      <h2>Income</h2><table class="rt">${Object.entries(inc).sort((a, b) => b[1] - a[1]).map(([c, v]) => `<tr><td>${esc(c)}</td><td class="num">${inr(v)}</td></tr>`).join('') || '<tr><td>No income</td></tr>'}</table>
      <h2>Net worth and investments (now)</h2><table class="rt"><tr><td>Net worth</td><td class="num">${inr(k.net_worth)}</td></tr>
        ${hs.map(h => `<tr><td>${esc(h.name)}</td><td class="num">${inr(h.value)}</td></tr>`).join('')}
        ${(inv() || {}).epf ? `<tr><td>EPF + EPS</td><td class="num">${inr(inv().epf.value)}</td></tr><tr><td>PPF</td><td class="num">${inr((inv().ppf || {}).balance)}</td></tr>` : ''}</table>
    </article>`;
}
function bindReport() {
  const s = document.getElementById('rp-m'); if (s) s.onchange = () => { state.reportMonth = s.value; render(); };
  const p = document.getElementById('rp-print'); if (p) p.onclick = () => window.print();
}

/* ======================================================================
   Coaching earnings (receipts kept by receipts.html on this phone)
   ====================================================================== */
function receiptsList() { try { return JSON.parse(localStorage.getItem('crm_receipts_v1') || '[]') || []; } catch (e) { return []; } }
function viewCoaching() {
  const list = receiptsList(), inrOf = r => r.cur === '₹' ? Number(r.total) || 0 : Number(r.inr) || 0, monthOf = r => r.month || String(r.date || '').slice(0, 7);
  const now = new Date(), ms = [...Array(6)].map((_, i) => { const d = new Date(now.getFullYear(), now.getMonth() - 5 + i, 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; });
  const byM = ms.map(m => ({ m, v: list.filter(r => monthOf(r) === m).reduce((s, r) => s + inrOf(r), 0) }));
  const max = Math.max(1, ...byM.map(x => x.v)), thisM = ms[5], lastM = ms[4];
  const fyFrom = fyStart(), byC = {};
  list.filter(r => pDate(r.date) >= fyFrom).forEach(r => { byC[r.client] = (byC[r.client] || 0) + inrOf(r); });
  const paidNow = new Set(list.filter(r => monthOf(r) === thisM).map(r => r.client));
  const notYet = [...new Set(list.filter(r => monthOf(r) === lastM).map(r => r.client))].filter(c => !paidNow.has(c));
  const sent = new Set(LS.get('rcSent', []));
  const cash = list.filter(r => /cash/i.test(r.method || '') && !sent.has(r.id));
  return `<section class="hero"><p class="hero-k" style="margin-top:0">Coaching this month</p><h2 class="hero-v num">${inr(byM[5].v)}</h2>
      <p class="hero-s" style="margin:0">${list.filter(r => monthOf(r) === thisM).length} receipts · last month ${inr(byM[4].v)}</p></section>
    <section class="sec box"><div class="sec-h"><h3>Last 6 months</h3></div>
      ${byM.map(x => `<div class="brow"><div class="row"><span>${MON[Number(x.m.slice(5)) - 1]} ${x.m.slice(0, 4)}</span><b class="num">${inr(x.v)}</b></div><div class="meter thin inc"><i style="width:${(x.v / max * 100).toFixed(1)}%"></i></div></div>`).join('')}</section>
    ${notYet.length ? `<section class="sec box"><div class="sec-h"><h3>Paid last month, not yet this month</h3></div>${notYet.map(c => `<div class="kv"><span>${esc(c)}</span><span class="small warn">no receipt yet</span></div>`).join('')}</section>` : ''}
    <section class="sec box"><div class="sec-h"><h3>By client, this financial year</h3></div>
      ${Object.entries(byC).sort((a, b) => b[1] - a[1]).map(([c, v]) => `<div class="kv"><span>${esc(c)}</span><span class="num">${inr(v)}</span></div>`).join('') || '<p class="small muted">No receipts yet.</p>'}</section>
    <section class="sec box"><div class="sec-h"><h3>Add to MyMoney</h3></div>
      <p class="small muted" style="margin:0 0 8px">Payments into your bank already reach Excel through the bank sync, so only cash receipts are added here (as Income → Other Income).</p>
      ${cash.length ? `${cash.slice(0, 10).map(r => `<div class="kv"><span>${esc(r.client)}<br><span class="small muted">${esc(fDay(pDate(r.date)))} · cash</span></span><span class="num">${inr(inrOf(r))}</span></div>`).join('')}
        <button class="primary" id="rc-add">Add ${cash.length} cash receipt${cash.length === 1 ? '' : 's'} to MyMoney</button>` : '<p class="small muted">No cash receipts waiting.</p>'}
    </section>
    ${viewCoachingMoney()}
    <a class="secondary btnlink" href="receipts.html">Open receipts</a>`;
}
function bindCoaching() {
  const b = document.getElementById('rc-add'); if (!b) return;
  b.onclick = async () => {
    b.disabled = true; b.textContent = 'Adding…';
    const sent = new Set(LS.get('rcSent', [])), inrOf = r => r.cur === '₹' ? Number(r.total) || 0 : Number(r.inr) || 0;
    let ok = 0;
    for (const r of receiptsList().filter(r => /cash/i.test(r.method || '') && !sent.has(r.id))) {
      const now = new Date(), entry = { id: `${now.toISOString().replace(/[-:]/g, '').slice(0, 15)}-${Math.random().toString(36).slice(2, 7)}`, date: String(r.date).slice(0, 10), type: 'Income',
        category: 'Other Income', amount: Math.round(inrOf(r) * 100) / 100, paidWith: 'Cash', note: `Coaching – ${r.client}${r.no ? ' ' + r.no : ''}`.slice(0, 80), createdAt: now.toISOString() };
      if (!(entry.amount > 0) || !/^\d{4}-\d{2}-\d{2}$/.test(entry.date)) { sent.add(r.id); continue; }
      try { const sha = await GH.putEntry(entry); state.pending.unshift({ id: entry.id, sha, entry }); sent.add(r.id); ok++; }
      catch (e) { toast(e.message || 'Couldn’t reach GitHub'); break; }
    }
    LS.set('rcSent', [...sent]); LS.set('pending', state.pending);
    if (ok) toast(`${ok} receipt${ok === 1 ? '' : 's'} added. Excel adds them next time it opens`);
    render();
  };
}

/* ======================================================================
   Groww import: paste the holdings file (Groww → Stocks → Holdings → Download)
   ====================================================================== */
function parseHoldingsText(txt) {
  const lines = String(txt).split(/\r?\n/).map(l => l.trim()).filter(Boolean); if (!lines.length) return [];
  const sep = lines.some(l => l.includes('\t')) ? '\t' : ',';
  const split = l => { const out = []; let cur = '', q = false; for (const ch of l) { if (ch === '"') q = !q; else if (ch === sep && !q) { out.push(cur); cur = ''; } else cur += ch; } out.push(cur); return out.map(s => s.trim()); };
  const rows = lines.map(split);
  const hi = rows.findIndex(r => r.some(c => /quantity|qty|units/i.test(c)) && r.some(c => /name|stock|symbol|scheme/i.test(c)));
  if (hi < 0) return [];
  const h = rows[hi].map(c => c.toLowerCase()), col = re => h.findIndex(c => re.test(c));
  const iName = col(/stock name|scheme name|^name|company/), iSym = col(/symbol|ticker/), iQty = col(/quantity|qty|units/), iAvg = col(/average|avg/), iInv = col(/invested|buy value|cost/), iIsin = col(/isin/);
  const nm = s => Number(String(s || '').replace(/[,₹\s]/g, ''));
  return rows.slice(hi + 1).map(r => {
    const qty = nm(r[iQty]), avg = iAvg >= 0 ? nm(r[iAvg]) : 0, inv = iInv >= 0 ? nm(r[iInv]) : 0;
    return { name: (r[iName >= 0 ? iName : iSym] || '').trim(), code: iSym >= 0 ? String(r[iSym]).trim().toUpperCase() : '', isin: iIsin >= 0 ? r[iIsin] : '', units: qty, amount: inv > 0 ? inv : Math.round(qty * avg * 100) / 100 };
  }).filter(x => x.name && x.units > 0 && x.amount > 0);
}
function viewImport() {
  const im = state.imp || (state.imp = { text: '', rows: [], date: isoToday(), kind: 'Stock' });
  return `<section class="card"><h2 style="margin-top:0">Import holdings</h2>
    <p class="small muted" style="margin:0 0 6px">In Groww: Stocks (or Mutual funds) → Holdings → download the holdings file. Open it, copy the table, and paste it here. Each holding becomes one buy with your total units and amount invested.</p>
    <div class="seg" id="im-kind">${[['Stock', 'Stocks'], ['MF', 'Mutual funds']].map(([k, l]) => `<button data-v="${k}" class="${im.kind === k ? 'on' : ''}">${l}</button>`).join('')}</div>
    <label class="f" for="im-text">Paste here</label>
    <textarea class="inp" id="im-text" rows="6" placeholder="Stock Name, Symbol, Quantity, Average buy price…">${esc(im.text)}</textarea>
    <label class="f" for="im-date">Date to use for these buys</label>
    <input class="inp" id="im-date" type="date" value="${esc(im.date)}" max="${isoToday()}">
    <button class="secondary" id="im-read">Read the table</button>
    ${im.rows.length ? `<div class="group" style="margin-top:10px">${im.rows.map(r => `<div class="kv"><span>${esc(r.name)}<br><span class="small muted">${fUnits(r.units)} units${r.code ? ' · ' + esc(r.code) : ''}</span></span><span class="num">${inr(r.amount)}</span></div>`).join('')}</div>
      <p class="small muted">${im.kind === 'Stock' ? 'Price code is the symbol – check it matches the NSE symbol.' : 'For funds, add the AMFI price code afterwards (Edit on each buy) to get live prices.'}</p>
      <button class="primary" id="im-save">Add ${im.rows.length} holding${im.rows.length === 1 ? '' : 's'}</button>` : ''}
  </section>`;
}
function bindImport() {
  const im = state.imp, $ = id => document.getElementById(id);
  document.querySelectorAll('#im-kind button').forEach(b => b.onclick = () => { im.text = $('im-text').value; im.kind = b.dataset.v; render(); });
  $('im-read').onclick = () => { im.text = $('im-text').value; im.date = $('im-date').value; im.rows = parseHoldingsText(im.text); if (!im.rows.length) toast('Couldn’t find a table with name and quantity columns'); render(); };
  const s = $('im-save'); if (s) s.onclick = async () => {
    im.date = $('im-date').value; if (!/^\d{4}-\d{2}-\d{2}$/.test(im.date)) { toast('Pick a date'); return; }
    s.disabled = true; s.textContent = 'Adding…'; let ok = 0;
    for (const r of im.rows) {
      const now = new Date(), entry = { id: `${now.toISOString().replace(/[-:]/g, '').slice(0, 15)}-${Math.random().toString(36).slice(2, 7)}`, date: im.date, side: 'Buy',
        name: r.name.slice(0, 120), code: im.kind === 'Stock' ? r.code.slice(0, 20) : '', type: im.kind === 'Stock' ? 'Stock' : 'Equity MF', units: Math.round(r.units * 10000) / 10000,
        amount: r.amount, platform: 'Groww', note: 'Imported holding', createdAt: now.toISOString() };
      try { await GH.putJson(`trades/${entry.id}.json`, entry, `Import ${entry.name}`); state.trades.unshift({ id: entry.id, entry }); ok++; }
      catch (e) { toast(e.message || 'Couldn’t reach GitHub'); break; }
    }
    LS.set('trades', state.trades); state.imp = null; toast(`${ok} holding${ok === 1 ? '' : 's'} added`); state.tab = 'invest'; render();
  };
}

/* ======================================================================
   Lock screen: PIN (+ fingerprint / face where the phone supports it)
   ====================================================================== */
const LOCK_AFTER = 5 * 60e3;
async function sha256(s) { const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s)); return [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, '0')).join(''); }
function lockCfg() { return LS.get('lock', null); }
function showLock() {
  const c = lockCfg(); if (!c) return;
  let el = document.getElementById('lock');
  if (!el) { el = document.createElement('div'); el.id = 'lock'; document.body.appendChild(el); }
  state.locked = true; state.pinTry = '';
  el.innerHTML = `<div class="lk"><div class="lk-logo">₹</div><h2>MyMoney is locked</h2><p class="small muted" id="lk-msg">Enter your PIN</p>
    <div class="lk-dots" id="lk-dots">${'<i></i>'.repeat(c.len)}</div>
    <div class="lk-pad">${[1, 2, 3, 4, 5, 6, 7, 8, 9, c.bio ? 'bio' : '', 0, 'del'].map(k => k === '' ? '<span></span>' : `<button data-k="${k}" aria-label="${k === 'bio' ? 'Use fingerprint' : k === 'del' ? 'Delete' : k}">${k === 'bio' ? '◉' : k === 'del' ? '⌫' : k}</button>`).join('')}</div></div>`;
  el.hidden = false; document.body.classList.add('locked');
  el.querySelectorAll('[data-k]').forEach(b => b.onclick = () => lockKey(b.dataset.k));
  if (c.bio) setTimeout(bioUnlock, 300);
}
async function lockKey(k) {
  const c = lockCfg(); if (k === 'bio') return bioUnlock();
  if (k === 'del') state.pinTry = state.pinTry.slice(0, -1); else if (state.pinTry.length < c.len) state.pinTry += k;
  document.querySelectorAll('#lk-dots i').forEach((d, i) => d.classList.toggle('on', i < state.pinTry.length));
  if (state.pinTry.length === c.len) {
    if (await sha256(c.salt + state.pinTry) === c.hash) unlock();
    else { document.getElementById('lk-msg').textContent = 'Wrong PIN – try again'; state.pinTry = ''; setTimeout(() => document.querySelectorAll('#lk-dots i').forEach(d => d.classList.remove('on')), 250); }
  }
}
async function bioUnlock() {
  const c = lockCfg(); if (!c || !c.bio || !window.PublicKeyCredential) return;
  try {
    const id = Uint8Array.from(atob(c.bio), ch => ch.charCodeAt(0));
    await navigator.credentials.get({ publicKey: { challenge: crypto.getRandomValues(new Uint8Array(32)), allowCredentials: [{ type: 'public-key', id }], userVerification: 'required', timeout: 60000 } });
    unlock();
  } catch (e) { /* cancelled: PIN still works */ }
}
function unlock() { state.locked = false; LS.set('lastActive', Date.now()); const el = document.getElementById('lock'); if (el) el.hidden = true; document.body.classList.remove('locked'); }
async function setPin(pin) { const salt = crypto.getRandomValues(new Uint32Array(2)).join('-'); LS.set('lock', { salt, hash: await sha256(salt + pin), len: pin.length, bio: (lockCfg() || {}).bio || null }); }
async function addBio() {
  if (!window.PublicKeyCredential || !(await PublicKeyCredential.isUserVerifyingPlatformAuthenticatorAvailable().catch(() => false))) { toast('This phone doesn’t offer fingerprint or face unlock to web apps'); return false; }
  const cred = await navigator.credentials.create({ publicKey: { challenge: crypto.getRandomValues(new Uint8Array(32)), rp: { name: 'MyMoney' },
    user: { id: crypto.getRandomValues(new Uint8Array(16)), name: 'mymoney', displayName: 'MyMoney' }, pubKeyCredParams: [{ type: 'public-key', alg: -7 }, { type: 'public-key', alg: -257 }],
    authenticatorSelection: { authenticatorAttachment: 'platform', userVerification: 'required', residentKey: 'discouraged' }, timeout: 60000 } });
  const c = lockCfg(); c.bio = btoa(String.fromCharCode(...new Uint8Array(cred.rawId))); LS.set('lock', c); return true;
}
function viewLockSettings() {
  const c = lockCfg(), theme = LS.get('theme', 'dark');
  return `<section class="card"><h2 style="margin-top:0">App lock</h2>
      <p class="small muted" style="margin:0 0 8px">${c ? `On – asks for your ${c.len}-digit PIN when the app opens and after 5 minutes away${c.bio ? ', or your fingerprint / face' : ''}.` : 'Off. Set a PIN so nobody can open your finances on this phone.'}</p>
      <label class="f" for="lk-new">${c ? 'New PIN' : 'PIN (4–6 digits)'}</label><input class="inp num" id="lk-new" type="password" inputmode="numeric" maxlength="6" autocomplete="off">
      <label class="f" for="lk-new2">Type it again</label><input class="inp num" id="lk-new2" type="password" inputmode="numeric" maxlength="6" autocomplete="off">
      <button class="primary" id="lk-save">${c ? 'Change PIN' : 'Turn on app lock'}</button>
      ${c ? `${c.bio ? '<button class="secondary" id="lk-bio-off">Stop using fingerprint / face</button>' : '<button class="secondary" id="lk-bio">Also unlock with fingerprint / face</button>'}
        <button class="secondary" id="lk-off" style="color:var(--bad)">Turn off app lock</button>` : ''}
    </section>
    ${viewPhoneCheck()}
    <section class="card"><h2 style="margin-top:0">Look</h2><div class="seg" id="theme-seg">${[['dark', 'Dark'], ['light', 'Light'], ['auto', 'Match phone']].map(([k, l]) => `<button data-v="${k}" class="${theme === k ? 'on' : ''}">${l}</button>`).join('')}</div></section>`;
}
function bindLockSettings() {
  const $ = id => document.getElementById(id);
  document.querySelectorAll('#theme-seg button').forEach(b => b.onclick = () => { LS.set('theme', b.dataset.v); applyTheme(); render(); });
  $('lk-save').onclick = async () => {
    const a = $('lk-new').value, b = $('lk-new2').value;
    if (!/^\d{4,6}$/.test(a)) { toast('PIN must be 4 to 6 digits'); return; }
    if (a !== b) { toast('The two PINs don’t match'); return; }
    await setPin(a); toast('App lock on'); render();
  };
  if ($('lk-off')) $('lk-off').onclick = () => { if (!confirmInline($('lk-off'))) return; LS.del('lock'); toast('App lock off'); render(); };
  if ($('lk-bio')) $('lk-bio').onclick = async () => { try { if (await addBio()) toast('Fingerprint / face unlock on'); } catch (e) { toast('Not set up – ' + (e.name === 'NotAllowedError' ? 'cancelled' : 'not supported here')); } render(); };
  if ($('lk-bio-off')) $('lk-bio-off').onclick = () => { const c = lockCfg(); c.bio = null; LS.set('lock', c); render(); };
}
function applyTheme() {
  const t = LS.get('theme', 'dark'), light = t === 'light' || (t === 'auto' && window.matchMedia && matchMedia('(prefers-color-scheme: light)').matches);
  document.documentElement.dataset.theme = light ? 'light' : 'dark';
  const m = document.querySelector('meta[name=theme-color]'); if (m) m.content = light ? '#f4f6f9' : '#0d0f12';
}

/* ======================================================================
   Accounts and balances (Balances sheet), low balance before SIPs / bills
   ====================================================================== */
function bankAccounts() {
  const sh = sheetList().find(s => s.name === 'Balances'); if (!sh) return [];
  let hdr = null; const out = [];
  for (const r of sh.rows) {
    if (r.c[0] === 'Account') { hdr = r.c; continue; }
    if (hdr && /total/i.test(r.c[0])) break;
    if (hdr && r.c[0] && /Bank|Cash/.test(r.c[1] || '')) { const used = r.c[hdr.indexOf('Balance used')]; out.push({ name: r.c[0], type: r.c[1], bal: num(used) || 0, asOf: r.c[hdr.indexOf('As of')] || '' }); }
  }
  return out;
}
/* Upcoming unpaid SIPs and bills (next 7 days) grouped by the account they're paid from */
function lowBalance() {
  const accts = Object.fromEntries(bankAccounts().map(a => [a.name, a])), due = {};
  for (const kind of ['sips', 'bills']) for (const { it, st } of statuses(kind)) {
    if (st.s !== 'due' || !it.account || !accts[it.account]) continue;
    if ((st.due - dayOnly(new Date())) / DAY > 7) continue;
    (due[it.account] || (due[it.account] = [])).push({ it, st });
  }
  return Object.entries(due).map(([a, list]) => ({ account: a, bal: accts[a].bal, need: list.reduce((s, x) => s + (Number(x.it.amount) || 0), 0), list }))
    .filter(x => x.need > x.bal);
}
function viewLowBalance() {
  const lb = lowBalance(); if (!lb.length) return '';
  return lb.map(x => `<section class="sec box warnbox"><div class="sec-h"><h3>Low balance in ${esc(x.account)}</h3></div>
    <p class="small" style="margin:0">${x.list.map(y => `${esc(y.it.name)} ${inr(y.it.amount)} (${esc(y.st.txt.toLowerCase())})`).join(', ')} – needs ${inr(x.need)}, the account had ${inr(x.bal)} at the last statement.</p></section>`).join('');
}

/* ======================================================================
   Change a bank transaction's category (becomes a rule at the top of Auto Rules)
   ====================================================================== */
function txKey(t) { return `${String(t.Date).slice(0, 10)}|${n(t.Amount)}|${String(t.Details || '').slice(0, 60)}`; }
function recatFor(t) { const k = txKey(t); return state.recats.find(r => r.entry.tx === k) || null; }
function upiParts(det) {
  const d = String(det || '').replace(/^bank:\s*\{[^}]*\}\s*/i, '');
  const m = d.match(/UPI\/(?:DR|CR)\/(\d{6,})\/([^/]+)/i);
  if (m) return { ref: m[1], payee: m[2].trim() };
  const ref = (d.match(/\b\d{9,}\b/) || [])[0] || '';
  const words = d.replace(/[^A-Za-z ]+/g, ' ').trim().split(/\s+/).filter(w => w.length > 2 && !/^(UPI|NEFT|IMPS|RTGS|POS|ATM|TFR|DEP|WDL|PAYMENT|FROM|PHONEPE|PAYM|BRANCH|TRANSFER)$/i.test(w));
  return { ref, payee: words.slice(0, 2).join(' ') };
}
function openRecat(t) {
  const p = upiParts(t.Details);
  state.recat = { t, type: t.Type, category: t.Category, scope: p.payee ? 'payee' : 'one', payee: p.payee, ref: p.ref };
  state.tab = 'recat'; window.scrollTo(0, 0); render();
}
function viewRecat() {
  const r = state.recat; if (!r) { state.tab = 'tx'; return viewTx(); }
  const t = r.t, cats = catsFor(r.type);
  return `<section class="card"><p class="small muted" style="margin:0">${esc(fDay(pDate(t.Date)))} · ${esc(t.Source || '')}</p>
      <h2 style="margin:4px 0 2px">${esc(describe(t).title)} <span class="num">${inr(Math.abs(n(t.Amount)))}</span></h2>
      <p class="small muted" style="margin:0;overflow-wrap:anywhere">${esc(fullDetails(t))}</p>
      <p class="small" style="margin:8px 0 0">Now: ${esc(t.Type)} → ${esc(t.Category)}</p></section>
    <section class="card">
      <div class="seg" id="rc-type">${['Expenses', 'Income', 'Savings', 'Transfer'].map(x => `<button data-v="${x}" class="${r.type === x ? 'on' : ''}">${x === 'Expenses' ? 'Spent' : x}</button>`).join('')}</div>
      <label class="f" for="rc-cat">New category</label>
      <select class="inp" id="rc-cat">${cats.map(c => `<option ${c === r.category ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select>
      <label class="f">Change</label>
      <div class="seg" id="rc-scope"><button data-v="one" class="${r.scope === 'one' ? 'on' : ''}" ${r.ref ? '' : 'disabled'}>Just this one</button><button data-v="payee" class="${r.scope === 'payee' ? 'on' : ''}">Always for this payee</button></div>
      ${r.scope === 'payee' ? `<label class="f" for="rc-kw">Payee text to match</label><input class="inp" id="rc-kw" value="${esc(r.payee)}" autocapitalize="characters">
        <p class="small muted" style="margin:6px 0 0">Every bank row with this text gets the new category – past and future. Keep it specific (e.g. SWIGGY, not PAY).</p>`
        : `<p class="small muted" style="margin:8px 0 0">Matched by its reference number ${esc(r.ref)}.</p>`}
      <button class="primary" id="rc-save">Save</button><button class="secondary" id="rc-cancel">Cancel</button>
      <p class="note">Excel adds this as a rule at the top of Auto Rules next time it opens, then totals update.</p></section>`;
}
function bindRecat() {
  const r = state.recat, $ = id => document.getElementById(id); if (!r) return;
  const read = () => { r.category = $('rc-cat').value; if ($('rc-kw')) r.payee = $('rc-kw').value.trim(); };
  document.querySelectorAll('#rc-type button').forEach(b => b.onclick = () => { read(); r.type = b.dataset.v; r.category = ''; render(); });
  document.querySelectorAll('#rc-scope button').forEach(b => b.onclick = () => { read(); r.scope = b.dataset.v; render(); });
  $('rc-cancel').onclick = () => { state.recat = null; state.tab = 'tx'; render(); };
  $('rc-save').onclick = async () => {
    read();
    const kw = r.scope === 'one' ? r.ref : r.payee;
    if (!kw || kw.length < 3) { toast('Payee text must be at least 3 letters'); return; }
    if (!r.category) { toast('Pick a category'); return; }
    const now = new Date(), entry = { id: `${now.toISOString().replace(/[-:]/g, '').slice(0, 15)}-${Math.random().toString(36).slice(2, 7)}`,
      keyword: kw.slice(0, 60), direction: n(r.t.Amount) < 0 || r.t.Type !== 'Income' ? 'Debit' : 'Credit', type: r.type, category: r.category,
      scope: r.scope, tx: txKey(r.t), note: `${r.scope === 'one' ? 'One transaction' : 'Payee'} set from phone (was ${r.t.Category})`, createdAt: now.toISOString() };
    try {
      await GH.putJson(`recat/${entry.id}.json`, entry, `Category change ${entry.id}`);
      state.recats.unshift({ id: entry.id, entry }); LS.set('recats', state.recats);
      toast('Saved. Excel adds the rule next time it opens'); state.recat = null; state.tab = 'tx'; render();
    } catch (e) { toast(e.message || 'Couldn’t reach GitHub – try again'); }
  };
}
async function refreshRecats() {
  const r = await GH.call('GET', 'recat'); if (r.status === 404) { state.recats = []; LS.set('recats', []); return; } if (!r.ok) return;
  const files = (await r.json()).filter(f => f.type === 'file' && /\.json$/.test(f.name)), known = Object.fromEntries(state.recats.map(x => [x.id, x])), out = [];
  for (const f of files) { const id = f.name.replace(/\.json$/, ''); let x = known[id]; if (!x) { try { x = { id, entry: await GH.getEntry(f.path) }; } catch (e) { continue; } } out.push(x); }
  state.recats = out; LS.set('recats', out);
}
/* a pending payee rule changes how matching rows look until Excel applies it */
function applyRecatsView(rows) {
  if (!state.recats.length) return rows;
  const one = Object.fromEntries(state.recats.filter(x => x.entry.scope === 'one').map(x => [x.entry.tx, x.entry]));
  const payee = state.recats.filter(x => x.entry.scope !== 'one').map(x => x.entry);
  return rows.map(t => {
    if (t._pending) return t;
    const hit = one[txKey(t)] || payee.find(e => String(t.Details || '').toLowerCase().includes(String(e.keyword).toLowerCase()) && (e.direction === 'Credit') === (t.Type === 'Income'));
    return hit && (hit.category !== t.Category || hit.type !== t.Type) ? { ...t, Type: hit.type, Category: hit.category, _recat: true } : t;
  });
}

/* ======================================================================
   Share to MyMoney / paste an SMS: amount, payee and date filled in
   ====================================================================== */
function parseSms(txt) {
  const s = String(txt || '').replace(/\s+/g, ' ');
  const am = s.match(/(?:rs\.?|inr|₹)\s*([\d,]+(?:\.\d{1,2})?)/i) || s.match(/([\d,]+\.\d{2})\b/);
  const amount = am ? Number(am[1].replace(/,/g, '')) : null;
  const credit = /\b(credited|received|deposited|refund)\b/i.test(s) && !/\bdebited\b/i.test(s);
  let payee = (s.match(/\b(?:to|at|towards|paid to|trf to)\s+([A-Za-z][A-Za-z0-9 .&'-]{2,30}?)(?=\s+(?:on|via|ref|upi|for|from|avl|a\/c|\.|,)|[.,]|$)/i) || [])[1] || '';
  if (!payee) payee = (s.match(/VPA\s+([\w.-]+)@/i) || [])[1] || '';
  let date = '';
  const d1 = s.match(/\b(\d{1,2})[-/ ](\d{1,2}|[A-Za-z]{3})[-/ ](\d{2,4})\b/);
  if (d1) {
    const mo = /^\d+$/.test(d1[2]) ? Number(d1[2]) : MON.findIndex(m => m.toLowerCase() === d1[2].slice(0, 3).toLowerCase()) + 1;
    const yr = d1[3].length === 2 ? 2000 + Number(d1[3]) : Number(d1[3]);
    if (mo >= 1 && mo <= 12 && Number(d1[1]) <= 31) date = `${yr}-${String(mo).padStart(2, '0')}-${String(d1[1]).padStart(2, '0')}`;
  }
  if (date > isoToday()) date = '';
  return { amount, type: credit ? 'Income' : 'Expenses', note: payee.trim().slice(0, 60), date };
}
function applySms(txt) {
  const p = parseSms(txt), a = state.add || (state.add = newAdd());
  if (!p.amount) { toast('Couldn’t find an amount in that text'); return false; }
  Object.assign(a, { amount: String(p.amount), type: p.type, note: p.note || a.note, date: p.date || a.date, paidWith: 'UPI', category: '' });
  return true;
}
function handleShared() {
  const q = new URLSearchParams(location.search), txt = [q.get('title'), q.get('text'), q.get('url')].filter(Boolean).join(' ');
  if (!txt) return;
  history.replaceState(null, '', location.pathname);
  state.add = newAdd(); if (applySms(txt)) { state.tab = 'add'; toast('Filled in from what you shared – check and save'); }
}

/* ======================================================================
   Net worth over time, loans with prepayment what-if, month-end forecast
   ====================================================================== */
function viewNetWorthChart() {
  const pts = ((state.history && state.history.nw) || []).map(p => ({ d: p.d, v: p.v, i: p.v }));
  if (!pts.length) return '';
  if (pts.length < 2) return `<section class="sec box"><div class="sec-h"><h3>Net worth over time</h3></div><p class="small muted">Starts today at ${inr(pts[0].v)} – a new point is added each day Excel sends figures.</p></section>`;
  const W = 320, H = 120, P = 6, vs = pts.map(p => p.v), lo = Math.min(...vs) * .98, hi = Math.max(...vs) * 1.02, xs = pts.map((_, i) => P + i * (W - 2 * P) / (pts.length - 1));
  const y = v => H - P - (v - lo) / (hi - lo || 1) * (H - 2 * P), ch = vs[vs.length - 1] - vs[0];
  return `<section class="sec box"><div class="sec-h"><h3>Net worth over time</h3><span class="small ${ch >= 0 ? 'good' : 'bad'}">${ch >= 0 ? '+' : '−'}${inr(Math.abs(ch))}</span></div>
    <svg viewBox="0 0 ${W} ${H}" width="100%" role="img" aria-label="Net worth over time"><path d="${pts.map((p, i) => `${i ? 'L' : 'M'}${xs[i].toFixed(1)},${y(p.v).toFixed(1)}`).join('')}" fill="none" stroke="var(--blue1)" stroke-width="2.2"/></svg>
    <p class="small muted" style="margin:4px 0 0">${esc(fDay(pDate(pts[0].d)))} – ${esc(fDay(pDate(pts[pts.length - 1].d)))}</p></section>`;
}
function loanRows() {
  const sh = sheetList().find(s => s.name === 'Loans'); if (!sh) return [];
  let hdr = null; const out = [];
  for (const r of sh.rows) {
    if (/^Lender/.test(r.c[0] || '')) { hdr = r.c; continue; }
    if (!hdr) continue; if (/^Total/i.test(r.c[0] || '')) break;
    const o = {}; hdr.forEach((h, i) => { o[h] = r.c[i]; }); o.name = r.c[0]; out.push(o);
  }
  return out;
}
function loanMonths(bal, emi, r) { if (bal <= 0) return 0; if (!r) return Math.ceil(bal / emi); const x = 1 - r * bal / emi; return x <= 0 ? Infinity : Math.ceil(-Math.log(x) / Math.log(1 + r)); }
function viewLoans() {
  const rows = loanRows(), act = rows.filter(r => /Active/i.test(r.Status || '')), rates = LS.get('loanRates', {}), pre = state.prepay || {};
  return `<section class="hero"><p class="hero-k" style="margin-top:0">Still owed</p><h2 class="hero-v num">${inr(act.reduce((s, r) => s + num(r['Still owed']), 0))}</h2>
      <p class="hero-s" style="margin:0">${act.length} active loan${act.length === 1 ? '' : 's'} · ${inr(act.reduce((s, r) => s + num(r.EMI), 0))} EMI a month</p></section>
    ${act.map(r => { const bal = num(r['Still owed']), emi = num(r.EMI), left = num(r['EMIs left']), key = r.name, rate = rates[key] ?? 18, mr = rate / 1200, p = pre[key] || 0;
      const n0 = loanMonths(bal, emi, mr), n1 = loanMonths(bal - p, emi, mr), saved = isFinite(n0) && isFinite(n1) ? emi * n0 - (emi * n1 + p) : 0;
      const end = new Date(); end.setMonth(end.getMonth() + (left || n0));
      return `<section class="sec box"><div class="sec-h"><h3>${esc(r.name)}</h3><span class="small">${left ? left + ' EMIs left' : ''}</span></div>
        <div class="kv"><span>Still owed</span><span class="num">${inr(bal)}</span></div><div class="kv"><span>EMI</span><span class="num">${inr(emi)}</span></div>
        <div class="kv"><span>Last EMI about</span><span>${MON[end.getMonth()]} ${end.getFullYear()}</span></div>
        <label class="f" for="lr-${esc(key)}">Interest rate (% a year)</label><input class="inp num" id="lr-${esc(key)}" data-lrate="${esc(key)}" inputmode="decimal" value="${rate}">
        <label class="f">What if I prepay <b class="num" id="lpv-${esc(key)}">${inr(p)}</b>?</label>
        <input type="range" data-lpre="${esc(key)}" min="0" max="${Math.max(1000, Math.floor(bal / 1000) * 1000)}" step="1000" value="${p}">
        <p class="small ${p ? 'good' : 'muted'}" id="lpr-${esc(key)}" style="margin:4px 0 0">${p ? `Ends ${n0 - n1} month${n0 - n1 === 1 ? '' : 's'} sooner and saves about ${inr(saved)} in interest` : 'Move the slider to see the interest you’d save'}</p></section>`; }).join('')
      || '<p class="small muted">No active loans – nice.</p>'}
    ${rows.filter(r => !/Active/i.test(r.Status || '')).length ? `<section class="sec box"><div class="sec-h"><h3>Closed</h3></div>${rows.filter(r => !/Active/i.test(r.Status || '')).map(r => `<div class="kv"><span>${esc(r.name)}</span><span class="small good">${esc(r.Status)}</span></div>`).join('')}</section>` : ''}
    <p class="note">From the Loans sheet. The rate is only kept on this phone – check your loan statement for it.</p>`;
}
function bindLoans() {
  document.querySelectorAll('[data-lrate]').forEach(i => i.onchange = () => { const r = LS.get('loanRates', {}); r[i.dataset.lrate] = Number(i.value) || 0; LS.set('loanRates', r); render(); });
  document.querySelectorAll('[data-lpre]').forEach(i => i.oninput = () => { (state.prepay || (state.prepay = {}))[i.dataset.lpre] = Number(i.value); const y = window.scrollY; render(); window.scrollTo(0, y); });
}
function forecast() {
  const accts = bankAccounts(), now = bankAccounts().reduce((s, a) => s + a.bal, 0); if (!accts.length) return null;
  const today = dayOnly(new Date()), end = new Date(today.getFullYear(), today.getMonth() + 1, 0), ym = isoToday().slice(0, 7);
  const salaryIn = tbl('AppRecent').some(t => t.Type === 'Income' && /Employment/i.test(t.Category) && String(t.Date).startsWith(ym));
  const net = netPay(), salary = !salaryIn && isFinite(net) ? net : 0;
  let fixed = 0; const fixedNames = [];
  for (const kind of ['sips', 'bills']) for (const { it, st } of statuses(kind)) if (st.s !== 'paid' && st.due <= end && st.due >= new Date(today.getFullYear(), today.getMonth(), 1)) { fixed += Number(it.amount) || 0; fixedNames.push(it.name); }
  const months = tbl('AppMonthly').slice(-4, -1), avgSpend = months.length ? months.reduce((s, m) => s + (Number(m.Spending) || 0), 0) / months.length : 0;
  const billMonthly = (state.docs.bills || []).reduce((s, b) => s + (Number(b.amount) || 0) / (b.every === 'Yearly' ? 12 : 1), 0);
  const be = bankEnd(), from = be && be > new Date(today.getFullYear(), today.getMonth(), 1) ? be : new Date(today.getFullYear(), today.getMonth(), 0);
  const daysLeft = Math.max(0, Math.round((end - from) / DAY)), dim = end.getDate();
  const variable = Math.max(0, avgSpend - billMonthly) * daysLeft / dim;
  return { now, salary, fixed, fixedNames, variable, end: now + salary - fixed - variable, endDate: end, asOf: be };
}
function viewForecast() {
  const f = forecast(); if (!f) return '';
  return `<section class="sec box"><div class="sec-h"><h3>Month-end forecast</h3><span class="small">${fDay(f.endDate).replace(/ \d{4}$/, '')}</span></div>
    <div class="kv"><span>In your banks now<br><span class="small muted">${f.asOf ? 'as of ' + esc(fDay(f.asOf)) : ''}</span></span><span class="num">${inr(f.now)}</span></div>
    ${f.salary ? `<div class="kv"><span>Salary still to come</span><span class="num good">+${inr(f.salary)}</span></div>` : ''}
    <div class="kv"><span>SIPs and bills still due<br><span class="small muted">${esc(f.fixedNames.slice(0, 4).join(', ')) || 'none tracked'}</span></span><span class="num">−${inr(f.fixed)}</span></div>
    <div class="kv"><span>Usual day-to-day spending<br><span class="small muted">3-month average, rest of the month</span></span><span class="num">−${inr(f.variable)}</span></div>
    <div class="kv"><span><b>Expected at month end</b></span><span class="num ${f.end < 0 ? 'bad' : ''}"><b>${inr(f.end)}</b></span></div></section>`;
}

/* ======================================================================
   Insurance cover check (app/profile.json)
   ====================================================================== */
function viewCover() {
  const p = state.docs.profile || {}, k = kpis(), net = netPay(), yearly = isFinite(net) ? net * 12 : 0, loans = n(k.loans_owed) || 0;
  const needLife = Math.round((yearly * 10 + loans) / 100000) * 100000, life = Number(p.lifeCover) || 0, health = Number(p.healthCover) || 0, people = Number(p.people) || 1;
  const needHealth = people > 2 ? 1000000 : 500000;
  const row = (l, have, need, why) => `<div class="kv"><span>${l}<br><span class="small muted">${why}</span></span><span class="num" style="text-align:right">${inr(have)}<br><span class="small ${have >= need ? 'good' : 'bad'}">${have >= need ? 'enough' : `short by ${inr(need - have)}`}</span></span></div>`;
  return `<section class="card"><h2 style="margin-top:0">Your cover</h2>
      <label class="f" for="cv-life">Life cover – all term / LIC policies (₹)</label><input class="inp num" id="cv-life" inputmode="numeric" value="${esc(p.lifeCover ?? '')}" placeholder="Sum assured">
      <label class="f" for="cv-health">Health cover – company + personal (₹)</label><input class="inp num" id="cv-health" inputmode="numeric" value="${esc(p.healthCover ?? '')}">
      <label class="f" for="cv-people">People covered by your health policy</label><input class="inp num" id="cv-people" inputmode="numeric" value="${esc(p.people ?? 1)}">
      <button class="primary" id="cv-save">Save</button></section>
    ${life || health ? `<section class="sec box"><div class="sec-h"><h3>Check</h3></div>
      ${row('Life cover', life, needLife, `About 10 × yearly take-home (${inr(yearly)}) + loans (${inr(loans)})`)}
      ${row('Health cover', health, needHealth, people > 2 ? '₹10 lakh is a sensible floor for a family' : '₹5 lakh is a sensible floor for one or two people')}
      <p class="small muted" style="margin:8px 0 0">Rules of thumb, not advice. Company health cover ends when you change jobs, so personal cover matters.</p></section>` : ''}`;
}
function bindCover() {
  const $ = id => document.getElementById(id);
  $('cv-save').onclick = async () => {
    const v = x => { const t = $(x).value.replace(/[,₹\s]/g, ''); return t === '' ? null : Number(t); };
    const prof = { ...(state.docs.profile || {}), lifeCover: v('cv-life'), healthCover: v('cv-health'), people: v('cv-people') || 1 };
    const old = state.docs.profile; state.docs.profile = prof;
    try { await saveDoc('profile', 'Insurance cover'); toast('Saved'); } catch (e) { state.docs.profile = old; toast(e.message || 'Couldn’t reach GitHub'); }
    render();
  };
}

/* ======================================================================
   Coaching: monthly fee per client (who still owes), costs, profit, tax set aside (app/coaching.json)
   ====================================================================== */
function viewCoachingMoney() {
  const c = state.docs.coaching || {}, fees = c.fees || {}, list = receiptsList(), inrOf = r => r.cur === '₹' ? Number(r.total) || 0 : Number(r.inr) || 0;
  const ym = isoToday().slice(0, 7), monthOf = r => r.month || String(r.date || '').slice(0, 7);
  const clients = [...new Set(list.map(r => r.client).filter(Boolean))].sort();
  const got = cl => list.filter(r => r.client === cl && monthOf(r) === ym).reduce((s, r) => s + inrOf(r), 0);
  const owed = Object.entries(fees).filter(([cl, f]) => Number(f) > 0).map(([cl, f]) => ({ cl, f: Number(f), got: got(cl) })).filter(x => x.got < x.f);
  const inc = list.filter(r => monthOf(r) === ym).reduce((s, r) => s + inrOf(r), 0), costs = Number(c.costs) || 0, taxPct = Number(c.taxPct ?? 10);
  return `<section class="sec box"><div class="sec-h"><h3>Still to pay this month</h3></div>
      ${owed.length ? owed.map(x => `<div class="kv"><span>${esc(x.cl)}<br><span class="small muted">fee ${inr(x.f)}${x.got ? `, paid ${inr(x.got)}` : ''}</span></span><span class="num warn">${inr(x.f - x.got)}</span></div>`).join('')
        : `<p class="small muted" style="margin:0">${Object.keys(fees).length ? 'Everyone has paid this month.' : 'Set each client’s monthly fee below to see who still owes you.'}</p>`}</section>
    <section class="sec box"><div class="sec-h"><h3>Profit this month</h3></div>
      <div class="kv"><span>Coaching income</span><span class="num">${inr(inc)}</span></div>
      <div class="kv"><span>Costs (tools, transfer fees…)</span><span class="num">−${inr(costs)}</span></div>
      <div class="kv"><span><b>Profit</b></span><span class="num"><b>${inr(inc - costs)}</b></span></div>
      <div class="kv"><span>Set aside for tax (${taxPct}%)</span><span class="num warn">${inr(Math.max(0, (inc - costs) * taxPct / 100))}</span></div></section>
    <section class="card"><h2 style="margin-top:0">Fees and costs</h2>
      ${clients.map((cl, i) => `<label class="f" for="cf-${i}">${esc(cl)} – monthly fee (₹)</label><input class="inp num" id="cf-${i}" data-cl="${esc(cl)}" inputmode="decimal" value="${esc(fees[cl] ?? '')}" placeholder="leave empty if not monthly">`).join('') || '<p class="small muted">Clients appear here once you have receipts.</p>'}
      <label class="f" for="cf-costs">Monthly costs (₹)</label><input class="inp num" id="cf-costs" inputmode="decimal" value="${esc(c.costs ?? '')}">
      <label class="f" for="cf-tax">Set aside for tax (%)</label><input class="inp num" id="cf-tax" inputmode="decimal" value="${esc(taxPct)}">
      <button class="primary" id="cf-save">Save</button></section>`;
}
function bindCoachingMoney() {
  const s = document.getElementById('cf-save'); if (!s) return;
  s.onclick = async () => {
    const fees = {}; document.querySelectorAll('[data-cl]').forEach(i => { const v = i.value.replace(/[,₹\s]/g, ''); if (v !== '' && Number(v) > 0) fees[i.dataset.cl] = Number(v); });
    const num2 = id => { const v = document.getElementById(id).value.replace(/[,₹\s%]/g, ''); return v === '' ? null : Number(v); };
    const old = state.docs.coaching; state.docs.coaching = { fees, costs: num2('cf-costs') || 0, taxPct: num2('cf-tax') ?? 10 };
    try { await saveDoc('coaching', 'Coaching fees and costs'); toast('Saved'); } catch (e) { state.docs.coaching = old; toast(e.message || 'Couldn’t reach GitHub'); }
    render();
  };
}

/* ======================================================================
   This phone: install / notifications / storage check, token expiry
   ====================================================================== */
function viewPhoneCheck() {
  const standalone = matchMedia('(display-mode: standalone)').matches || navigator.standalone === true;
  const perm = 'Notification' in window ? Notification.permission : 'unsupported', sw = 'serviceWorker' in navigator && !!navigator.serviceWorker.controller;
  const ios = /iPhone|iPad/.test(navigator.userAgent);
  const line = (ok, t, fix) => `<div class="kv"><span>${t}${!ok && fix ? `<br><span class="small muted">${fix}</span>` : ''}</span><span class="small ${ok ? 'good' : 'warn'}">${ok ? '✓' : 'Fix'}</span></div>`;
  const te = tokenDays();
  return `<section class="card"><h2 style="margin-top:0">This phone</h2>
    ${line(standalone, 'Installed on the home screen', ios ? 'Safari → Share → Add to Home Screen, then open MyMoney from the icon' : 'Chrome menu ⋮ → Add to Home screen / Install app, then open it from the icon')}
    ${line(perm === 'granted', 'Notifications allowed', perm === 'denied' ? 'Blocked – allow them in the phone’s Settings → Apps → Chrome / MyMoney → Notifications' : 'Turn them on in Notifications and settings')}
    ${line(sw, 'Works offline', 'Close and reopen the app once')}
    ${line(!!state.storageSafe, 'Data kept safe from clean-up', 'Install the app – installed apps keep their storage')}
    ${line(te == null || te > 14, te == null ? 'GitHub token expiry (Excel reports it after the next sync)' : `GitHub token expires in ${te} days`, 'Make a new token on GitHub and paste it here and in Excel’s AppLink')}
  </section>`;
}
function tokenDays() {
  const s = state.data && state.data.tokenExpires; if (!s) return null;
  const d = new Date(String(s).replace(' UTC', 'Z').replace(' ', 'T').replace(/ ([+-]\d{4})$/, '$1'));
  return isNaN(d) ? null : Math.floor((d - Date.now()) / DAY);
}

/* ======================================================================
   Monthly review: last month vs plan, what went over, three suggestions
   ====================================================================== */
function monthReview() {
  const k = kpis(), ym = String(k.month || '').slice(0, 7), mrow = tbl('AppMonthly').find(r => String(r.Month).startsWith(ym)) || {};
  const plan = activePlan(), net = netPay(), limit = isFinite(net) ? net - (Number(plan.savings) || 0) : null;
  const rows = tbl('AppBudget').filter(r => r.Category).map(r => {
    const lim = plan.limits && plan.limits[r.Category] != null ? Number(plan.limits[r.Category]) : n(r.Target);
    return { cat: r.Category, spent: n(r.Spent) || 0, lim: lim || 0 };
  });
  const over = rows.filter(r => r.lim > 0 && r.spent > r.lim * 1.02).sort((a, b) => (b.spent - b.lim) - (a.spent - a.lim));
  const sugg = [];
  const rate = Number(k.savings_rate) || 0, target = Number(k.savings_target) || 0.2;
  const flex = over.find(o => !FIXED.has(o.cat) && !/Credit Card|UPI – Other/i.test(o.cat));     // fixed costs (EMIs, rent) can't simply be capped
  if (flex) sugg.push(`Cap ${flex.cat} at ${inr(flex.lim)} – it was ${inr(flex.spent)}, ${inr(flex.spent - flex.lim)} over.`);
  const cc = over.find(o => /Credit Card/i.test(o.cat));
  if (cc && sugg.length < 2) sugg.push(`Card bills were ${inr(cc.spent)}. Check the card statement for spends that belong in other categories, and keep new card spends inside the plan.`);
  const upi = rows.find(r => /UPI – Other/i.test(r.cat));
  if (upi && upi.spent > 5000) sugg.push(`${inr(upi.spent)} went to “UPI – Other”. Sort those payees into real categories so the plan is accurate (Money → Sort UPI).`);
  if (rate < target) sugg.push(`You saved ${Math.round(rate * 100)}% of income – aim for ${Math.round(target * 100)}%. Moving ${inr((target - rate) * (Number(mrow.Income) || 0))} a month into a SIP right after salary day closes the gap.`);
  const tr = categoryTrends().find(x => x.diff > 2000);
  if (tr && sugg.length < 3) sugg.push(`${tr.name} is rising: ${inr(tr.now)} against a usual ${inr(tr.avg)}.`);
  if (sugg.length < 3 && (inv() || {}).epf) sugg.push('Check Tax → 80C before March so nothing is left unused.');
  if (sugg.length < 3 && n(k.loans_owed) > 0) sugg.push(`Loans still owed: ${inr(k.loans_owed)}. Even a small prepayment saves interest (Money → Loans).`);
  return { ym, mrow, limit, spent: Number(mrow.Spending) || n(k.spending_month) || 0, over, sugg: sugg.slice(0, 3), rate, target };
}
function viewReview() {
  const r = monthReview(), d = pDate(r.ym + '-01');
  return `<section class="hero"><p class="hero-k" style="margin-top:0">${d ? MONF[d.getMonth()] + ' ' + d.getFullYear() : 'Last month'} in review</p>
      <h2 class="hero-v num">${inr(r.spent)}<span style="font-size:16px;font-weight:500"> spent</span></h2>
      <p class="hero-s" style="margin:0">${r.limit != null ? (r.spent > r.limit ? `${inr(r.spent - r.limit)} over your net-pay limit of ${inr(r.limit)}` : `${inr(r.limit - r.spent)} under your net-pay limit`) : ''}</p>
      <div class="hero-tiles"><div><span>Income</span><b class="num">${inr(r.mrow.Income)}</b></div><div><span>Saved</span><b class="num">${inr(r.mrow.Saved)} · ${Math.round(r.rate * 100)}%</b></div></div></section>
    <section class="sec box"><div class="sec-h"><h3>Three things to change</h3></div>
      ${r.sugg.map((s, i) => `<div class="kv" style="justify-content:flex-start"><span class="sugg-n">${i + 1}</span><span>${esc(s)}</span></div>`).join('') || '<p class="small muted">Nothing stands out – a good month.</p>'}</section>
    <section class="sec box"><div class="sec-h"><h3>Over the plan</h3><span class="small">${r.over.length}</span></div>
      ${r.over.slice(0, 8).map(o => `<div class="kv"><span>${esc(o.cat)}<br><span class="small muted">plan ${inr(o.lim)}</span></span><span class="num bad">+${inr(o.spent - o.lim)}</span></div>`).join('') || '<p class="small muted">Every category stayed within plan.</p>'}</section>
    <button class="secondary" data-gotab="budget">Change the plan</button>`;
}

/* ======================================================================
   Sort "UPI – Other": group by payee, give each payee a real category
   ====================================================================== */
function upiGroups() {
  const since = new Date(); since.setMonth(since.getMonth() - 6);
  const g = {};
  for (const t of allTx()) {
    if (t._pending || t.Type !== 'Expenses' || !/UPI – Other|Uncategorised/i.test(t.Category)) continue;
    const d = pDate(t.Date); if (!d || d < since) continue;
    const p = upiParts(t.Details).payee; if (!p) continue;
    const key = p.toUpperCase(), o = g[key] || (g[key] = { payee: p, n: 0, sum: 0, last: t, cat: t.Category });
    o.n++; o.sum += Math.abs(n(t.Amount)); if (t.Date > o.last.Date) o.last = t;
  }
  return Object.values(g).sort((a, b) => b.sum - a.sum);
}
function viewSort() {
  const gs = upiGroups(), cats = catsFor('Expenses'), tot = gs.reduce((s, x) => s + x.sum, 0);
  return `<section class="hero"><p class="hero-k" style="margin-top:0">“UPI – Other”, last 6 months</p><h2 class="hero-v num">${inr(tot)}</h2>
      <p class="hero-s" style="margin:0">${gs.length} payees. Pick a category for the big ones – every past and future payment to them moves.</p></section>
    <section class="sec box">${gs.slice(0, 40).map((x, i) => `<div class="sortrow"><div class="row"><span><b>${esc(x.payee)}</b><br><span class="small muted">${x.n} payment${x.n === 1 ? '' : 's'} · last ${esc(fDay(pDate(x.last.Date)))}</span></span><b class="num">${inr(x.sum)}</b></div>
        <div class="row" style="gap:8px;margin-top:6px"><select class="inp" data-sort-cat="${i}"><option value="">Choose category…</option>${cats.filter(c => !/UPI – Other/.test(c)).map(c => `<option>${esc(c)}</option>`).join('')}</select>
        <button class="chip" data-sort-go="${i}">Save</button></div></div>`).join('') || '<p class="small muted">Nothing left in “UPI – Other”.</p>'}</section>
    <p class="note">Each choice is saved as a rule for that payee; Excel adds it to the top of Auto Rules next time it opens.</p>`;
}
function bindSort() {
  const gs = upiGroups();
  document.querySelectorAll('[data-sort-go]').forEach(b => b.onclick = async () => {
    const i = Number(b.dataset.sortGo), sel = document.querySelector(`[data-sort-cat="${i}"]`), g = gs[i];
    if (!sel.value) { toast('Choose a category first'); return; }
    if (g.payee.length < 3) { toast('Payee name too short to match safely'); return; }
    b.disabled = true; b.textContent = 'Saving…';
    const now = new Date(), entry = { id: `${now.toISOString().replace(/[-:]/g, '').slice(0, 15)}-${Math.random().toString(36).slice(2, 7)}`, keyword: g.payee.slice(0, 60),
      direction: 'Debit', type: 'Expenses', category: sel.value, scope: 'payee', tx: txKey(g.last), note: `Payee sorted from phone (was ${g.cat})`, createdAt: now.toISOString() };
    try { await GH.putJson(`recat/${entry.id}.json`, entry, `Category change ${entry.id}`); state.recats.unshift({ id: entry.id, entry }); LS.set('recats', state.recats); toast(`${g.payee} → ${sel.value}`); render(); }
    catch (e) { toast(e.message || 'Couldn’t reach GitHub'); b.disabled = false; b.textContent = 'Save'; }
  });
}

/* ======================================================================
   Shared expenses: who owes whom (app/shared.json)
   ====================================================================== */
function viewShared() {
  const list = state.docs.shared || [], open = list.filter(x => !x.settled), f = state.shEdit;
  const by = {}; open.forEach(x => { by[x.person] = (by[x.person] || 0) + (x.dir === 'owe' ? -1 : 1) * (Number(x.amount) || 0); });
  const owedToMe = Object.values(by).filter(v => v > 0).reduce((s, v) => s + v, 0), iOwe = -Object.values(by).filter(v => v < 0).reduce((s, v) => s + v, 0);
  return `<section class="hero"><div class="hero-tiles" style="margin-top:0"><div><span>Owed to you</span><b class="num">${inr(owedToMe)}</b></div><div><span>You owe</span><b class="num">${inr(iOwe)}</b></div></div></section>
    ${f ? `<section class="card"><h2 style="margin-top:0">${f.id ? 'Edit' : 'Add'}</h2>
      <div class="seg" id="sh-dir"><button data-v="owed" class="${f.dir !== 'owe' ? 'on' : ''}">They owe me</button><button data-v="owe" class="${f.dir === 'owe' ? 'on' : ''}">I owe them</button></div>
      <label class="f" for="sh-person">Person</label><input class="inp" id="sh-person" list="sh-people" value="${esc(f.person || '')}"><datalist id="sh-people">${[...new Set(list.map(x => x.person))].map(p => `<option value="${esc(p)}">`).join('')}</datalist>
      <label class="f" for="sh-amt">Amount (₹)</label><input class="inp num" id="sh-amt" inputmode="decimal" value="${esc(f.amount || '')}">
      <label class="f" for="sh-note">For</label><input class="inp" id="sh-note" maxlength="60" value="${esc(f.note || '')}" placeholder="e.g. Dinner, movie tickets">
      <label class="f" for="sh-date">Date</label><input class="inp" id="sh-date" type="date" value="${esc(f.date || isoToday())}" max="${isoToday()}">
      <button class="primary" id="sh-save">Save</button><button class="secondary" id="sh-cancel">Cancel</button></section>` : `<button class="primary" id="sh-add">＋ Add who owes what</button>`}
    ${Object.keys(by).length ? `<section class="sec box"><div class="sec-h"><h3>By person</h3></div>${Object.entries(by).filter(([, v]) => Math.abs(v) >= 1).sort((a, b) => Math.abs(b[1]) - Math.abs(a[1])).map(([p, v]) => `<div class="kv"><span>${esc(p)}</span><span class="num ${v > 0 ? 'good' : 'bad'}">${v > 0 ? 'owes you ' : 'you owe '}${inr(Math.abs(v))}</span></div>`).join('')}</section>` : ''}
    <section class="sec box"><div class="sec-h"><h3>Open</h3><span class="small">${open.length}</span></div>
      ${open.map(x => `<div class="kv"><span>${esc(x.person)} · ${esc(x.note || '')}<br><span class="small muted">${esc(fDay(pDate(x.date)))} · ${x.dir === 'owe' ? 'you owe' : 'owes you'}</span></span>
        <span class="num" style="text-align:right">${inr(x.amount)}<br><button class="linkbtn" data-sh-settle="${esc(x.id)}">Settled</button> · <button class="linkbtn" data-sh-edit="${esc(x.id)}">Edit</button></span></div>`).join('') || '<p class="small muted">Nothing open.</p>'}</section>
    ${list.some(x => x.settled) ? `<p class="small muted" style="text-align:center">${list.filter(x => x.settled).length} settled earlier</p>` : ''}`;
}
function bindShared() {
  const $ = id => document.getElementById(id), save = async (list, msg) => {
    const old = state.docs.shared; state.docs.shared = list;
    try { await saveDoc('shared', msg); toast('Saved'); state.shEdit = null; } catch (e) { state.docs.shared = old; toast(e.message || 'Couldn’t reach GitHub'); }
    render();
  };
  if ($('sh-add')) $('sh-add').onclick = () => { state.shEdit = { dir: 'owed' }; render(); };
  document.querySelectorAll('[data-sh-edit]').forEach(b => b.onclick = () => { state.shEdit = { ...(state.docs.shared || []).find(x => x.id === b.dataset.shEdit) }; render(); window.scrollTo(0, 0); });
  document.querySelectorAll('[data-sh-settle]').forEach(b => b.onclick = () => save((state.docs.shared || []).map(x => x.id === b.dataset.shSettle ? { ...x, settled: isoToday() } : x), 'Settled'));
  const f = state.shEdit; if (!f) return;
  const read = () => { f.person = $('sh-person').value.trim(); f.amount = $('sh-amt').value.replace(/[,₹\s]/g, ''); f.note = $('sh-note').value.trim(); f.date = $('sh-date').value; };
  document.querySelectorAll('#sh-dir button').forEach(b => b.onclick = () => { read(); f.dir = b.dataset.v; render(); });
  $('sh-cancel').onclick = () => { state.shEdit = null; render(); };
  $('sh-save').onclick = () => {
    read(); if (!f.person) { toast('Enter the person'); return; } if (!(Number(f.amount) > 0)) { toast('Enter the amount'); return; }
    const item = { id: f.id || uid('x'), dir: f.dir || 'owed', person: f.person, amount: Number(f.amount), note: f.note, date: f.date || isoToday() };
    const list = state.docs.shared || [];
    save(f.id ? list.map(x => x.id === f.id ? item : x) : [...list, item], `${item.dir === 'owe' ? 'I owe' : 'Owed by'} ${item.person}`);
  };
}

/* ======================================================================
   Voice entry: "spent 250 on vegetables cash"
   ====================================================================== */
const CAT_WORDS = { Groceries: /vegetable|grocer|milk|fruit|kirana|zepto|blinkit|instamart|bigbasket/i, 'Food Online': /swiggy|zomato|food|lunch|dinner|breakfast|snack|tea|coffee/i,
  Travel: /auto|cab|uber|ola|rapido|bus|train|metro|petrol|fuel|flight|ticket/i, Medical: /medicine|doctor|hospital|pharmacy|clinic/i, Shopping: /shirt|clothes|shoes|amazon|flipkart|shopping/i,
  Utilities: /electric|power|water bill|gas|recharge|internet|wifi|broadband/i, Housing: /rent/i, Donation: /donat|temple|church|mosque/i };
function parseSpoken(txt) {
  const s = String(txt || '').toLowerCase();
  const words = { hundred: 100, thousand: 1000, lakh: 100000 };
  let amount = null;
  const m = s.match(/(\d[\d,]*(?:\.\d+)?)\s*(k|thousand|lakh|hundred)?/);
  if (m) { amount = Number(m[1].replace(/,/g, '')); if (m[2]) amount *= m[2] === 'k' ? 1000 : words[m[2]]; }
  const type = /\b(got|received|earned|income|salary|credited)\b/.test(s) ? 'Income' : 'Expenses';
  const paid = /\bcash\b/.test(s) ? 'Cash' : /\b(upi|gpay|google pay|phonepe|paytm)\b/.test(s) ? 'UPI' : /\bcard\b/.test(s) ? 'Card' : 'Cash';
  const note = (s.match(/\b(?:on|for|at|from)\s+(.+?)(?:\s+(?:by|in|with|using|via)\s+(?:cash|upi|card|gpay|phonepe|paytm)|\s+(?:cash|upi|card)\b|$)/) || [])[1] || '';
  let category = '';
  const cats = catsFor(type);
  if (type === 'Expenses') for (const [c, re] of Object.entries(CAT_WORDS)) if (re.test(s) && cats.includes(c)) { category = c; break; }
  if (!category) category = cats.find(c => !/^cash$/i.test(c) && s.includes(c.toLowerCase())) || '';
  if (!category && type === 'Income') category = cats.find(c => /other income/i.test(c)) || '';
  const date = /\byesterday\b/.test(s) ? (() => { const d = new Date(Date.now() - DAY); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; })() : isoToday();
  return { amount, type, paidWith: paid, note: note.replace(/\b(rupees?|rs)\b/g, '').trim().replace(/^./, c => c.toUpperCase()).slice(0, 60), category, date };
}
function voiceSupported() { return !!(window.SpeechRecognition || window.webkitSpeechRecognition); }
function startVoice() {
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SR) { toast('Voice input isn’t available in this browser – try Chrome'); return; }
  const r = new SR(); r.lang = 'en-IN'; r.interimResults = false; r.maxAlternatives = 1;
  const b = document.getElementById('add-voice'); if (b) { b.classList.add('on'); b.textContent = '● Listening…'; }
  r.onresult = e => { const txt = e.results[0][0].transcript; applySpoken(txt); };
  r.onerror = e => { toast(e.error === 'not-allowed' ? 'Allow the microphone for MyMoney to use voice' : 'Didn’t catch that – try again'); render(); };
  r.onend = () => { const bb = document.getElementById('add-voice'); if (bb) { bb.classList.remove('on'); bb.textContent = '🎤 Say it'; } };
  r.start();
}
function applySpoken(txt) {
  const p = parseSpoken(txt), a = state.add || (state.add = newAdd());
  if (!p.amount) { toast(`Heard “${txt}” – no amount in it`); return; }
  Object.assign(a, { amount: String(p.amount), type: p.type, paidWith: p.paidWith, note: p.note, date: p.date, category: p.category });
  state.tab = 'add'; render(); toast(`Heard: “${txt}” – check and save`);
}

/* ======================================================================
   Year in review
   ====================================================================== */
function viewYear() {
  const years = [...new Set(tbl('AppRecent').map(t => String(t.Date).slice(0, 4)))].filter(y => /^\d{4}$/.test(y)).sort().reverse();
  const y = state.yearPick && years.includes(state.yearPick) ? state.yearPick : years[0];
  const tx = tbl('AppRecent').filter(t => String(t.Date).startsWith(y));
  const sum = ty => tx.filter(t => t.Type === ty).reduce((s, t) => s + Math.abs(n(t.Amount)), 0);
  const inc = sum('Income'), sp = sum('Expenses'), sv = sum('Savings');
  const cats = {}; tx.filter(t => t.Type === 'Expenses').forEach(t => { cats[t.Category] = (cats[t.Category] || 0) + Math.abs(n(t.Amount)); });
  const top = Object.entries(cats).sort((a, b) => b[1] - a[1]).slice(0, 5), maxC = top.length ? top[0][1] : 1;
  const months = {}; tx.filter(t => t.Type === 'Expenses').forEach(t => { const m = String(t.Date).slice(5, 7); months[m] = (months[m] || 0) + Math.abs(n(t.Amount)); });
  const mm = Object.entries(months).sort((a, b) => b[1] - a[1]);
  const big = tx.filter(t => t.Type === 'Expenses' && !/Loans|Credit Card|Housing/i.test(t.Category)).sort((a, b) => Math.abs(n(b.Amount)) - Math.abs(n(a.Amount))).slice(0, 3);
  const payees = {}; tx.filter(t => t.Type === 'Expenses').forEach(t => { const p = upiParts(t.Details).payee; if (p) payees[p.toUpperCase()] = (payees[p.toUpperCase()] || 0) + 1; });
  const fav = Object.entries(payees).sort((a, b) => b[1] - a[1])[0];
  const nw = ((state.history && state.history.nw) || []).filter(p => p.d.startsWith(y));
  return `<div class="chips" style="padding-top:0">${years.slice(0, 5).map(x => `<button class="chip ${x === y ? 'on' : ''}" data-yr="${x}">${x}</button>`).join('')}</div>
    <section class="hero"><p class="hero-k" style="margin-top:0">${y} in review</p><h2 class="hero-v num">${inr(inc)}<span style="font-size:16px;font-weight:500"> came in</span></h2>
      <div class="hero-tiles"><div><span>Spent</span><b class="num">${inr(sp)}</b></div><div><span>Saved & invested</span><b class="num">${inr(sv)} · ${inc ? Math.round(sv / inc * 100) : 0}%</b></div></div></section>
    <section class="sec box"><div class="sec-h"><h3>Where it went</h3></div>${top.map(([c, v]) => `<div class="brow"><div class="row"><span>${esc(c)}</span><b class="num">${inr(v)}</b></div><div class="meter thin"><i style="width:${(v / maxC * 100).toFixed(1)}%"></i></div></div>`).join('')}</section>
    <section class="sec box"><div class="sec-h"><h3>Highlights</h3></div>
      ${mm.length ? `<div class="kv"><span>Biggest spending month</span><span>${MONF[Number(mm[0][0]) - 1]} · <b class="num">${inr(mm[0][1])}</b></span></div>
      <div class="kv"><span>Lightest month</span><span>${MONF[Number(mm[mm.length - 1][0]) - 1]} · <b class="num">${inr(mm[mm.length - 1][1])}</b></span></div>` : ''}
      ${fav ? `<div class="kv"><span>Paid most often</span><span>${esc(fav[0])} · ${fav[1]} times</span></div>` : ''}
      <div class="kv"><span>Transactions</span><span class="num">${tx.length.toLocaleString('en-IN')}</span></div>
      ${nw.length > 1 ? `<div class="kv"><span>Net worth change (tracked days)</span><span class="num ${nw[nw.length - 1].v >= nw[0].v ? 'good' : 'bad'}">${inr(nw[nw.length - 1].v - nw[0].v)}</span></div>` : ''}</section>
    ${big.length ? `<section class="sec box"><div class="sec-h"><h3>Biggest single spends</h3></div>${big.map(t => `<div class="kv"><span>${esc(describe(t).title)}<br><span class="small muted">${esc(fDay(pDate(t.Date)))} · ${esc(t.Category)}</span></span><span class="num">${inr(Math.abs(n(t.Amount)))}</span></div>`).join('')}</section>` : ''}
    <button class="secondary noprint" onclick="window.print()">Save as PDF</button>`;
}
function bindYear() { document.querySelectorAll('[data-yr]').forEach(b => b.onclick = () => { state.yearPick = b.dataset.yr; render(); }); }

/* ======================================================================
   Family view: view-only phones, limited to the screens you choose (app/family.json)
   ====================================================================== */
const FAMILY_SCREENS = [['home', 'Home'], ['salary', 'Salary'], ['money', 'Money'], ['invest', 'Invest'], ['future', 'Retire'], ['goals', 'Goals'], ['bills', 'Bills'], ['tx', 'Transactions']];
function isViewer() { return !!(state.cfg && state.cfg.viewOnly); }
function famAllowed() { const f = state.docs.family || {}; return Array.isArray(f.screens) && f.screens.length ? f.screens : ['home', 'money', 'future', 'goals']; }
function viewFamily() {
  if (isViewer()) return `<section class="card"><h2 style="margin-top:0">Family view</h2><p class="small muted">This phone is view-only. Ask the owner to change what you can see.</p>
    <button class="secondary" id="fam-leave" style="color:var(--bad)">Stop view-only on this phone</button></section>`;
  const allowed = new Set(famAllowed()), site = location.origin + location.pathname.replace(/index\.html$/, '');
  const link = `${site}#family=${encodeURIComponent(state.cfg.owner)}/${encodeURIComponent(state.cfg.repo)}`;
  return `<section class="card"><h2 style="margin-top:0">Family view</h2>
      <p class="small muted" style="margin:0 0 8px">Let your spouse or family see your finances on their phone – read-only, with their own PIN, and only the screens you tick.</p>
      ${FAMILY_SCREENS.map(([k, l]) => `<label class="tick"><input type="checkbox" data-fam="${k}" ${allowed.has(k) ? 'checked' : ''}> ${l}</label>`).join('')}
      <button class="primary" id="fam-save">Save screens</button></section>
    <section class="card"><h2 style="margin-top:0">Set up their phone</h2>
      <ol class="steps small">
        <li>On GitHub, make a <b>new fine-grained token</b> for them: only your data repo, <b>Contents: Read-only</b>. A read-only token means their phone can’t change anything, even by mistake.</li>
        <li>Send them this link and open it on their phone:<br><code class="linkbox" id="fam-link">${esc(link)}</code><button class="secondary" id="fam-copy">Copy link</button></li>
        <li>They paste the read-only token, tap Save, then set their own PIN in More → Lock.</li>
      </ol></section>`;
}
function bindFamily() {
  const $ = id => document.getElementById(id);
  if ($('fam-leave')) $('fam-leave').onclick = () => { if (!confirmInline($('fam-leave'))) return; state.cfg.viewOnly = false; LS.set('cfg', state.cfg); state.tab = 'home'; render(); };
  if ($('fam-copy')) $('fam-copy').onclick = () => { navigator.clipboard.writeText($('fam-link').textContent).then(() => toast('Link copied'), () => toast('Copy it by pressing on the link')); };
  if ($('fam-save')) $('fam-save').onclick = async () => {
    const screens = [...document.querySelectorAll('[data-fam]')].filter(i => i.checked).map(i => i.dataset.fam);
    if (!screens.length) { toast('Tick at least one screen'); return; }
    const old = state.docs.family; state.docs.family = { screens };
    try { await saveDoc('family', 'Family view screens'); toast('Saved'); } catch (e) { state.docs.family = old; toast(e.message || 'Couldn’t reach GitHub'); }
    render();
  };
}
/* opened from the family link: owner/repo filled in, view-only */
function handleFamilyLink() {
  const m = location.hash.match(/^#family=([^/]+)\/(.+)$/); if (!m) return;
  history.replaceState(null, '', location.pathname);
  if (state.cfg && !state.cfg.viewOnly) { toast('This phone is already linked as the owner'); return; }
  state.pendingFamily = { owner: decodeURIComponent(m[1]), repo: decodeURIComponent(m[2]) };
  state.tab = 'settings';
}

/* Retirement: ring + projection, quick sliders, savings today (format from the reference) */
function retireCard(x, b) {
  const r = retirePlan(x, b), p = Math.max(0, Math.min(1, r.atPace / (r.corpus || 1))), R = 46, C = 2 * Math.PI * R, ok = r.gap <= 0;
  return `<div class="rc"><svg viewBox="0 0 110 110" width="112" height="112" aria-label="${Math.round(p * 100)}% of goal">
      <circle cx="55" cy="55" r="${R}" fill="none" stroke="rgba(255,255,255,.18)" stroke-width="10"/>
      <circle cx="55" cy="55" r="${R}" fill="none" stroke="#e6f1ff" stroke-width="10" stroke-linecap="round" stroke-dasharray="${(C * p).toFixed(1)} ${C.toFixed(1)}" transform="rotate(-90 55 55)"/>
      <text x="55" y="55" text-anchor="middle" font-size="20" font-weight="600" fill="#fff">${Math.round(p * 100)}%</text>
      <text x="55" y="72" text-anchor="middle" font-size="10" fill="rgba(255,255,255,.75)">of goal</text></svg>
    <div><p class="hero-k" style="margin:0">Projected at age ${x.retire}</p><h2 class="hero-v num" style="font-size:30px;margin:2px 0 4px">${cr(r.atPace)}</h2>
      <p class="hero-s" style="margin:0">${ok ? 'On track for your goal' : `Need ${cr(r.corpus)}, short by ${cr(r.gap)}`}</p></div></div>`;
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
        ${row('data-go="coaching"', 'Coaching earnings', '')}
        ${row('href="receipts.html#history"', 'Past receipts', rc.count ? `${rc.count} receipts, ${inr(rc.all)} in all${rc.last ? `. Last: ${esc(rc.last.client)}, ${fDay(pDate(rc.last.date))}` : ''}` : '')}
      </div></section>
    <section class="sec"><div class="sec-h"><h3>Plans</h3></div>
      <div class="group">
        ${row('data-go="future"', 'Retirement', '')}
        ${row('data-go="goals"', 'Goals', '')}
        ${row('data-go="sips"', 'SIPs', '')}
        ${row('data-go="bills"', 'Bills and subscriptions', '')}
        ${row('data-go="tax"', 'Tax: 80C, 80D and capital gains', '')}
        ${row('data-go="loans"', 'Loans', '')}
        ${row('data-go="cover"', 'Insurance cover check', '')}
        ${row('data-go="report"', 'Monthly report (PDF)', '')}
        ${row('data-go="review"', 'Monthly review', '')}
        ${row('data-go="year"', 'Year in review', '')}
        ${row('data-go="shared"', 'Shared expenses', '')}
        ${row('data-go="sort"', 'Sort UPI – Other', '')}
        ${row('data-go="family"', 'Family view', isViewer() ? 'This phone is view-only' : '')}
        ${row('data-go="budget"', 'Net-pay plan', '')}
        ${row('data-go="lockset"', 'Lock, look and this phone', lockCfg() ? 'Lock on' : '')}
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
  const qp = a.editId ? [] : quickPicks();
  return `${qp.length ? `<div class="chips" id="add-quick" style="padding-top:0">${qp.map((e, i) => `<button class="chip" data-q="${i}">${esc((e.note || e.category).slice(0, 22))} ${inr(e.amount)}</button>`).join('')}</div>` : ''}
    ${a.editId ? '<p class="small warn" style="margin:0 0 8px">Editing an entry that’s waiting for Excel</p>' : `<details class="sms"><summary>Paste a bank SMS or UPI message</summary>
      <textarea class="inp" id="add-sms" rows="3" placeholder="e.g. Rs.250.00 debited from A/c XX1234 on 17-11-26 to VPA swiggy@icici…"></textarea>
      <button class="secondary" id="add-sms-go">Fill in from message</button></details>`}<section class="card">
    ${voiceSupported() && !a.editId ? '<button class="secondary voice" id="add-voice">🎤 Say it</button><p class="small muted" style="margin:6px 0 10px;text-align:center">e.g. “spent 250 on vegetables cash”</p>' : ''}
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
    <button class="primary" id="add-save">${a.editId ? 'Save changes' : 'Save entry'}</button>
    ${a.editId ? '<button class="secondary" id="add-cancel">Cancel</button>' : ''}
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
  const qp = quickPicks();
  document.querySelectorAll('#add-quick .chip').forEach(b => b.onclick = () => { const e = qp[Number(b.dataset.q)]; Object.assign(a, { type: e.type, category: e.category, amount: String(e.amount), note: e.note || '', paidWith: e.paidWith || 'Cash', date: isoToday() }); render(); });
  const vb = document.getElementById('add-voice'); if (vb) vb.onclick = () => { sync1(); startVoice(); };
  const sg = document.getElementById('add-sms-go'); if (sg) sg.onclick = () => { sync1(); if (applySms(document.getElementById('add-sms').value)) render(); };
  const ac = document.getElementById('add-cancel'); if (ac) ac.onclick = () => { state.add = newAdd(); state.tab = 'tx'; render(); };
  document.getElementById('add-save').onclick = async () => {
    sync1();
    const amt = Number(String(a.amount).replace(/[,₹\s]/g, ''));
    if (!(amt > 0)) { toast('Enter an amount'); document.getElementById('add-amt').focus(); return; }
    if (!a.category) { toast('Pick a category'); return; }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(a.date)) { toast('Pick a date'); return; }
    const now = new Date(), stamp = now.toISOString().replace(/[-:]/g, '').slice(0, 15);
    const entry = { id: a.editId || `${stamp}-${Math.random().toString(36).slice(2, 7)}`, date: a.date, type: a.type, category: a.category,
      amount: Math.round(amt * 100) / 100, paidWith: a.paidWith, note: a.note.trim().slice(0, 80), createdAt: a.createdAt || now.toISOString(), ...(a.editId ? { editedAt: now.toISOString() } : {}) };
    if (a.editId) {
      try {
        await GH.putJson(`inbox/${entry.id}.json`, entry, `Phone entry edited ${entry.id}`);
        const sha = await GH.sha(`inbox/${entry.id}.json`);
        state.pending = state.pending.map(p => p.id === entry.id ? { id: entry.id, sha, entry } : p); LS.set('pending', state.pending);
        toast('Changes saved');
      } catch (e) { toast(e.message || 'Couldn’t reach GitHub – try again'); return; }
      state.add = newAdd(); state.tab = 'tx'; render(); return;
    }
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
  const c = state.cfg || (state.pendingFamily ? { ...state.pendingFamily, token: '', branch: 'main' } : { owner: '', repo: 'mymoney-data', token: '', branch: 'main' });
  return `${state.cfg ? viewNotif() : ''}<section class="card">
    <h2>${state.cfg ? 'GitHub link' : state.pendingFamily ? 'Family view – view only' : 'Connect to your data'}</h2>
    ${state.pendingFamily && !state.cfg ? '<p class="small" style="margin:0 0 6px">Paste the read-only token you were given. This phone will only be able to look, not change anything.</p>' : ''}
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
      token: document.getElementById('s-token').value.trim(), branch: document.getElementById('s-branch').value.trim() || 'main',
      ...((state.cfg && state.cfg.viewOnly) || (!state.cfg && state.pendingFamily) ? { viewOnly: true } : {}) };
    if (!cfg.owner || !cfg.repo || !cfg.token) { msg.className = 'small bad'; msg.textContent = 'Fill in username, repo and token.'; return; }
    const prev = state.cfg; state.cfg = cfg; msg.className = 'small muted'; msg.textContent = 'Checking…';
    try {
      const repo = await GH.testRepo();
      if (!repo.private) { msg.className = 'small bad'; msg.textContent = 'That repo is PUBLIC – anyone could read your figures. Make it private on GitHub first.'; state.cfg = prev; return; }
      LS.set('cfg', cfg); state.pendingFamily = null; state.tab = 'home'; state.lastError = null; render(); sync({ toast: true });
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
document.getElementById('btn-back').onclick = goBack;
document.getElementById('btn-settings').onclick = () => { state.tab = 'more'; window.scrollTo(0, 0); render(); };
window.addEventListener('online', () => sync());
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden') { if (!state.locked) LS.set('lastActive', Date.now()); return; }
  if (lockCfg() && !state.locked && Date.now() - (LS.get('lastActive', 0) || 0) > LOCK_AFTER) showLock();
  if (document.visibilityState === 'visible' && state.cfg && (!state.fetchedAt || Date.now() - new Date(state.fetchedAt) > 5 * 60e3)) sync();
});
if ('serviceWorker' in navigator && (location.protocol === 'https:' || /^(localhost|127\.0\.0\.1)$/.test(location.hostname))) {
  const hadSw = !!navigator.serviceWorker.controller;
  navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).then(r => r.update()).catch(() => {});
  navigator.serviceWorker.addEventListener('controllerchange', () => { if (hadSw && !window.__mmReloaded) { window.__mmReloaded = true; location.reload(); } });   // new version: show it at once
}
{ const h = location.hash.slice(1); if (h && h !== 'sheet' && (TITLES[h] || h === 'more') && state.cfg) state.tab = h; }
applyTheme();
handleShared();
handleFamilyLink();
if (location.hash === '#voice' && state.cfg) { history.replaceState(null, '', location.pathname); state.tab = 'add'; setTimeout(() => { if (voiceSupported()) startVoice(); }, 600); }
if (navigator.storage && navigator.storage.persist) navigator.storage.persisted().then(v => v || navigator.storage.persist()).then(v => { state.storageSafe = !!v; }).catch(() => {});
if (window.matchMedia) matchMedia('(prefers-color-scheme: light)').addEventListener?.('change', applyTheme);
render();
if (lockCfg()) showLock();
if (state.cfg) sync();
