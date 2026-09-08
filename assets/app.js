/* ═══════════════════════════════════════════════════════════════
   NEXUS Ranking Command Center — Application Logic
   Unified HLTV + VRS · KZ Studio · by Kilzys
═══════════════════════════════════════════════════════════════ */

const $ = sel => document.querySelector(sel);
const $$ = sel => document.querySelectorAll(sel);

/* ════════════════════════════════════════
   CONSTANTS
════════════════════════════════════════ */
const REGION_LABELS = { "AM": "Americas", "EU": "Europe", "AS/SIS/ESEA": "Asia" };
const REGION_CLASS = { "AM": "region-am", "EU": "region-eu", "AS/SIS/ESEA": "region-as" };
const OLD_URL = "file/old.xlsx";
const NEW_URL = "file/ranking.xlsx";
const HISTORY_URL = "file/history.xlsx";
const ARCHIVE_DIR = "file/archives/";
const MANIFEST_URL = ARCHIVE_DIR + "manifest.json";

/* ════════════════════════════════════════
   GLOBAL STATE
════════════════════════════════════════ */
let oldRowsCache = null;
let newRowsCache = null;
let hltvHeaders = [];
let hltvRows = [];
let currentNewTeams = null;
let lastCombined = null;
const charts = {};

/* Pro Analyses state */
let historyData = [];
let paFilteredTeam = '';
let paFilteredOpponent = '';

/* Team logos */
let teamLogosMap = {};

/* Time Navigator (old.xlsx / ranking.xlsx / archived snapshots) state */
let archiveManifest = null;   // raw manifest.json contents
let navChain = [];            // [{ label, url, date, kind }] ordered oldest → newest
let navFromIndex = null;      // index within navChain currently shown as "Previous"
let navToIndex = null;        // index within navChain currently shown as "Actual"

/* Dashboard state */
let dashSortDesc = true;

/* Change History state */
let changeHistory = [];
const CHANGE_HISTORY_KEY = 'nexus_change_history';

/* ════════════════════════════════════════
   SIDEBAR TOGGLE
════════════════════════════════════════ */
const sidebarToggle = $('#sidebarToggle');
const appShell = $('#appShell');

sidebarToggle.addEventListener('click', () => {
  appShell.classList.toggle('collapsed');
  sidebarToggle.textContent = appShell.classList.contains('collapsed') ? '▶' : '◀';
  // Give charts time to resize
  setTimeout(() => {
    Object.values(charts).forEach(c => { if (c) c.resize(); });
  }, 350);
});

/* ════════════════════════════════════════
   PANEL SWITCHING
════════════════════════════════════════ */
const PANEL_TITLES = {
  dashboard: 'Dashboard',
  charts: 'Charts & Analytics',
  global: 'Global Ranking',
  region: 'Regional Ranking',
  matchup: 'Matchup',
  legends: 'Legends Hall',
  'pro-analyses': 'Pro Analyses',
  'seeding': 'Seeding',
  'events': 'Events',
  'finances': 'Finances',
  'database': 'Database',
  'events-org': 'Events Organizer'
};

function switchPanel(name) {
  $$('.panel').forEach(p => p.classList.remove('active'));
  $$('.nav-item[data-panel]').forEach(n => n.classList.remove('active'));
  const panel = $(`#panel-${name}`);
  if (panel) panel.classList.add('active');
  const nav = $(`[data-panel="${name}"]`);
  if (nav) nav.classList.add('active');
  $('#pageTitle').textContent = PANEL_TITLES[name] || name;

  // Lazy-load iframes
  const iframe = panel && panel.querySelector('iframe.panel-iframe');
  if (iframe && !iframe.src && iframe.dataset.src) {
    iframe.src = iframe.dataset.src;
  }
}

/* ════════════════════════════════════════
   STATUS HELPERS
════════════════════════════════════════ */
function setStatus(msg, state = 'loading') {
  $('#statusText').textContent = msg;
  $('#statusDot').className = 'status-dot ' + state;
}

/* ════════════════════════════════════════
   CHART DEFAULTS & HELPERS
════════════════════════════════════════ */
Chart.defaults.color = '#8a8275';
Chart.defaults.font.family = "'IBM Plex Mono', monospace";
Chart.defaults.font.size = 11;

function baseOpts(extra = {}) {
  return {
    responsive: true, maintainAspectRatio: false,
    animation: { duration: 300 },
    plugins: {
      legend: { display: false },
      tooltip: {
        backgroundColor: 'rgba(244,241,234,0.96)', titleColor: '#1c1a17', bodyColor: '#4a463e',
        borderColor: 'rgba(140,47,36,0.2)', borderWidth: 1, padding: 10,
        titleFont: { weight: '700', size: 12 }, bodyFont: { size: 11 }
      }
    },
    layout: { padding: 4 },
    scales: {
      x: { ticks: { color: '#8a8275', maxRotation: 35, font: { size: 10 } }, grid: { color: 'rgba(140,47,36,0.04)' } },
      y: { ticks: { color: '#8a8275', font: { size: 10 }, precision: 0 }, grid: { color: 'rgba(140,47,36,0.05)' }, beginAtZero: true }
    },
    ...extra
  };
}

function destroyChart(k) {
  if (charts[k]) { charts[k].destroy(); charts[k] = null; }
}

function toggleViz(id, show, msg = '') {
  const c = document.getElementById(id);
  const s = document.getElementById(id + 'State');
  if (c) c.style.display = show ? 'block' : 'none';
  if (s) { s.style.display = show ? 'none' : 'flex'; if (!show && msg) s.innerHTML = msg; }
}

/* ════════════════════════════════════════
   NUMBER PARSING — handles EU comma decimals
════════════════════════════════════════ */
function num(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  const s = v.toString().trim();
  if (!s) return 0;
  const commas = (s.match(/,/g) || []).length;
  const dots = (s.match(/\./g) || []).length;
  let n;
  if (commas === 1 && dots === 0) n = s.replace(',', '.');
  else if (commas > 1 && dots === 0) n = s.replace(/,/g, '');
  else if (dots === 1 && commas === 0) n = s;
  else if (dots > 1 && commas === 0) n = s.replace(/\./g, '');
  else if (commas >= 1 && dots >= 1) {
    const lc = s.lastIndexOf(','), ld = s.lastIndexOf('.');
    n = lc > ld ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  } else n = s;
  const r = parseFloat(n);
  return isNaN(r) ? 0 : r;
}

function toNumber(v) {
  if (v == null) return NaN;
  if (typeof v === 'number') return v;
  const s = String(v).replace(/\./g, '').replace(/,/g, '.').trim();
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}

function ptsFmt(v) {
  const n = typeof v === 'number' ? v : toNumber(v);
  return Number.isFinite(n) ? Math.trunc(n).toLocaleString() : '0';
}

const avg = arr => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0;

const escHtml = s => String(s ?? '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;')
  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/* ════════════════════════════════════════
   TEAM LOGOS
════════════════════════════════════════ */
const TEAM_LOGO_URL = 'assets/team_logo.json';

async function loadTeamLogos() {
  try {
    const res = await fetch(TEAM_LOGO_URL, { cache: 'no-store' });
    if (!res.ok) { console.warn('team_logo.json not found (' + res.status + ')'); return; }
    const arr = await res.json();
    if (!Array.isArray(arr)) return;
    arr.forEach(entry => {
      if (entry.name && entry.logo) {
        teamLogosMap[entry.name.trim().toLowerCase()] = entry.logo.trim();
      }
    });
    console.log(`Loaded ${Object.keys(teamLogosMap).length} team logo(s).`);
  } catch (e) { console.warn('Failed to load team logos:', e); }
}

function getTeamLogoUrl(teamName) {
  if (!teamName) return '';
  return teamLogosMap[teamName.trim().toLowerCase()] || '';
}

function teamLogoImgHtml(teamName, sizeClass) {
  const url = getTeamLogoUrl(teamName);
  if (!url) return '';
  const cls = sizeClass || 'team-logo-sm';
  return `<img class="team-logo ${cls}" src="${escHtml(url)}" alt="" loading="lazy" onerror="this.style.display='none'">`;
}

/* ════════════════════════════════════════
   EXCEL LOADING
════════════════════════════════════════ */
async function readExcel(url) {
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error('HTTP ' + res.status);
  const lastModified = res.headers.get('Last-Modified') || null;
  const data = await res.arrayBuffer();
  const wb = XLSX.read(data, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return {
    json: XLSX.utils.sheet_to_json(ws, { defval: '', raw: true }),
    raw: XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' }),
    lastModified
  };
}

async function loadAll(force = false) {
  const bust = force ? `?v=${Date.now()}` : '';
  setStatus('Loading old.xlsx…', 'loading');
  const oldData = await readExcel(OLD_URL + bust);
  oldRowsCache = oldData.json;
  recordChange('old.xlsx', oldData.lastModified, force ? 'reload' : 'load');

  setStatus('Loading ranking.xlsx…', 'loading');
  const newData = await readExcel(NEW_URL + bust);
  newRowsCache = newData.json;
  recordChange('ranking.xlsx', newData.lastModified, force ? 'reload' : 'load');

  hltvHeaders = (newData.raw[0] || []).map(h => (h ?? '').toString().trim());
  hltvRows = newData.raw.slice(1).filter(r => r.some(c => c !== undefined && c !== ''));

  setStatus('Loading history.xlsx…', 'loading');
  try {
    const histData = await readExcel(HISTORY_URL + bust);
    historyData = parseHistoryData(histData.json);
    recordChange('history.xlsx', histData.lastModified, force ? 'reload' : 'load');
  } catch (e) {
    console.warn('history.xlsx not found or failed to load:', e);
    historyData = [];
  }

  setStatus('Processing…', 'loading');
  computeVRS();
  renderDashboard();
  renderHltvCharts();
  initProAnalyses();

  const histMsg = historyData.length ? ` · ${historyData.length} matches` : ' · history.xlsx not found';
  setStatus(`Loaded — ${hltvRows.length} teams${histMsg}`, 'ok');

  enableFilters(true);
  enableRegion(true);
  enablePredictor(true);
  $('#btnDownloadCsv').disabled = false;

  renderChangeHistory();
}

/* ════════════════════════════════════════
   TIME NAVIGATOR
   Walks the full snapshot chain:
   r0.xlsx → r1.xlsx → … → r[N].xlsx → old.xlsx → ranking.xlsx
   Two modes share the same chain:
     - Arrows (←/→): step one snapshot at a time, always comparing
       two ADJACENT entries (Previous = Actual - 1).
     - Custom range (the two <select> dropdowns): compare ANY two
       entries, e.g. r1.xlsx vs ranking.xlsx, to see a team's
       accumulated movement across a whole stretch of time.
   navFromIndex / navToIndex always reflect whichever pair is
   currently on screen, regardless of which mode produced it.
════════════════════════════════════════ */
function todayIso() {
  const d = new Date();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${d.getFullYear()}-${mm}-${dd}`;
}

function fmtDateDisplay(iso) {
  if (!iso) return '—';
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
}

function navOptionLabel(entry) {
  const suffix = entry.kind === 'live' ? ' (today)' : '';
  return `${entry.label} · ${fmtDateDisplay(entry.date)}${suffix}`;
}

function populateNavSelects() {
  const fromSel = $('#navFromSelect'), toSel = $('#navToSelect');
  if (!fromSel || !toSel) return;
  const optionsHtml = navChain.map((entry, i) => `<option value="${i}">${escHtml(navOptionLabel(entry))}</option>`).join('');
  fromSel.innerHTML = optionsHtml;
  toSel.innerHTML = optionsHtml;
}

function syncNavSelectsToIndices(fromIdx, toIdx) {
  const fromSel = $('#navFromSelect'), toSel = $('#navToSelect');
  if (fromSel) fromSel.value = String(fromIdx);
  if (toSel) toSel.value = String(toIdx);
}

async function loadArchiveManifest() {
  try {
    const res = await fetch(MANIFEST_URL, { cache: 'no-store' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    archiveManifest = await res.json();
  } catch (e) {
    console.warn('manifest.json not found or failed to load:', e);
    archiveManifest = { archives: [], old: { date: null } };
  }

  const chain = (archiveManifest.archives || [])
    .map(entry => ({ label: entry.file, url: ARCHIVE_DIR + entry.file, date: entry.date, kind: 'archive' }))
    .sort((a, b) => (a.date || '').localeCompare(b.date || ''));

  chain.push({ label: 'old.xlsx', url: OLD_URL, date: archiveManifest.old?.date || null, kind: 'old' });
  chain.push({ label: 'ranking.xlsx', url: NEW_URL, date: todayIso(), kind: 'live' });

  navChain = chain;
  navToIndex = navChain.length - 1;   // old.xlsx/ranking.xlsx pair is already loaded by loadAll()
  navFromIndex = navToIndex > 0 ? navToIndex - 1 : 0;

  populateNavSelects();
  updateNavUI();
}

/* Arrow-driven adjacent stepping: Previous is always Actual - 1 */
async function loadPairAt(index) {
  const toIdx = index;
  const fromIdx = toIdx > 0 ? toIdx - 1 : null;
  await loadSnapshotPair(fromIdx, toIdx, { allowEmpty: true });
}

/* Dropdown-driven custom range: any two entries, in any order */
async function loadCustomPair(rawFromIdx, rawToIdx) {
  if (rawFromIdx === rawToIdx) { setStatus('Pick two different snapshots to compare.', 'err'); return; }
  const fromIdx = Math.min(rawFromIdx, rawToIdx);
  const toIdx = Math.max(rawFromIdx, rawToIdx);
  await loadSnapshotPair(fromIdx, toIdx, { allowEmpty: false, label: 'custom range' });
}

/* Shared loader: fetches the "to" snapshot into newRowsCache and the
   "from" snapshot (if any) into oldRowsCache, then reuses the existing
   computeVRS()/render pipeline exactly as the live old/ranking pair does. */
async function loadSnapshotPair(fromIdx, toIdx, { allowEmpty = false, label = null } = {}) {
  const to = navChain[toIdx];
  if (!to) return;
  const from = fromIdx !== null && fromIdx !== undefined ? navChain[fromIdx] : null;

  setNavBusy(true);
  try {
    setStatus(`Loading ${to.label}…`, 'loading');
    const toData = await readExcel(to.url + '?v=' + Date.now());
    newRowsCache = toData.json;
    hltvHeaders = (toData.raw[0] || []).map(h => (h ?? '').toString().trim());
    hltvRows = toData.raw.slice(1).filter(r => r.some(c => c !== undefined && c !== ''));
    recordChange(to.label, toData.lastModified, 'nav');

    if (from) {
      setStatus(`Loading ${from.label}…`, 'loading');
      const fromData = await readExcel(from.url + '?v=' + Date.now());
      oldRowsCache = fromData.json;
      recordChange(from.label, fromData.lastModified, 'nav');
    } else if (allowEmpty) {
      oldRowsCache = []; // oldest snapshot has no predecessor — every team renders as NEW
    }

    computeVRS();
    renderDashboard();
    renderHltvCharts();
    renderChangeHistory();

    navFromIndex = from ? fromIdx : (toIdx > 0 ? toIdx - 1 : 0);
    navToIndex = toIdx;
    syncNavSelectsToIndices(navFromIndex, navToIndex);
    updateNavUI();

    const suffix = label ? ` (${label})` : '';
    setStatus(`Viewing ${to.label}` + (from ? ` vs ${from.label}${suffix}` : ' (oldest snapshot — no predecessor)'), 'ok');
  } catch (e) {
    console.error(e);
    setStatus(`Error loading comparison`, 'err');
  } finally {
    setNavBusy(false);
  }
}

function setNavBusy(busy) {
  const prevBtn = $('#navPrevBtn'), nextBtn = $('#navNextBtn'), resetBtn = $('#navResetBtn'), compareBtn = $('#navCustomCompareBtn');
  if (prevBtn) prevBtn.disabled = busy || navToIndex <= 0;
  if (nextBtn) nextBtn.disabled = busy || navToIndex >= navChain.length - 1;
  if (resetBtn) resetBtn.disabled = busy;
  if (compareBtn) compareBtn.disabled = busy;
}

function updateNavUI() {
  const bar = $('#timeNavBar'), customBar = $('#timeNavCustomBar');
  if (!bar) return;
  if (!navChain.length || navToIndex === null) { bar.style.display = 'none'; if (customBar) customBar.style.display = 'none'; return; }
  bar.style.display = '';
  if (customBar) customBar.style.display = '';

  const actual = navChain[navToIndex];
  const prev = navFromIndex !== null ? navChain[navFromIndex] : null;

  $('#navActualLabel').textContent = actual ? navOptionLabel(actual) : '—';
  $('#navPrevLabel').textContent = prev ? navOptionLabel(prev) : '— (oldest snapshot)';

  syncNavSelectsToIndices(navFromIndex ?? 0, navToIndex);
  setNavBusy(false);
}

/* ════════════════════════════════════════
   VRS DATA HELPERS
════════════════════════════════════════ */
function normalizeHeader(h) { return String(h || '').trim().toLowerCase(); }

function getColumn(row, hm, desired) {
  const cands = {
    team: ['team', 'time', 'equipe'], points: ['points', 'pontos'], tier: ['tier'],
    region: ['region', 'região', 'regiao'],
    victories: ['victories', 'vitorias', 'vitórias'],
    streak: ['streak', 'sequencia', 'sequência'],
    loses: ['loses', 'losses', 'derrotas', 'perdas'],
    prestige: ['prestige', 'prestígio'], majors: ['majors'],
    trophies: ['tournaments trophies', 'trophies', 'troféus']
  }[desired];
  for (const c of cands) { const k = hm[c]; if (k !== undefined) return row[k]; }
  return undefined;
}

function buildHeaderMap(rows) {
  const m = {};
  const s = rows[0] || {};
  Object.keys(s).forEach(k => { m[normalizeHeader(k)] = k; });
  return m;
}

function buildRanking(rows) {
  const hm = buildHeaderMap(rows);
  const teams = [];
  for (const r of rows) {
    const team = String(getColumn(r, hm, 'team') ?? '').trim();
    if (!team) continue;
    const points = toNumber(getColumn(r, hm, 'points'));
    teams.push({
      team, points: Number.isFinite(points) ? points : 0,
      tier: String(getColumn(r, hm, 'tier') ?? '').trim(),
      region: String(getColumn(r, hm, 'region') ?? '').trim(),
      victories: toNumber(getColumn(r, hm, 'victories')) || 0,
      streak: toNumber(getColumn(r, hm, 'streak')) || 0,
      loses: toNumber(getColumn(r, hm, 'loses')) || 0,
      prestige: toNumber(getColumn(r, hm, 'prestige')) || 0,
      majors: toNumber(getColumn(r, hm, 'majors')) || 0,
      trophies: toNumber(getColumn(r, hm, 'trophies')) || 0
    });
  }
  teams.sort((a, b) => b.points !== a.points ? b.points - a.points : a.team.localeCompare(b.team, 'en'));
  teams.forEach((t, i) => t.pos = i + 1);
  return teams;
}

function mapPositions(ts) { const m = new Map(); ts.forEach(t => m.set(t.team.toLowerCase(), t.pos)); return m; }
function mapPoints(ts) { const m = new Map(); ts.forEach(t => m.set(t.team.toLowerCase(), t.points)); return m; }

function computeDelta(o, n) {
  const d = o - n;
  if (d === 0) return { text: '—', cls: 'flat' };
  return d > 0 ? { text: '+' + d, cls: 'up' } : { text: '' + d, cls: 'down' };
}

function computePointsDiff(o, n) {
  const d = Math.trunc(n) - Math.trunc(o);
  if (d === 0) return { text: '—', cls: 'flat' };
  return d > 0 ? { text: '+' + d.toLocaleString(), cls: 'up' } : { text: d.toLocaleString(), cls: 'down' };
}

/* ════════════════════════════════════════
   DISPLAY HELPERS
════════════════════════════════════════ */
function regionLabel(v) { return REGION_LABELS[v] || v || '—'; }
function regionCls(v) { return REGION_CLASS[v] || ''; }

function tierBadgeClass(t) {
  if (!t) return '';
  const l = t.toLowerCase();
  if (l.includes('s')) return 'tier-s';
  if (l.includes('1')) return 'tier-1';
  if (l.includes('2')) return 'tier-2';
  return 'tier-3';
}

function rankBadgeHtml(p) {
  if (p === 1) return '<span class="rank-badge rank-badge-1">1</span>';
  if (p === 2) return '<span class="rank-badge rank-badge-2">2</span>';
  if (p === 3) return '<span class="rank-badge rank-badge-3">3</span>';
  return `<span class="rank-badge rank-badge-n">${p}</span>`;
}

function teamNameHtml(n, p) {
  const logo = teamLogoImgHtml(n, 'team-logo-md');
  if (p === 1) return `<span class="team-name-with-logo"><span class="team-name-animated team-name-1">${logo}${escHtml(n)}</span></span>`;
  if (p === 2) return `<span class="team-name-with-logo"><span class="team-name-animated team-name-2">${logo}${escHtml(n)}</span></span>`;
  if (p === 3) return `<span class="team-name-with-logo"><span class="team-name-animated team-name-3">${logo}${escHtml(n)}</span></span>`;
  if (p <= 10) return `<span class="team-name-with-logo"><span class="team-name-animated team-name-top">${logo}${escHtml(n)}</span></span>`;
  return `<span class="team-name-with-logo"><span class="team-name-default">${logo}${escHtml(n)}</span></span>`;
}

function rowClass(p) {
  if (p === 1) return 'rank-row rank-1';
  if (p === 2) return 'rank-row rank-2';
  if (p === 3) return 'rank-row rank-3';
  if (p <= 10) return 'rank-row rank-top';
  return '';
}

function streakHtml(s) {
  const v = Number(s) || 0;
  return v === 0 ? '<span class="streak-badge cold">—</span>' : `<span class="streak-badge hot">▲${v}</span>`;
}

/* ════════════════════════════════════════
   VRS COMPUTE
════════════════════════════════════════ */
function computeVRS() {
  if (!oldRowsCache || !newRowsCache) return;

  const oldTeams = buildRanking(oldRowsCache);
  const newTeams = buildRanking(newRowsCache);
  const oldPosMap = mapPositions(oldTeams);
  const oldPtsMap = mapPoints(oldTeams);

  const combined = newTeams.map(t => {
    const k = t.team.toLowerCase();
    const oPos = oldPosMap.get(k);
    const oPts = oldPtsMap.get(k);
    const pd = oPos ? computeDelta(oPos, t.pos) : { text: 'NEW', cls: 'new' };
    const ptd = oPts !== undefined ? computePointsDiff(oPts, t.points) : { text: 'NEW', cls: 'new' };
    return {
      ...t, oldPos: oPos ?? null, newPos: t.pos,
      deltaText: oPos ? pd.text : 'NEW', deltaCls: oPos ? pd.cls : 'new',
      pointsDiffText: oPts !== undefined ? ptd.text : 'NEW',
      pointsDiffCls: oPts !== undefined ? ptd.cls : 'new'
    };
  });

  let up = 0, down = 0, flat = 0, newly = 0;
  combined.forEach(r => {
    if (r.deltaText === 'NEW') newly++;
    else if (r.deltaText === '—' || r.deltaText === '0') flat++;
    else if (r.deltaText.startsWith('+')) up++;
    else down++;
  });

  lastCombined = combined;
  currentNewTeams = newTeams;

  // Stats
  const avgPts = avg(combined.map(r => r.points));
  ['countTeams', 'sideTeams'].forEach(id => { const e = document.getElementById(id); if (e) e.textContent = combined.length; });
  ['countUp', 'sideUp'].forEach(id => { const e = document.getElementById(id); if (e) e.textContent = up; });
  ['countDown', 'sideDown'].forEach(id => { const e = document.getElementById(id); if (e) e.textContent = down; });
  $('#countFlat').textContent = flat;
  $('#countNew').textContent = newly;
  $('#countTop').textContent = combined[0] ? combined[0].team : '—';
  $('#countAvg').textContent = Math.round(avgPts);
  $('#statsRow').style.display = 'flex';

  updateRegionOptions(combined, '#filterRegion');
  enableFilters(true);
  renderGlobalTable(combined);

  updateRegionOptions(combined, '#regionRankingSelect');
  enableRegion(true);
  recomputeRegionRanking();

  updateMatchTeamOptions(newTeams);
  enablePredictor(true);
  renderMatchPredictor();

  // Populate seeding team list
  seedTeamData = [];
  populateSeedTeamList();
}

/* ════════════════════════════════════════
   GLOBAL TABLE (VRS)
════════════════════════════════════════ */
function applyFilters(c) {
  const txt = $('#filterText').value.trim().toLowerCase();
  const reg = $('#filterRegion').value;
  const tier = $('#filterTier').value;
  return c.filter(r =>
    (!txt || r.team.toLowerCase().includes(txt)) &&
    (!reg || r.region === reg) &&
    (!tier || r.tier === tier)
  );
}

function renderGlobalTable(combined) {
  const tb = $('#tbody');
  const rows = applyFilters(combined);
  if (!rows.length) { tb.innerHTML = '<tr><td colspan="9" class="empty-state">No results for current filters.</td></tr>'; return; }
  tb.innerHTML = rows.map(r => {
    const o2n = (r.oldPos || '—') + ' → ' + r.newPos;
    const dc = r.deltaText === 'NEW' ? '<span class="delta new">✦ NEW</span>' : `<span class="delta ${r.deltaCls}">${r.deltaText}</span>`;
    const pc = r.pointsDiffText === 'NEW' ? '<span class="delta new">✦ NEW</span>' : `<span class="delta ${r.pointsDiffCls}">${r.pointsDiffText}</span>`;
    const tc = tierBadgeClass(r.tier);
    const th = r.tier ? `<span class="tier-badge ${tc}">${escHtml(r.tier)}</span>` : '<span class="text-muted text-xs">—</span>';
    const rc = regionCls(r.region);
    const rh = r.region ? `<span class="region-tag ${rc}"><span class="region-dot"></span>${escHtml(regionLabel(r.region))}</span>` : '<span class="text-muted">—</span>';
    return `<tr class="${rowClass(r.newPos)}">
      <td>${rankBadgeHtml(r.newPos)}</td><td>${teamNameHtml(r.team, r.newPos)}</td>
      <td class="right mono">${ptsFmt(r.points)}</td><td class="right">${pc}</td>
      <td class="right">${streakHtml(r.streak)}</td><td>${rh}</td><td>${th}</td>
      <td class="right">${dc}</td><td class="right mono text-xs text-muted">${o2n}</td>
    </tr>`;
  }).join('');
}

/* ════════════════════════════════════════
   DASHBOARD — responds to metric filter
════════════════════════════════════════ */
function renderDashboard() {
  if (!hltvRows.length) return;

  const metric = $('#dashMetric').value;
  const idxTeam = hltvHeaders.indexOf('Team');
  const idxRegion = hltvHeaders.indexOf('Region');
  const idxMetric = hltvHeaders.indexOf(metric);

  // Region chart
  const buckets = { AM: [], EU: [], 'AS/SIS/ESEA': [] };
  hltvRows.forEach(r => {
    const reg = (r[idxRegion] || '').toString().trim();
    if (buckets[reg] !== undefined && idxMetric >= 0) buckets[reg].push(num(r[idxMetric]));
  });
  renderRegionChart(['Americas', 'Europe', 'Asia'], [avg(buckets.AM), avg(buckets.EU), avg(buckets['AS/SIS/ESEA'])]);

  // Top 10 chart — by selected metric
  if (idxMetric >= 0) {
    const isAsc = false; // always desc for top
    const sorted = [...hltvRows].sort((a, b) => num(b[idxMetric]) - num(a[idxMetric])).slice(0, 10);
    renderTopChart(sorted, idxTeam, idxMetric, metric);
    $('#top10Subtitle').textContent = 'by ' + metric;
  }

  // Top 20 table — sorted by selected metric
  renderDashTable(metric);
}

function renderRegionChart(labels, values) {
  destroyChart('regionChart');
  const el = document.getElementById('regionChart');
  charts.regionChart = new Chart(el.getContext('2d'), {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        data: values.map(v => +v.toFixed(2)),
        backgroundColor: ['rgba(140,47,36,0.7)', 'rgba(168,93,58,0.7)', 'rgba(63,125,82,0.7)'],
        borderColor: ['#8c2f24', '#a85d3a', '#3f7d52'],
        borderWidth: 1, borderRadius: 4, maxBarThickness: 52
      }]
    },
    options: baseOpts()
  });
  toggleViz('regionChart', true);
}

function renderTopChart(data, ti, mi, metric) {
  destroyChart('topChart');
  const el = document.getElementById('topChart');
  const ctx = el.getContext('2d');
  const g = ctx.createLinearGradient(0, 0, 0, 250);
  g.addColorStop(0, 'rgba(169,132,46,0.9)');
  g.addColorStop(1, 'rgba(150,110,40,0.7)');
  charts.topChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: data.map(r => (r[ti] || '?').toString()),
      datasets: [{ label: metric, data: data.map(r => num(r[mi])), backgroundColor: g, borderColor: 'rgba(169,132,46,0.4)', borderWidth: 1, borderRadius: 4, maxBarThickness: 48 }]
    },
    options: { ...baseOpts(), scales: { x: { ticks: { color: '#8a8275', font: { size: 10 }, maxRotation: 35 }, grid: { display: false } }, y: { ticks: { color: '#8a8275', font: { size: 10 }, precision: 0 }, grid: { color: 'rgba(140,47,36,0.05)' }, beginAtZero: true } } }
  });
  toggleViz('topChart', true);
}

function renderDashTable(metric) {
  if (!hltvRows.length) return;
  const idxTeam = hltvHeaders.indexOf('Team');
  const idxRegion = hltvHeaders.indexOf('Region');
  const idxTier = hltvHeaders.indexOf('Tier');
  const idxPts = hltvHeaders.indexOf('Points');
  const idxVictories = hltvHeaders.indexOf('Victories');
  const idxLoses = hltvHeaders.indexOf('Loses');
  const idxStreak = hltvHeaders.indexOf('Streak');
  const idxPrestige = hltvHeaders.indexOf('Prestige');
  const idxMajors = hltvHeaders.indexOf('Majors');
  const idxTrophies = hltvHeaders.indexOf('Tournaments Trophies');
  const idxMetric = hltvHeaders.indexOf(metric);

  // Sort ALL teams by selected metric
  const sorted = idxMetric >= 0
    ? [...hltvRows].sort((a, b) => dashSortDesc ? num(b[idxMetric]) - num(a[idxMetric]) : num(a[idxMetric]) - num(b[idxMetric]))
    : [...hltvRows];

  const sortLabel = dashSortDesc ? '↓ Desc' : '↑ Asc';
  const sortBtn = document.getElementById('btnDashSortOrder');
  if (sortBtn) sortBtn.textContent = sortLabel;

  $('#dashTableSubtitle').textContent = 'sorted by ' + metric + (dashSortDesc ? ' (desc)' : ' (asc)');

  // Determine which column is the active metric for highlighting
  const metricKey = metric;

  const table = document.getElementById('dashTable');
  let html = `<thead><tr>
    <th>#</th><th>Team</th>
    <th class="right dash-col-header ${metricKey === 'Points' ? 'dash-col-active' : ''}" data-dash-col="Points">Pts</th>
    <th class="right dash-col-header ${metricKey === 'Victories' ? 'dash-col-active' : ''}" data-dash-col="Victories">W</th>
    <th class="right dash-col-header ${metricKey === 'Loses' ? 'dash-col-active' : ''}" data-dash-col="Loses">L</th>
    <th class="right dash-col-header ${metricKey === 'Streak' ? 'dash-col-active' : ''}" data-dash-col="Streak">Streak</th>
    <th>Region</th><th>Tier</th>
    <th class="right dash-col-header ${metricKey === 'Prestige' ? 'dash-col-active' : ''}" data-dash-col="Prestige">Prestige</th>
    <th class="right dash-col-header ${metricKey === 'Majors' ? 'dash-col-active' : ''}" data-dash-col="Majors">Majors</th>
    <th class="right dash-col-header ${metricKey === 'Tournaments Trophies' ? 'dash-col-active' : ''}" data-dash-col="Tournaments Trophies">Trophies</th>
  </tr></thead><tbody>`;

  sorted.forEach((r, i) => {
    const teamName = (r[idxTeam] || '?').toString();
    const ptsVal = idxPts >= 0 ? num(r[idxPts]) : 0;
    const wVal = idxVictories >= 0 ? num(r[idxVictories]) : 0;
    const lVal = idxLoses >= 0 ? num(r[idxLoses]) : 0;
    const sVal = idxStreak >= 0 ? num(r[idxStreak]) : 0;
    const reg = idxRegion >= 0 ? (r[idxRegion] || '').toString().trim() : '';
    const tier = idxTier >= 0 ? (r[idxTier] || '').toString().trim() : '';
    const presVal = idxPrestige >= 0 ? num(r[idxPrestige]) : 0;
    const majVal = idxMajors >= 0 ? num(r[idxMajors]) : 0;
    const tropVal = idxTrophies >= 0 ? num(r[idxTrophies]) : 0;

    const rc = regionCls(reg);
    const rh = reg ? `<span class="region-tag ${rc}"><span class="region-dot"></span>${escHtml(regionLabel(reg))}</span>` : '—';
    const tc = tierBadgeClass(tier);
    const th = tier ? `<span class="tier-badge ${tc}">${escHtml(tier)}</span>` : '—';
    const pos = i + 1;

    // Streak display with color
    const streakDisplay = sVal === 0
      ? '<span class="streak-badge cold">—</span>'
      : `<span class="streak-badge hot">▲${sVal}</span>`;

    // Highlight class for the active metric column
    const hlPts = metricKey === 'Points' ? ' dash-cell-active' : '';
    const hlW = metricKey === 'Victories' ? ' dash-cell-active' : '';
    const hlL = metricKey === 'Loses' ? ' dash-cell-active' : '';
    const hlS = metricKey === 'Streak' ? ' dash-cell-active' : '';
    const hlPres = metricKey === 'Prestige' ? ' dash-cell-active' : '';
    const hlMaj = metricKey === 'Majors' ? ' dash-cell-active' : '';
    const hlTrop = metricKey === 'Tournaments Trophies' ? ' dash-cell-active' : '';

    html += `<tr class="${rowClass(pos)}">
      <td>${rankBadgeHtml(pos)}</td>
      <td>${teamNameHtml(teamName, pos)}</td>
      <td class="right mono${hlPts}">${ptsVal.toFixed(0)}</td>
      <td class="right mono${hlW}">${wVal}</td>
      <td class="right mono${hlL}">${lVal}</td>
      <td class="center${hlS}">${streakDisplay}</td>
      <td>${rh}</td><td>${th}</td>
      <td class="right mono${hlPres}">${presVal.toFixed(2)}</td>
      <td class="right mono${hlMaj}">${majVal}</td>
      <td class="right mono${hlTrop}">${tropVal}</td>
    </tr>`;
  });

  html += '</tbody>';
  table.innerHTML = html;
  table.style.display = 'table';
  document.getElementById('dashTableState').style.display = 'none';
}

/* ════════════════════════════════════════
   HLTV CHARTS
════════════════════════════════════════ */
function renderHltvCharts() {
  if (!hltvRows.length) return;
  const reg = $('#chartRegionFilter').value;
  const idxTeam = hltvHeaders.indexOf('Team');
  const idxRegion = hltvHeaders.indexOf('Region');
  const pool = hltvRows.filter(r => reg === 'Global' || (r[idxRegion] || '').toString().trim() === reg);
  const idx = h => hltvHeaders.indexOf(h);

  const render = (col, chartId, bg, border) => {
    const i = idx(col);
    if (i >= 0) { const t = [...pool].sort((a, b) => num(b[i]) - num(a[i])).slice(0, 10); renderMiniBar(chartId, t.map(r => (r[idxTeam] || '?').toString()), t.map(r => num(r[i])), bg, border); }
    else toggleViz(chartId, false, 'No data');
  };

  render('Streak', 'streakChart', 'rgba(168,93,58,0.75)', '#a85d3a');
  render('Victories', 'victoriesChart', 'rgba(63,125,82,0.75)', '#3f7d52');
  render('Prestige', 'prestigeChart', 'rgba(169,132,46,0.8)', '#a9842e');

  const iPts = idx('Points');
  if (iPts >= 0) { const t = [...pool].sort((a, b) => num(b[iPts]) - num(a[iPts])).slice(0, 20); renderMiniLine('comparisonChart', t.map(r => (r[idxTeam] || '?').toString()), t.map(r => num(r[iPts])), '#8c2f24', 'rgba(140,47,36,0.08)', '#c08a5a', 'Points'); }
  else toggleViz('comparisonChart', false, 'No data');

  const iMaj = idx('Majors');
  if (iMaj >= 0) { const t = [...pool].sort((a, b) => num(b[iMaj]) - num(a[iMaj])).slice(0, 20); renderMiniLine('majorsChart', t.map(r => (r[idxTeam] || '?').toString()), t.map(r => num(r[iMaj])), '#ff6b35', 'rgba(255,107,53,0.08)', '#fb923c', 'Majors'); }
  else toggleViz('majorsChart', false, 'No data');

  populatePieSelect();
  renderTeamPie();
  renderLegendsTable();
}

function renderMiniBar(id, labels, values, bg, border) {
  destroyChart(id);
  const el = document.getElementById(id);
  charts[id] = new Chart(el.getContext('2d'), {
    type: 'bar',
    data: { labels, datasets: [{ data: values, backgroundColor: bg, borderColor: border, borderWidth: 1, borderRadius: 4, maxBarThickness: 42 }] },
    options: { ...baseOpts(), scales: { x: { ticks: { color: '#8a8275', font: { size: 10 }, maxRotation: 35 }, grid: { display: false } }, y: { ticks: { color: '#8a8275', font: { size: 10 }, precision: 0 }, grid: { color: 'rgba(140,47,36,0.04)' }, beginAtZero: true } } }
  });
  el.style.display = 'block';
  const s = document.getElementById(id + 'State'); if (s) s.style.display = 'none';
}

function renderMiniLine(id, labels, values, lineColor, fillColor, pointColor, label) {
  destroyChart(id);
  const el = document.getElementById(id);
  charts[id] = new Chart(el.getContext('2d'), {
    type: 'line',
    data: { labels, datasets: [{ label, data: values, borderColor: lineColor, backgroundColor: fillColor, pointBackgroundColor: pointColor, pointRadius: 3, pointHoverRadius: 5, tension: 0.28, fill: true, borderWidth: 2 }] },
    options: { ...baseOpts(), plugins: { ...baseOpts().plugins, legend: { display: true, labels: { color: '#8a8275', font: { size: 10 } } } } }
  });
  el.style.display = 'block';
  const s = document.getElementById(id + 'State'); if (s) s.style.display = 'none';
}

/* ════════════════════════════════════════
   PIE CHART
════════════════════════════════════════ */
function populatePieSelect() {
  const sel = $('#teamPieSelect');
  const idxTeam = hltvHeaders.indexOf('Team');
  const names = hltvRows.map(r => (r[idxTeam] || '').toString()).filter(Boolean).sort();
  sel.innerHTML = '<option value="">— Select —</option>' + names.map(n => `<option value="${escHtml(n)}">${escHtml(n)}</option>`).join('');
}

function renderTeamPie() {
  const sel = $('#teamPieSelect');
  if (!sel.value) { toggleViz('teamPieChart', false, 'Select a team'); return; }
  const idxTeam = hltvHeaders.indexOf('Team'), idxV = hltvHeaders.indexOf('Victories'), idxL = hltvHeaders.indexOf('Loses');
  const row = hltvRows.find(r => (r[idxTeam] || '').toString() === sel.value);
  if (!row) { toggleViz('teamPieChart', false, 'Team not found'); return; }
  const w = num(row[idxV]), l = num(row[idxL]);
  destroyChart('teamPieChart');
  const el = document.getElementById('teamPieChart');
  charts.teamPieChart = new Chart(el.getContext('2d'), {
    type: 'doughnut',
    data: { labels: ['Wins', 'Losses'], datasets: [{ data: [w, l], backgroundColor: ['rgba(63,125,82,0.8)', 'rgba(163,57,44,0.8)'], borderColor: ['#3f7d52', '#a3392c'], borderWidth: 1 }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: true, labels: { color: '#4a463e', font: { weight: '700' }, padding: 14 } }, tooltip: { backgroundColor: 'rgba(244,241,234,0.96)', titleColor: '#1c1a17', bodyColor: '#4a463e', borderColor: 'rgba(140,47,36,0.2)', borderWidth: 1 } } }
  });
  toggleViz('teamPieChart', true);
}

/* ════════════════════════════════════════
   LEGENDS TABLE
════════════════════════════════════════ */
function renderLegendsTable() {
  const legendRegion = $('#legendsRegionFilter').value;
  const idxTeam = hltvHeaders.indexOf('Team'), idxPts = hltvHeaders.indexOf('Points'), idxRegion = hltvHeaders.indexOf('Region');
  const legends = hltvRows
    .filter(r => { const reg = (r[idxRegion] || '').toString().trim(); return num(r[idxPts]) > 1700 && (legendRegion === 'Global' || reg === legendRegion); })
    .sort((a, b) => num(b[idxPts]) - num(a[idxPts]));
  const table = document.getElementById('legendsTable');
  if (!legends.length) { table.style.display = 'none'; document.getElementById('legendsTableState').style.display = 'flex'; document.getElementById('legendsTableState').innerHTML = 'No Legends (1700+ pts) found.'; return; }
  let html = '<thead><tr><th>#</th><th>Team</th><th>Region</th><th class="right">Points</th></tr></thead><tbody>';
  legends.forEach((r, i) => {
    const cls = i === 0 ? 'legend-gold' : i === 1 ? 'legend-silver' : i === 2 ? 'legend-bronze' : 'legend-other';
    const reg = idxRegion >= 0 ? (r[idxRegion] || '').toString().trim() : '';
    const bc = reg === 'AM' ? 'badge-AM' : reg === 'EU' ? 'badge-EU' : reg.includes('AS') ? 'badge-AS' : '';
    const badge = bc ? `<span class="badge ${bc}">${escHtml(reg)}</span>` : escHtml(reg);
    const rk = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : `#${i + 1}`;
    html += `<tr class="${cls}"><td>${rk}</td><td style="font-weight:700"><span class="team-name-with-logo">${teamLogoImgHtml((r[idxTeam] || '').toString(), 'team-logo-md')}${escHtml((r[idxTeam] || '').toString())}</span></td><td>${badge}</td><td class="right" style="font-weight:800">${num(r[idxPts]).toFixed(2)}</td></tr>`;
  });
  html += '</tbody>';
  table.innerHTML = html; table.style.display = 'table';
  document.getElementById('legendsTableState').style.display = 'none';
}

/* ════════════════════════════════════════
   REGIONAL RANKING (VRS)
════════════════════════════════════════ */
function updateRegionOptions(combined, selId) {
  const sel = $(selId); const cur = sel.value;
  const regs = new Set(); combined.forEach(r => { if (r.region) regs.add(r.region); });
  const order = ['AM', 'EU', 'AS/SIS/ESEA'];
  const list = Array.from(regs).sort((a, b) => { const ia = order.indexOf(a), ib = order.indexOf(b); return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib); });
  const ph = selId === '#filterRegion' ? '<option value="">All Regions</option>' : '';
  sel.innerHTML = ph + list.map(v => `<option value="${escHtml(v)}">${escHtml(regionLabel(v))}</option>`).join('');
  sel.value = list.includes(cur) ? cur : (selId === '#regionRankingSelect' ? list[0] || '' : '');
}

function computeRegionCombined(rv) {
  if (!oldRowsCache || !newRowsCache) return [];
  const oHM = buildHeaderMap(oldRowsCache), nHM = buildHeaderMap(newRowsCache);
  const oF = oldRowsCache.filter(r => String(getColumn(r, oHM, 'region') ?? '').trim() === rv);
  const nF = newRowsCache.filter(r => String(getColumn(r, nHM, 'region') ?? '').trim() === rv);
  const oT = buildRanking(oF), nT = buildRanking(nF);
  const oPM = mapPositions(oT), oPtM = mapPoints(oT);
  return nT.map(t => {
    const k = t.team.toLowerCase(); const oP = oPM.get(k); const oPts = oPtM.get(k);
    const pd = oP ? computeDelta(oP, t.pos) : { text: 'NEW', cls: 'new' };
    const ptd = oPts !== undefined ? computePointsDiff(oPts, t.points) : { text: 'NEW', cls: 'new' };
    return { ...t, oldPos: oP ?? null, newPos: t.pos, deltaText: oP ? pd.text : 'NEW', deltaCls: oP ? pd.cls : 'new', pointsDiffText: oPts !== undefined ? ptd.text : 'NEW', pointsDiffCls: oPts !== undefined ? ptd.cls : 'new' };
  });
}

function recomputeRegionRanking() {
  const v = $('#regionRankingSelect').value;
  if (!v) { $('#tbodyRegion').innerHTML = '<tr><td colspan="8" class="empty-state">Select a region above.</td></tr>'; return; }
  renderRegionTable(computeRegionCombined(v));
}

function renderRegionTable(combined) {
  const tb = $('#tbodyRegion');
  if (!combined || !combined.length) { tb.innerHTML = '<tr><td colspan="8" class="empty-state">No data for selected region.</td></tr>'; return; }
  tb.innerHTML = combined.map(r => {
    const o2n = (r.oldPos || '—') + ' → ' + r.newPos;
    const dc = r.deltaText === 'NEW' ? '<span class="delta new">✦ NEW</span>' : `<span class="delta ${r.deltaCls}">${r.deltaText}</span>`;
    const pc = r.pointsDiffText === 'NEW' ? '<span class="delta new">✦ NEW</span>' : `<span class="delta ${r.pointsDiffCls}">${r.pointsDiffText}</span>`;
    const tc = tierBadgeClass(r.tier); const th = r.tier ? `<span class="tier-badge ${tc}">${escHtml(r.tier)}</span>` : '—';
    return `<tr class="${rowClass(r.newPos)}"><td>${rankBadgeHtml(r.newPos)}</td><td>${teamNameHtml(r.team, r.newPos)}</td><td class="right mono">${ptsFmt(r.points)}</td><td class="right">${pc}</td><td class="right">${streakHtml(r.streak)}</td><td>${th}</td><td class="right">${dc}</td><td class="right mono text-xs text-muted">${o2n}</td></tr>`;
  }).join('');
}

/* ════════════════════════════════════════
   ENABLE/DISABLE CONTROLS
════════════════════════════════════════ */
function enableFilters(on) {
  ['filterRegion', 'filterTier', 'filterText', 'btnClearFilters'].forEach(id => { const e = document.getElementById(id); if (e) e.disabled = !on; });
}
function enableRegion(on) { $('#regionRankingSelect').disabled = !on; }
function enablePredictor(on) {
  ['matchTeamA', 'matchTeamB', 'matchResult'].forEach(id => { const e = document.getElementById(id); if (e) e.disabled = !on; });
}

function updateMatchTeamOptions(teams) {
  const sA = $('#matchTeamA'), sB = $('#matchTeamB');
  const cA = sA.value, cB = sB.value;
  const opts = '<option value="">— Select —</option>' + (teams || []).map(t => `<option value="${escHtml(t.team)}">${escHtml(t.team)}</option>`).join('');
  sA.innerHTML = opts; sB.innerHTML = opts;
  const names = new Set((teams || []).map(t => t.team));
  sA.value = names.has(cA) ? cA : '';
  sB.value = names.has(cB) ? cB : '';
}

/* ════════════════════════════════════════
   ELO ENGINE
════════════════════════════════════════ */
const ELO_K_BASE = 30, ELO_DIVISOR = 400, ELO_STREAK_ALPHA = 0.10, ELO_STREAK_SCALE = 5, ELO_K_CAP_MULT = 1.35;
const ELO_TOURNEY = { regional: 0.05, global: 0.10, major: 0.20 };

function eloExpected(ra, rb) { return 1 / (1 + Math.pow(10, (rb - ra) / ELO_DIVISOR)); }
function tanhL(x) { const e = Math.exp(2 * x); return (e - 1) / (e + 1); }
function streakKBonus(s) { return ELO_STREAK_ALPHA * tanhL((s || 0) / ELO_STREAK_SCALE); }

function getBonusForTeam(lName) {
  if (!newRowsCache) return { type: 'none', remaining: 0 };
  const hm = buildHeaderMap(newRowsCache);
  for (const r of newRowsCache) {
    const t = String(getColumn(r, hm, 'team') ?? '').trim();
    if (t.toLowerCase() !== lName) continue;
    let btk = null, brk = null;
    for (const k of Object.keys(r)) { const nk = normalizeHeader(k); if (nk === 'bonustype') btk = k; if (nk === 'bonusremaining') brk = k; }
    const type = btk ? String(r[btk] ?? '').trim().toLowerCase() : 'none';
    const rem = brk ? toNumber(r[brk]) : 0;
    return { type: ELO_TOURNEY[type] !== undefined ? type : 'none', remaining: Number.isFinite(rem) ? rem : 0 };
  }
  return { type: 'none', remaining: 0 };
}

function kFinalFor(team) {
  const sb = streakKBonus(team.streak || 0);
  const b = getBonusForTeam(team.team.toLowerCase());
  const tb = (b.type !== 'none' && b.remaining > 0) ? (ELO_TOURNEY[b.type] || 0) : 0;
  return Math.min(ELO_K_BASE * (1 + sb) * (1 + tb), ELO_K_BASE * ELO_K_CAP_MULT);
}

function computeHypoRanking(teams, updates) {
  const h = teams.map(t => ({ ...t }));
  for (const ht of h) { const u = updates.get(ht.team.toLowerCase()); if (u) { ht.points = u.newPoints; ht.streak = u.newStreak; } }
  h.sort((a, b) => b.points !== a.points ? b.points - a.points : a.team.localeCompare(b.team, 'en'));
  h.forEach((t, i) => t.pos = i + 1);
  return h;
}

function renderMatchPredictor() {
  const tb = $('#tbodyMatchPredict'), ms = $('#matchStatus'), vm = $('#vsMeter');
  if (!currentNewTeams || !currentNewTeams.length) { tb.innerHTML = '<tr><td colspan="8" class="empty-state">No data loaded.</td></tr>'; ms.innerHTML = '<div class="status-dot"></div><span>Load ranking data first.</span>'; vm.style.display = 'none'; return; }
  const aName = $('#matchTeamA').value, bName = $('#matchTeamB').value, result = $('#matchResult').value;
  if (!aName || !bName) { tb.innerHTML = '<tr><td colspan="8" class="empty-state">Select two teams.</td></tr>'; ms.innerHTML = '<div class="status-dot"></div><span>Select Team A and Team B above.</span>'; vm.style.display = 'none'; return; }
  if (aName === bName) { tb.innerHTML = '<tr><td colspan="8" class="empty-state">Pick two different teams.</td></tr>'; ms.innerHTML = '<div class="status-dot err"></div><span>Team A and Team B must be different.</span>'; vm.style.display = 'none'; return; }
  const byName = new Map(currentNewTeams.map(t => [t.team.toLowerCase(), t]));
  const a = byName.get(aName.toLowerCase()), b = byName.get(bName.toLowerCase());
  if (!a || !b) { tb.innerHTML = '<tr><td colspan="8" class="empty-state">Team not found.</td></tr>'; vm.style.display = 'none'; return; }
  const ra = a.points, rb = b.points, ea = eloExpected(ra, rb), eb = 1 - ea;
  const kA = kFinalFor(a), kB = kFinalFor(b);
  let aD, bD, aNS, bNS;
  if (result === 'A') { aD = kA * (1 - ea); bD = kB * (0 - eb); aNS = (a.streak || 0) + 1; bNS = 0; }
  else { aD = kA * (0 - ea); bD = kB * (1 - eb); aNS = 0; bNS = (b.streak || 0) + 1; }
  const aNP = ra + aD, bNP = rb + bD;
  const updates = new Map([[a.team.toLowerCase(), { newPoints: aNP, newStreak: aNS }], [b.team.toLowerCase(), { newPoints: bNP, newStreak: bNS }]]);
  const hypo = computeHypoRanking(currentNewTeams, updates);
  const oPM = mapPositions(currentNewTeams), nPM = mapPositions(hypo);
  const aOP = oPM.get(a.team.toLowerCase()), bOP = oPM.get(b.team.toLowerCase());
  const aNPos = nPM.get(a.team.toLowerCase()), bNPos = nPM.get(b.team.toLowerCase());
  const aM = computeDelta(aOP, aNPos), bM = computeDelta(bOP, bNPos), aPD = computePointsDiff(ra, aNP), bPD = computePointsDiff(rb, bNP);

  vm.style.display = '';
  const vsALogoHtml = teamLogoImgHtml(a.team, 'team-logo-lg');
  const vsBLogoHtml = teamLogoImgHtml(b.team, 'team-logo-lg');
  $('#vsAName').innerHTML = vsALogoHtml + escHtml(a.team);
  $('#vsBName').innerHTML = vsBLogoHtml + escHtml(b.team);
  $('#vsAPts').textContent = ptsFmt(ra); $('#vsBPts').textContent = ptsFmt(rb);
  const aProb = Math.round(ea * 100), bProb = 100 - aProb;
  $('#vsAProb').textContent = aProb + '%'; $('#vsBProb').textContent = bProb + '%';
  $('#vsProbBarA').style.width = aProb + '%'; $('#vsProbBarB').style.width = bProb + '%';

  function mkRow(team, oP, nP, pts, streak, pd, mv, newPts) {
    const logo = teamLogoImgHtml(team.team, 'team-logo-md');
    return `<tr><td>${rankBadgeHtml(oP)}</td><td><span class="team-name-with-logo">${logo}${escHtml(team.team)}</span></td><td class="right mono">${ptsFmt(pts)}</td><td class="right">${streakHtml(streak)}</td><td class="right"><span class="delta ${pd.cls}">${pd.text}</span></td><td class="right mono">${ptsFmt(newPts)}</td><td class="right"><span class="delta ${mv.cls}">${mv.text}</span></td><td class="right mono text-xs text-muted">${oP} → ${nP}</td></tr>`;
  }
  tb.innerHTML = mkRow(a, aOP, aNPos, ra, aNS, aPD, aM, aNP) + mkRow(b, bOP, bNPos, rb, bNS, bPD, bM, bNP);
  const winner = result === 'A' ? a.team : b.team;
  ms.innerHTML = `<div class="status-dot ok"></div><span>Scenario: <strong>${escHtml(winner)}</strong> wins. ${escHtml(a.team)}: ${aPD.text} pts (${aM.text} positions). ${escHtml(b.team)}: ${bPD.text} pts (${bM.text} positions).</span>`;
}

/* ════════════════════════════════════════
   VRS ANALYSIS — Enhanced
════════════════════════════════════════ */
function runVrsAnalysis() {
  if (!lastCombined || !lastCombined.length) return;
  const c = lastCombined, n = c.length;

  const risers = c.filter(r => r.deltaText !== 'NEW' && r.deltaText.startsWith('+')).sort((a, b) => parseInt(b.deltaText) - parseInt(a.deltaText)).slice(0, 5);
  const fallers = c.filter(r => r.deltaText !== 'NEW' && r.deltaText.startsWith('-')).sort((a, b) => parseInt(a.deltaText) - parseInt(b.deltaText)).slice(0, 5);
  const ptGainers = c.filter(r => r.pointsDiffText !== 'NEW' && r.pointsDiffText.startsWith('+')).sort((a, b) => parseInt(b.pointsDiffText.replace(/[^0-9]/g, '')) - parseInt(a.pointsDiffText.replace(/[^0-9]/g, ''))).slice(0, 5);
  const ptLosers = c.filter(r => r.pointsDiffText !== 'NEW' && r.pointsDiffText.startsWith('-')).sort((a, b) => parseInt(a.pointsDiffText.replace(/[^0-9-]/g, '')) - parseInt(b.pointsDiffText.replace(/[^0-9-]/g, ''))).slice(0, 5);
  const hotStreaks = currentNewTeams ? [...currentNewTeams].sort((a, b) => b.streak - a.streak).filter(t => t.streak >= 2).slice(0, 5) : [];
  const newTeams = c.filter(r => r.deltaText === 'NEW');
  const regionDist = {}; c.forEach(r => { regionDist[r.region] = (regionDist[r.region] || 0) + 1; });
  const tierDist = {}; c.forEach(r => { tierDist[r.tier || 'Unknown'] = (tierDist[r.tier || 'Unknown'] || 0) + 1; });
  const pts = c.map(r => r.points);
  const avgPts = pts.reduce((a, b) => a + b, 0) / n;
  const maxPts = Math.max(...pts), minPts = Math.min(...pts);
  const top1 = c[0], top5 = c.slice(0, 5);
  const gap12 = c.length >= 2 ? Math.trunc(c[0].points - c[1].points) : 0;

  // Regional averages
  const regAvg = {};
  c.forEach(r => { if (!regAvg[r.region]) regAvg[r.region] = []; regAvg[r.region].push(r.points); });
  const regAvgEntries = Object.entries(regAvg).map(([k, v]) => ({ reg: k, avg: avg(v), count: v.length, top: c.filter(t => t.region === k)[0] })).sort((a, b) => b.avg - a.avg);

  // Win rate leaders
  const wrLeaders = currentNewTeams ? [...currentNewTeams]
    .map(t => ({ ...t, games: t.victories + t.loses, wr: (t.victories + t.loses) > 0 ? t.victories / (t.victories + t.loses) : 0 }))
    .filter(t => t.games >= 3)
    .sort((a, b) => b.wr - a.wr).slice(0, 5) : [];

  // Closest battles (tightest point gaps in top 10)
  const closeBattles = [];
  for (let i = 0; i < Math.min(c.length - 1, 15); i++) {
    const diff = Math.abs(c[i].points - c[i + 1].points);
    if (diff < 5) closeBattles.push({ a: c[i], b: c[i + 1], diff, posA: i + 1, posB: i + 2 });
  }
  closeBattles.sort((a, b) => a.diff - b.diff);

  let html = `<div class="ai-analysis-box">
    <div class="ai-header"><div class="ai-icon">🧠</div><div>
      <div class="ai-title">Ranking Analysis Report</div>
      <div class="ai-subtitle mono">Generated from ${n} teams · ${new Date().toLocaleDateString()}</div>
    </div></div>`;

  // Overview
  html += `<div class="ai-section"><div class="ai-section-title">Overview</div><div class="ai-insight-grid">
    <div class="ai-insight-card"><div class="ai-insight-label">Leader</div><div class="ai-insight-value">${escHtml(top1.team)}</div><div class="ai-insight-sub">${ptsFmt(top1.points)} pts · ${escHtml(regionLabel(top1.region))}</div></div>
    <div class="ai-insight-card"><div class="ai-insight-label">#1 vs #2 Gap</div><div class="ai-insight-value" style="color:var(--warn)">${ptsFmt(gap12)} pts</div><div class="ai-insight-sub">${gap12 > 200 ? 'Strong dominance' : gap12 > 50 ? 'Moderate lead' : 'Very tight at the top'}</div></div>
    <div class="ai-insight-card"><div class="ai-insight-label">Avg Points</div><div class="ai-insight-value">${ptsFmt(avgPts)}</div><div class="ai-insight-sub">Range: ${ptsFmt(minPts)} – ${ptsFmt(maxPts)}</div></div>
    <div class="ai-insight-card"><div class="ai-insight-label">Movement</div><div class="ai-insight-value">${risers.length + fallers.length}</div><div class="ai-insight-sub">${risers.length} ↑ risen · ${fallers.length} ↓ fallen</div></div>
    <div class="ai-insight-card"><div class="ai-insight-label">New Entries</div><div class="ai-insight-value" style="color:var(--warn)">${newTeams.length}</div><div class="ai-insight-sub">${newTeams.map(t => t.team).slice(0, 3).join(', ') || 'None'}${newTeams.length > 3 ? '…' : ''}</div></div>
    <div class="ai-insight-card"><div class="ai-insight-label">Hot Streaks</div><div class="ai-insight-value" style="color:var(--good)">${hotStreaks.length}</div><div class="ai-insight-sub">${hotStreaks.map(t => t.team + ' (' + t.streak + ')').slice(0, 2).join(', ') || 'None'}</div></div>
  </div></div>`;

  // Region distribution
  html += `<div class="ai-section"><div class="ai-section-title">Region Distribution</div><div class="ai-insight-grid">
    ${regAvgEntries.map(r => `<div class="ai-insight-card"><div class="ai-insight-label">${escHtml(regionLabel(r.reg))}</div><div class="ai-insight-value">${r.count} teams</div><div class="ai-insight-sub">Avg: ${ptsFmt(r.avg)} pts · Top: ${escHtml(r.top ? r.top.team : '—')}</div></div>`).join('')}
  </div></div>`;

  // Tier distribution
  html += `<div class="ai-section"><div class="ai-section-title">Tier Distribution</div><div class="ai-insight-grid">
    ${Object.entries(tierDist).sort((a, b) => b[1] - a[1]).map(([t, cnt]) => `<div class="ai-insight-card"><div class="ai-insight-label">${escHtml(t)}</div><div class="ai-insight-value">${cnt}</div><div class="ai-insight-sub">${Math.round(cnt / n * 100)}% of field</div></div>`).join('')}
  </div></div>`;

  // Risers
  if (risers.length) html += `<div class="ai-section"><div class="ai-section-title">Biggest Risers</div><div class="ai-list">${risers.map((r, i) => `<div class="ai-list-item"><div class="ai-list-num">${i + 1}</div><div><strong>${escHtml(r.team)}</strong> <span class="delta up" style="margin-left:8px">${r.deltaText} positions</span> <span class="delta up" style="margin-left:8px">${r.pointsDiffText} pts</span><br><span class="text-xs text-muted mono">${escHtml(regionLabel(r.region))} · ${escHtml(r.tier || '—')} · Now #${r.newPos}</span></div></div>`).join('')}</div></div>`;

  // Fallers
  if (fallers.length) html += `<div class="ai-section"><div class="ai-section-title">Biggest Fallers</div><div class="ai-list">${fallers.map((r, i) => `<div class="ai-list-item"><div class="ai-list-num" style="background:rgba(163,57,44,0.1);color:var(--bad)">${i + 1}</div><div><strong>${escHtml(r.team)}</strong> <span class="delta down" style="margin-left:8px">${r.deltaText} positions</span> <span class="delta down" style="margin-left:8px">${r.pointsDiffText} pts</span><br><span class="text-xs text-muted mono">${escHtml(regionLabel(r.region))} · ${escHtml(r.tier || '—')} · Now #${r.newPos}</span></div></div>`).join('')}</div></div>`;

  // Point gainers
  if (ptGainers.length) html += `<div class="ai-section"><div class="ai-section-title">Biggest Point Gains</div><div class="ai-list">${ptGainers.map((r, i) => `<div class="ai-list-item"><div class="ai-list-num" style="background:rgba(63,125,82,0.1);color:var(--good)">${i + 1}</div><div><strong>${escHtml(r.team)}</strong> <span class="delta up">${r.pointsDiffText} pts</span><br><span class="text-xs text-muted mono">${escHtml(regionLabel(r.region))} · #${r.newPos}</span></div></div>`).join('')}</div></div>`;

  // Point losers
  if (ptLosers.length) html += `<div class="ai-section"><div class="ai-section-title">Biggest Point Losses</div><div class="ai-list">${ptLosers.map((r, i) => `<div class="ai-list-item"><div class="ai-list-num" style="background:rgba(163,57,44,0.1);color:var(--bad)">${i + 1}</div><div><strong>${escHtml(r.team)}</strong> <span class="delta down">${r.pointsDiffText} pts</span><br><span class="text-xs text-muted mono">${escHtml(regionLabel(r.region))} · #${r.newPos}</span></div></div>`).join('')}</div></div>`;

  // Hot streaks
  if (hotStreaks.length) html += `<div class="ai-section"><div class="ai-section-title">Active Win Streaks</div><div class="ai-list">${hotStreaks.map((t, i) => `<div class="ai-list-item"><div class="ai-list-num" style="background:rgba(181,134,42,0.1);color:var(--warn)">${i + 1}</div><div><strong>${escHtml(t.team)}</strong> <span style="margin-left:8px">🔥 ${t.streak}-win streak</span><br><span class="text-xs text-muted mono">${escHtml(regionLabel(t.region))} · #${t.pos} · ${ptsFmt(t.points)} pts</span></div></div>`).join('')}</div></div>`;

  // Win rate leaders
  if (wrLeaders.length) html += `<div class="ai-section"><div class="ai-section-title">Win Rate Leaders (min 3 games)</div><div class="ai-list">${wrLeaders.map((t, i) => `<div class="ai-list-item"><div class="ai-list-num">${i + 1}</div><div><strong>${escHtml(t.team)}</strong> <span style="margin-left:8px;color:var(--good);font-weight:700;font-family:var(--font-mono)">${(t.wr * 100).toFixed(1)}%</span><br><span class="text-xs text-muted mono">${t.victories}W-${t.loses}L · ${escHtml(regionLabel(t.region))} · #${t.pos}</span></div></div>`).join('')}</div></div>`;

  // Close battles
  if (closeBattles.length) html += `<div class="ai-section"><div class="ai-section-title">Tightest Battles</div><div class="ai-list">${closeBattles.slice(0, 5).map((b, i) => `<div class="ai-list-item"><div class="ai-list-num" style="background:rgba(181,134,42,0.1);color:var(--warn)">${i + 1}</div><div>#${b.posA} <strong>${escHtml(b.a.team)}</strong> vs #${b.posB} <strong>${escHtml(b.b.team)}</strong> — <span style="color:var(--warn);font-weight:700">${b.diff.toFixed(2)} pts apart</span></div></div>`).join('')}</div></div>`;

  // Key takeaways
  html += `<div class="ai-section"><div class="ai-section-title">Key Takeaways</div><div class="ai-list">${[
    top1 ? `<strong>${escHtml(top1.team)}</strong> leads with ${ptsFmt(top1.points)} points, a ${ptsFmt(gap12)}-point gap over the runner-up.` : null,
    hotStreaks.length ? `${escHtml(hotStreaks[0].team)} is on the hottest streak (${hotStreaks[0].streak} consecutive wins), amplifying ELO K-factor.` : null,
    newTeams.length ? `${newTeams.length} new team${newTeams.length > 1 ? 's' : ''} entered the ranking: ${escHtml(newTeams.map(t => t.team).join(', '))}.` : null,
    risers.length ? `Most impressive riser: <strong>${escHtml(risers[0].team)}</strong>, climbing ${risers[0].deltaText} positions.` : null,
    fallers.length ? `<strong>${escHtml(fallers[0].team)}</strong> suffered the steepest drop: ${fallers[0].deltaText} positions.` : null,
    `Average points: ${ptsFmt(avgPts)}. Spread (${ptsFmt(maxPts)} − ${ptsFmt(minPts)}) = ${ptsFmt(maxPts - minPts)} pts → ${(maxPts - minPts) > 500 ? 'highly stratified' : 'competitive'} field.`,
    regAvgEntries.length ? `Strongest region by avg points: <strong>${escHtml(regionLabel(regAvgEntries[0].reg))}</strong> with ${ptsFmt(regAvgEntries[0].avg)} avg pts.` : null
  ].filter(Boolean).map((txt, i) => `<div class="ai-list-item"><div class="ai-list-num">${i + 1}</div><div>${txt}</div></div>`).join('')}</div></div>`;

  html += `</div>`;
  $('#aiContentVrs').innerHTML = html;
}

/* ════════════════════════════════════════
   DATA INSIGHTS (HLTV typewriter)
════════════════════════════════════════ */
function runHltvAnalysis() {
  if (!hltvRows.length) return;
  const output = $('#aiOutputHltv'); output.innerHTML = '';
  const idx = h => hltvHeaders.indexOf(h);
  const iTeam = idx('Team'), iPts = idx('Points'), iVict = idx('Victories'), iLose = idx('Loses');
  const iStrk = idx('Streak'), iReg = idx('Region'), iTier = idx('Tier'), iMaj = idx('Majors'), iPrest = idx('Prestige');
  const get = (r, i) => i >= 0 ? num(r[i]) : 0;
  const name = r => (r[iTeam] || '?').toString();
  const byPts = [...hltvRows].sort((a, b) => get(b, iPts) - get(a, iPts));
  const t1 = byPts[0], t2 = byPts[1];
  const regions = {};
  hltvRows.forEach(r => { const reg = (iReg >= 0 ? r[iReg] : 'Unknown') || 'Unknown'; if (!regions[reg]) regions[reg] = { teams: [], pts: [] }; regions[reg].teams.push(r); regions[reg].pts.push(get(r, iPts)); });
  const regList = Object.entries(regions).map(([k, v]) => ({ name: k, count: v.teams.length, avgPts: avg(v.pts), topTeam: [...v.teams].sort((a, b) => get(b, iPts) - get(a, iPts))[0] })).sort((a, b) => b.avgPts - a.avgPts);
  const topReg = regList[0];
  const byStrk = iStrk >= 0 ? [...hltvRows].sort((a, b) => get(b, iStrk) - get(a, iStrk)) : [];
  const hotStrk = byStrk[0]; const hotVal = hotStrk ? get(hotStrk, iStrk) : 0;
  const perfect = iLose >= 0 ? hltvRows.filter(r => get(r, iLose) === 0 && get(r, iVict) > 0) : [];
  const wrTeams = hltvRows.map(r => { const w = get(r, iVict), l = get(r, iLose); return { name: name(r), games: w + l, wr: w + l > 0 ? w / (w + l) : 0, wins: w, losses: l }; }).filter(t => t.games >= 3).sort((a, b) => b.wr - a.wr);
  const bestWR = wrTeams[0]; const legends = byPts.filter(r => get(r, iPts) > 1700);
  const gap = t1 && t2 ? (get(t1, iPts) - get(t2, iPts)).toFixed(2) : 0;
  const tierCount = {}; if (iTier >= 0) hltvRows.forEach(r => { const t = r[iTier] || 'Unknown'; tierCount[t] = (tierCount[t] || 0) + 1; });
  let closest = null, closestD = Infinity;
  for (let i = 0; i < Math.min(byPts.length - 1, 9); i++) { const d = Math.abs(get(byPts[i], iPts) - get(byPts[i + 1], iPts)); if (d < closestD) { closestD = d; closest = { a: byPts[i], b: byPts[i + 1], diff: d }; } }

  const L = [];
  L.push('▸ LEADERBOARD SNAPSHOT');
  L.push(`  ${hltvRows.length} teams tracked · ${legends.length} at Legend status (1700+ pts) · ${regList.length} active regions`);
  if (t1) L.push(`  Current #1: ${name(t1)} — ${get(t1, iPts).toFixed(2)} pts${t2 ? ' · leads ' + name(t2) + ' by ' + gap + ' pts' : ''}`);
  L.push('');
  L.push('▸ REGIONAL POWER');
  if (topReg) { L.push(`  ${topReg.name} dominates with avg ${topReg.avgPts.toFixed(2)} pts across ${topReg.count} teams`); if (topReg.topTeam) L.push(`  Region flagship: ${name(topReg.topTeam)} (${get(topReg.topTeam, iPts).toFixed(2)} pts)`); }
  regList.forEach(r => { if (r !== topReg) L.push(`  ${r.name}: ${r.count} teams · avg ${r.avgPts.toFixed(2)} pts · best: ${r.topTeam ? name(r.topTeam) : '—'}`); });
  L.push('');
  L.push('▸ MOMENTUM');
  if (hotStrk && hotVal > 0) L.push(`  Hottest streak: ${name(hotStrk)} on a ${hotVal}-win run`);
  if (perfect.length > 0) L.push(`  Unbeaten: ${perfect.map(r => name(r)).join(', ')}`); else L.push('  No unbeaten teams — everyone has taken a loss');
  if (bestWR) L.push(`  Best win rate (min 3 games): ${bestWR.name} at ${(bestWR.wr * 100).toFixed(1)}% (${bestWR.wins}W-${bestWR.losses}L)`);
  L.push('');
  L.push('▸ NOTABLE PATTERNS');
  if (closest) { const pA = byPts.indexOf(closest.a) + 1, pB = byPts.indexOf(closest.b) + 1; L.push(`  Tightest race: #${pA} ${name(closest.a)} vs #${pB} ${name(closest.b)} — ${closest.diff.toFixed(2)} pts apart`); }
  if (iTier >= 0 && Object.keys(tierCount).length) L.push(`  Tier distribution: ${Object.entries(tierCount).sort((a, b) => b[1] - a[1]).map(([k, v]) => k + ':' + v).join('  ')}`);
  L.push('');
  L.push('▸ TOP 5 STANDINGS');
  byPts.slice(0, 5).forEach((r, i) => { const m = ['◆', '◇', '▲', '△', '○'][i]; L.push(`  ${m} #${i + 1}  ${name(r).padEnd(22)} ${get(r, iPts).toFixed(2)} pts   ${iVict >= 0 ? get(r, iVict) : '—'}W-${iLose >= 0 ? get(r, iLose) : '—'}L`); });

  const full = L.join('\n'); let pos = 0;
  function tick() {
    if (pos >= full.length) return;
    for (let b = 0; b < 6 && pos < full.length; b++) {
      const ch = full[pos];
      if (ch === '▸') { const end = full.indexOf('\n', pos); const le = end === -1 ? full.length : end; const span = document.createElement('span'); span.className = 'ai-header-line'; span.textContent = full.slice(pos, le); output.appendChild(span); pos = le; break; }
      else { output.appendChild(document.createTextNode(ch)); pos++; }
    }
    requestAnimationFrame(tick);
  }
  tick();
}

/* ════════════════════════════════════════
   PRO ANALYSES — Match History & H2H
════════════════════════════════════════ */

const MONTH_NAMES = ['','January','February','March','April','May','June','July','August','September','October','November','December'];

function parseHistoryData(rows) {
  const data = [];
  for (const r of rows) {
    const team1 = String(r['Team 1'] ?? r['team 1'] ?? r['team1'] ?? '').trim();
    const team2 = String(r['Team 2'] ?? r['team 2'] ?? r['team2'] ?? '').trim();
    const map = String(r['Map'] ?? r['map'] ?? '').trim();
    const result = String(r['Result'] ?? r['result'] ?? '').trim();
    const event = String(r['Event'] ?? r['event'] ?? '').trim();
    const details = String(r['Details'] ?? r['details'] ?? '').trim();
    let dateRaw = r['Date'] ?? r['date'] ?? '';
    if (!team1 || !team2 || !result) continue;

    // Parse date
    let dateObj = null, dateStr = '';
    if (dateRaw) {
      if (typeof dateRaw === 'number') {
        // Excel serial date
        dateObj = new Date((dateRaw - 25569) * 86400000);
      } else {
        dateObj = new Date(dateRaw);
      }
      if (dateObj && !isNaN(dateObj.getTime())) {
        dateStr = dateObj.toISOString().split('T')[0];
      } else {
        dateStr = String(dateRaw).trim();
        dateObj = null;
      }
    }

    // Parse result (e.g., "13x4", "14x16", "16x9")
    const resParts = result.toLowerCase().split('x').map(s => parseInt(s.trim()));
    let score1 = resParts[0] || 0, score2 = resParts[1] || 0;
    let winner = score1 > score2 ? team1 : score2 > score1 ? team2 : 'Draw';

    data.push({
      team1, team2, map, result, event, details, dateStr, dateObj,
      score1, score2, winner,
      month: dateObj ? dateObj.getMonth() + 1 : 0,
      year: dateObj ? dateObj.getFullYear() : 0
    });
  }
  // Sort newest first
  data.sort((a, b) => {
    if (a.dateObj && b.dateObj) return b.dateObj - a.dateObj;
    return b.dateStr.localeCompare(a.dateStr);
  });
  return data;
}

function getAllTeamsFromHistory() {
  const set = new Set();
  historyData.forEach(m => { set.add(m.team1); set.add(m.team2); });
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' }));
}

function getAllMapsFromHistory() {
  const set = new Set();
  historyData.forEach(m => { if (m.map) set.add(m.map); });
  return Array.from(set).sort();
}

function getAllEventsFromHistory() {
  const set = new Set();
  historyData.forEach(m => { if (m.event) set.add(m.event); });
  return Array.from(set).sort();
}

function getAllDetailsFromHistory() {
  const set = new Set();
  historyData.forEach(m => { if (m.details) set.add(m.details); });
  return Array.from(set).sort();
}

function getAllMonthsFromHistory() {
  const set = new Set();
  historyData.forEach(m => { if (m.month) set.add(m.month); });
  return Array.from(set).sort((a, b) => a - b);
}

function getAllYearsFromHistory() {
  const set = new Set();
  historyData.forEach(m => { if (m.year) set.add(m.year); });
  return Array.from(set).sort((a, b) => b - a);
}

/* ─── Tab switching ─── */
function switchPaTab(tab) {
  document.querySelectorAll('.pa-tab').forEach(t => t.classList.toggle('active', t.dataset.paTab === tab));
  document.querySelectorAll('.pa-tab-content').forEach(c => c.classList.toggle('active', c.id === 'pa-tab-' + tab));
}

/* ─── Autocomplete helper ─── */
function setupAutocomplete(inputId, listId, getItems, onSelect) {
  const input = document.getElementById(inputId);
  const list = document.getElementById(listId);
  if (!input || !list) return;

  input.addEventListener('input', () => {
    let val = input.value.trim().toLowerCase();
    // Strip trailing period for autocomplete suggestions (period is only for filtering)
    const cleanVal = val.endsWith('.') ? val.slice(0, -1) : val;
    if (!cleanVal) { list.innerHTML = ''; list.style.display = 'none'; return; }

    const items = getItems().filter(t => t.toLowerCase().includes(cleanVal));
    if (!items.length) { list.innerHTML = ''; list.style.display = 'none'; return; }

    // Sort: exact match first, then startsWith, then includes
    items.sort((a, b) => {
      const aL = a.toLowerCase(), bL = b.toLowerCase();
      const aExact = aL === cleanVal, bExact = bL === cleanVal;
      if (aExact && !bExact) return -1;
      if (!aExact && bExact) return 1;
      const aStarts = aL.startsWith(cleanVal), bStarts = bL.startsWith(cleanVal);
      if (aStarts && !bStarts) return -1;
      if (!aStarts && bStarts) return 1;
      return a.localeCompare(b);
    });

    list.innerHTML = items.slice(0, 15).map(t => {
      const idx = t.toLowerCase().indexOf(cleanVal);
      const hl = idx >= 0 ? escHtml(t.slice(0, idx)) + '<mark>' + escHtml(t.slice(idx, idx + cleanVal.length)) + '</mark>' + escHtml(t.slice(idx + cleanVal.length)) : escHtml(t);
      return `<div class="pa-autocomplete-item" data-value="${escHtml(t)}">${hl}</div>`;
    }).join('');
    list.style.display = 'block';
  });

  list.addEventListener('click', e => {
    const item = e.target.closest('.pa-autocomplete-item');
    if (!item) return;
    input.value = item.dataset.value;
    list.innerHTML = ''; list.style.display = 'none';
    if (onSelect) onSelect(item.dataset.value);
  });

  input.addEventListener('blur', () => { setTimeout(() => { list.style.display = 'none'; }, 200); });
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      const first = list.querySelector('.pa-autocomplete-item');
      if (first) { input.value = first.dataset.value; list.innerHTML = ''; list.style.display = 'none'; if (onSelect) onSelect(first.dataset.value); }
      e.preventDefault();
    }
  });
}

/* ─── Init Pro Analyses ─── */
let paInitialized = false;
function initProAnalyses() {
  if (!historyData.length) return;

  // Populate filter dropdowns
  const maps = getAllMapsFromHistory();
  const events = getAllEventsFromHistory();
  const details = getAllDetailsFromHistory();
  const months = getAllMonthsFromHistory();
  const years = getAllYearsFromHistory();

  const mapSel = document.getElementById('paFilterMap');
  mapSel.innerHTML = '<option value="">All Maps</option>' + maps.map(m => `<option value="${escHtml(m)}">${escHtml(m)}</option>`).join('');

  // Also populate H2H map filter
  const h2hMapSel = document.getElementById('h2hMapFilter');
  if (h2hMapSel) {
    h2hMapSel.innerHTML = '<option value="">All Maps</option>' + maps.map(m => `<option value="${escHtml(m)}">${escHtml(m)}</option>`).join('');
  }

  const eventSel = document.getElementById('paFilterEvent');
  eventSel.innerHTML = '<option value="">All Events</option>' + events.map(e => `<option value="${escHtml(e)}">${escHtml(e)}</option>`).join('');

  const detailsSel = document.getElementById('paFilterDetails');
  detailsSel.innerHTML = '<option value="">All Details</option>' + details.map(d => `<option value="${escHtml(d)}">${escHtml(d)}</option>`).join('');

  const monthSel = document.getElementById('paFilterMonth');
  monthSel.innerHTML = '<option value="">All Months</option>' + months.map(m => `<option value="${m}">${MONTH_NAMES[m]}</option>`).join('');

  const yearSel = document.getElementById('paFilterYear');
  yearSel.innerHTML = '<option value="">All Years</option>' + years.map(y => `<option value="${y}">${y}</option>`).join('');

  // Only bind event listeners once
  if (!paInitialized) {
    paInitialized = true;

    // Setup autocompletes
    const allTeams = getAllTeamsFromHistory;
    setupAutocomplete('paFilterTeam', 'paTeamSuggestions', allTeams, val => { paFilteredTeam = val; renderHistoryTable(); });
    setupAutocomplete('paFilterOpponent', 'paOpponentSuggestions', allTeams, val => { paFilteredOpponent = val; renderHistoryTable(); });
    setupAutocomplete('h2hTeamA', 'h2hTeamASugg', allTeams, null);
    setupAutocomplete('h2hTeamB', 'h2hTeamBSugg', allTeams, null);

    // Filter listeners
    ['paFilterMap', 'paFilterEvent', 'paFilterDetails', 'paFilterMonth', 'paFilterYear'].forEach(id => {
      document.getElementById(id).addEventListener('change', renderHistoryTable);
    });
    document.getElementById('paFilterDate').addEventListener('change', renderHistoryTable);
    document.getElementById('paFilterTeam').addEventListener('input', () => {
      paFilteredTeam = document.getElementById('paFilterTeam').value.trim();
      renderHistoryTable();
    });
    document.getElementById('paFilterOpponent').addEventListener('input', () => {
      paFilteredOpponent = document.getElementById('paFilterOpponent').value.trim();
      renderHistoryTable();
    });

    // Clear filters
    document.getElementById('btnClearHistoryFilters').addEventListener('click', () => {
      document.getElementById('paFilterTeam').value = '';
      document.getElementById('paFilterOpponent').value = '';
      document.getElementById('paFilterMap').value = '';
      document.getElementById('paFilterEvent').value = '';
      document.getElementById('paFilterDetails').value = '';
      document.getElementById('paFilterDate').value = '';
      document.getElementById('paFilterMonth').value = '';
      document.getElementById('paFilterYear').value = '';
      paFilteredTeam = '';
      paFilteredOpponent = '';
      renderHistoryTable();
    });

    // H2H button
    document.getElementById('btnH2hAnalyze').addEventListener('click', runH2hAnalysis);

    // H2H map filter: re-run analysis when map changes (if teams are already selected)
    const h2hMapFilterEl = document.getElementById('h2hMapFilter');
    if (h2hMapFilterEl) {
      h2hMapFilterEl.addEventListener('change', () => {
        const a = document.getElementById('h2hTeamA').value.trim();
        const b = document.getElementById('h2hTeamB').value.trim();
        if (a && b) runH2hAnalysis();
      });
    }
  }

  renderHistoryTable();
}

/* ─── Team name match helper — supports "." suffix for strict exact match ─── */
function teamNameMatch(teamFromData, filterValue) {
  if (!filterValue) return true;
  const dataName = teamFromData.toLowerCase().trim();
  // If filter ends with ".", force strict exact match (strip the period)
  if (filterValue.endsWith('.')) {
    const exact = filterValue.slice(0, -1).toLowerCase().trim();
    return dataName === exact;
  }
  // Default: exact match
  return dataName === filterValue.toLowerCase().trim();
}

/* ─── Filter & render history ─── */
function getFilteredHistory() {
  const team = paFilteredTeam.trim();
  const opp = paFilteredOpponent.trim();
  const map = document.getElementById('paFilterMap').value;
  const event = document.getElementById('paFilterEvent').value;
  const details = document.getElementById('paFilterDetails').value;
  const date = document.getElementById('paFilterDate').value;
  const month = document.getElementById('paFilterMonth').value;
  const year = document.getElementById('paFilterYear').value;

  return historyData.filter(m => {
    if (team && !(teamNameMatch(m.team1, team) || teamNameMatch(m.team2, team))) return false;
    if (opp) {
      if (!team) return false; // opponent filter only works with a team
      const teamMatch = teamNameMatch(m.team1, team) || teamNameMatch(m.team2, team);
      const oppMatch = teamNameMatch(m.team1, opp) || teamNameMatch(m.team2, opp);
      if (!teamMatch || !oppMatch) return false;
    }
    if (map && m.map !== map) return false;
    if (event && m.event !== event) return false;
    if (details && m.details !== details) return false;
    if (date && m.dateStr !== date) return false;
    if (month && m.month !== parseInt(month)) return false;
    if (year && m.year !== parseInt(year)) return false;
    return true;
  });
}

function renderHistoryTable() {
  const body = document.getElementById('paHistoryBody');
  const filtered = getFilteredHistory();
  document.getElementById('paMatchCount').textContent = filtered.length + ' match' + (filtered.length !== 1 ? 'es' : '');

  // Update tab badge
  const badge = document.getElementById('paTabHistoryBadge');
  if (badge) badge.textContent = historyData.length ? historyData.length : '';

  if (!filtered.length) {
    body.innerHTML = '<tr><td colspan="8" class="empty-state">No matches found for current filters.</td></tr>';
    return;
  }

  const teamFilter = paFilteredTeam.trim();

  body.innerHTML = filtered.map(m => {
    // Format date
    const dDisplay = m.dateObj ? m.dateObj.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : m.dateStr;

    // Determine outcome for the filtered team
    let outcomeHtml = '';
    if (teamFilter) {
      const isTeam1 = teamNameMatch(m.team1, teamFilter);
      const isTeam2 = teamNameMatch(m.team2, teamFilter);
      const selectedTeam = isTeam1 ? m.team1 : isTeam2 ? m.team2 : '';
      if (selectedTeam) {
        const won = m.winner === selectedTeam;
        const cls = won ? 'pa-outcome-win' : 'pa-outcome-loss';
        const label = won ? 'WIN' : 'LOSS';
        outcomeHtml = `<span class="${cls}">${label}</span>`;
      }
    }

    // Highlight team names
    let t1Html = escHtml(m.team1);
    let t2Html = escHtml(m.team2);
    if (teamFilter) {
      if (teamNameMatch(m.team1, teamFilter)) t1Html = `<strong class="pa-team-highlight">${escHtml(m.team1)}</strong>`;
      if (teamNameMatch(m.team2, teamFilter)) t2Html = `<strong class="pa-team-highlight">${escHtml(m.team2)}</strong>`;
    }

    // Color the result based on winner
    let resultCls = '';
    if (teamFilter) {
      const isTeam1 = teamNameMatch(m.team1, teamFilter);
      const isTeam2 = teamNameMatch(m.team2, teamFilter);
      if (isTeam1 && m.score1 > m.score2) resultCls = 'pa-result-win';
      else if (isTeam1 && m.score1 < m.score2) resultCls = 'pa-result-loss';
      else if (isTeam2 && m.score2 > m.score1) resultCls = 'pa-result-win';
      else if (isTeam2 && m.score2 < m.score1) resultCls = 'pa-result-loss';
    }

    return `<tr>
      <td class="mono text-xs">${dDisplay}</td>
      <td>${t1Html}</td>
      <td class="center mono ${resultCls}">${escHtml(m.result)}</td>
      <td>${t2Html}</td>
      <td><span class="pa-map-badge">${escHtml(m.map)}</span></td>
      <td class="text-xs">${escHtml(m.event)}</td>
      <td class="text-xs">${escHtml(m.details)}</td>
      <td class="center">${outcomeHtml}</td>
    </tr>`;
  }).join('');
}

/* ════════════════════════════════════════
   H2H ANALYSIS ENGINE
════════════════════════════════════════ */
function runH2hAnalysis() {
  const nameA = document.getElementById('h2hTeamA').value.trim();
  const nameB = document.getElementById('h2hTeamB').value.trim();
  const output = document.getElementById('h2hOutput');
  const selectedMap = (document.getElementById('h2hMapFilter')?.value || '').trim();

  if (!nameA || !nameB) { output.innerHTML = '<div class="empty-state">Please select both teams.</div>'; return; }
  if (nameA.toLowerCase() === nameB.toLowerCase()) { output.innerHTML = '<div class="empty-state">Please select two different teams.</div>'; return; }

  const aL = nameA.toLowerCase(), bL = nameB.toLowerCase();

  // Get all matches for each team (optionally filtered by map)
  let matchesA = historyData.filter(m => m.team1.toLowerCase() === aL || m.team2.toLowerCase() === aL);
  let matchesB = historyData.filter(m => m.team1.toLowerCase() === bL || m.team2.toLowerCase() === bL);

  // H2H matches (all maps)
  let h2hMatchesAll = historyData.filter(m =>
    (m.team1.toLowerCase() === aL && m.team2.toLowerCase() === bL) ||
    (m.team1.toLowerCase() === bL && m.team2.toLowerCase() === aL)
  );

  // Apply map filter if selected
  if (selectedMap) {
    matchesA = matchesA.filter(m => m.map === selectedMap);
    matchesB = matchesB.filter(m => m.map === selectedMap);
  }

  // H2H matches (filtered)
  let h2hMatches = selectedMap
    ? h2hMatchesAll.filter(m => m.map === selectedMap)
    : h2hMatchesAll;

  // Stats helper
  function teamStats(matches, teamNameLower) {
    let wins = 0, losses = 0;
    const mapStats = {};
    matches.forEach(m => {
      const won = m.winner.toLowerCase() === teamNameLower;
      if (won) wins++; else losses++;
      if (m.map) {
        if (!mapStats[m.map]) mapStats[m.map] = { wins: 0, losses: 0, total: 0 };
        mapStats[m.map].total++;
        if (won) mapStats[m.map].wins++; else mapStats[m.map].losses++;
      }
    });

    // Streak (most recent consecutive results)
    let streak = 0, streakType = '';
    for (const m of matches) {
      const won = m.winner.toLowerCase() === teamNameLower;
      const res = won ? 'W' : 'L';
      if (!streakType) { streakType = res; streak = 1; }
      else if (res === streakType) streak++;
      else break;
    }

    return { wins, losses, total: matches.length, wr: matches.length > 0 ? (wins / matches.length * 100).toFixed(1) : '0.0', mapStats, streak, streakType };
  }

  const statsA = teamStats(matchesA, aL);
  const statsB = teamStats(matchesB, bL);

  // H2H specific stats
  let h2hWinsA = 0, h2hWinsB = 0;
  const h2hMapStats = {};
  h2hMatches.forEach(m => {
    if (m.winner.toLowerCase() === aL) h2hWinsA++;
    else if (m.winner.toLowerCase() === bL) h2hWinsB++;
    if (m.map) {
      if (!h2hMapStats[m.map]) h2hMapStats[m.map] = { winsA: 0, winsB: 0, total: 0 };
      h2hMapStats[m.map].total++;
      if (m.winner.toLowerCase() === aL) h2hMapStats[m.map].winsA++;
      else h2hMapStats[m.map].winsB++;
    }
  });

  // Last 5 matches for each team
  const last5A = matchesA.slice(0, 5);
  const last5B = matchesB.slice(0, 5);

  // Map proficiency: sort by total played, highlight best(yellow) and worst(purple)
  function mapProficiency(mapStats) {
    const entries = Object.entries(mapStats).map(([map, s]) => ({
      map, wins: s.wins, losses: s.losses, total: s.total, wr: s.total > 0 ? (s.wins / s.total * 100) : 0
    })).sort((a, b) => b.total - a.total);

    if (entries.length === 0) return entries;

    // Best: highest WR with at least some games; Worst: most losses > wins
    let bestWr = -1, worstWr = 101;
    entries.forEach(e => {
      if (e.wr > bestWr) bestWr = e.wr;
      if (e.wr < worstWr) worstWr = e.wr;
    });
    entries.forEach(e => {
      e.isBest = e.wr === bestWr && e.total > 0;
      e.isWorst = e.wr === worstWr && e.total > 0 && entries.length > 1;
    });
    return entries;
  }

  const mapsA = mapProficiency(statsA.mapStats);
  const mapsB = mapProficiency(statsB.mapStats);

  // === Build Verdict ===
  let verdictText = '', verdictColor = '';
  const totalFactors = { a: 0, b: 0 };

  // Factor: H2H
  if (h2hMatches.length > 0) {
    if (h2hWinsA > h2hWinsB) totalFactors.a += 2; else if (h2hWinsB > h2hWinsA) totalFactors.b += 2;
  }
  // Factor: Win rate
  if (parseFloat(statsA.wr) > parseFloat(statsB.wr)) totalFactors.a += 1; else if (parseFloat(statsB.wr) > parseFloat(statsA.wr)) totalFactors.b += 1;
  // Factor: Form (streak)
  if (statsA.streakType === 'W' && statsA.streak >= 2) totalFactors.a += 1;
  if (statsB.streakType === 'W' && statsB.streak >= 2) totalFactors.b += 1;
  // Factor: Total games experience
  if (statsA.total > statsB.total * 1.2) totalFactors.a += 0.5;
  if (statsB.total > statsA.total * 1.2) totalFactors.b += 0.5;

  // === Win Probability ===
  // Base: use win rates normalized, then apply H2H bonus
  let baseA = parseFloat(statsA.wr) || 50;
  let baseB = parseFloat(statsB.wr) || 50;
  // Normalize to probabilities
  let probA = baseA / (baseA + baseB) * 100;
  let probB = 100 - probA;
  // H2H bonus: shift ~3% per H2H win difference (capped)
  if (h2hMatches.length > 0) {
    const h2hDiff = h2hWinsA - h2hWinsB;
    const h2hShift = Math.max(-12, Math.min(12, h2hDiff * 3));
    probA += h2hShift;
    probB -= h2hShift;
  }
  // Form bonus: +2% for active win streak ≥ 2
  if (statsA.streakType === 'W' && statsA.streak >= 2) { probA += 2; probB -= 2; }
  if (statsB.streakType === 'W' && statsB.streak >= 2) { probB += 2; probA -= 2; }
  // Clamp
  probA = Math.max(5, Math.min(95, probA));
  probB = Math.max(5, Math.min(95, 100 - probA));

  if (totalFactors.a > totalFactors.b) {
    verdictText = `${nameA} has the edge based on H2H record, form, and win rate analysis.`;
    verdictColor = 'var(--accent)';
  } else if (totalFactors.b > totalFactors.a) {
    verdictText = `${nameB} has the edge based on H2H record, form, and win rate analysis.`;
    verdictColor = 'var(--accent2)';
  } else {
    verdictText = 'This matchup is extremely balanced. Expect a closely contested series.';
    verdictColor = 'var(--warn)';
  }

  // === Render ===
  let html = `<div class="h2h-report">`;

  // Map filter indicator
  if (selectedMap) {
    // Count overall H2H wins for context
    let h2hWinsAAll = 0, h2hWinsBAll = 0;
    h2hMatchesAll.forEach(m => {
      if (m.winner.toLowerCase() === aL) h2hWinsAAll++;
      else if (m.winner.toLowerCase() === bL) h2hWinsBAll++;
    });
    html += `<div class="h2h-map-filter-active">🗺️ Filtered by map: <strong>${escHtml(selectedMap)}</strong> — Overall H2H: ${h2hWinsAAll}–${h2hWinsBAll} (${h2hMatchesAll.length} total matches)</div>`;
  }

  // Header
  html += `<div class="h2h-header">
    <div class="h2h-team-card h2h-team-a">
      <div class="h2h-team-name"><span class="team-name-with-logo">${teamLogoImgHtml(nameA, 'team-logo-lg')}${escHtml(nameA)}</span></div>
      <div class="h2h-team-record">${statsA.wins}W — ${statsA.losses}L${selectedMap ? ' on ' + escHtml(selectedMap) : ''}</div>
      <div class="h2h-team-wr">${statsA.wr}% WR</div>
    </div>
    <div class="h2h-vs">
      <div class="h2h-vs-label">VS</div>
      <div class="h2h-vs-score">${h2hWinsA} — ${h2hWinsB}</div>
      <div class="h2h-vs-sub">${h2hMatches.length} H2H match${h2hMatches.length !== 1 ? 'es' : ''}${selectedMap ? ' on ' + escHtml(selectedMap) : ''}</div>
    </div>
    <div class="h2h-team-card h2h-team-b">
      <div class="h2h-team-name"><span class="team-name-with-logo">${teamLogoImgHtml(nameB, 'team-logo-lg')}${escHtml(nameB)}</span></div>
      <div class="h2h-team-record">${statsB.wins}W — ${statsB.losses}L${selectedMap ? ' on ' + escHtml(selectedMap) : ''}</div>
      <div class="h2h-team-wr">${statsB.wr}% WR</div>
    </div>
  </div>`;

  // Comparison bars
  html += `<div class="h2h-section">
    <div class="h2h-section-title">Overall Comparison</div>
    <div class="h2h-bars">
      ${h2hCompBar('Win Rate', parseFloat(statsA.wr), parseFloat(statsB.wr), '%')}
      ${h2hCompBar('Total Wins', statsA.wins, statsB.wins, '')}
      ${h2hCompBar('Total Losses', statsA.losses, statsB.losses, '', true)}
      ${h2hCompBar('Current Streak', statsA.streakType === 'W' ? statsA.streak : 0, statsB.streakType === 'W' ? statsB.streak : 0, 'W')}
    </div>
  </div>`;

  // Streaks section
  html += `<div class="h2h-section">
    <div class="h2h-section-title">Current Form</div>
    <div class="h2h-form-grid">
      <div class="h2h-form-card">
        <div class="h2h-form-team">${escHtml(nameA)}</div>
        <div class="h2h-form-streak ${statsA.streakType === 'W' ? 'streak-win' : 'streak-loss'}">${statsA.streak}${statsA.streakType} Streak</div>
        <div class="h2h-form-dots">${last5A.map(m => {
          const won = m.winner.toLowerCase() === aL;
          return `<span class="h2h-form-dot ${won ? 'dot-win' : 'dot-loss'}" title="${escHtml(m.team1)} ${m.result} ${escHtml(m.team2)}">${won ? 'W' : 'L'}</span>`;
        }).join('')}</div>
      </div>
      <div class="h2h-form-card">
        <div class="h2h-form-team">${escHtml(nameB)}</div>
        <div class="h2h-form-streak ${statsB.streakType === 'W' ? 'streak-win' : 'streak-loss'}">${statsB.streak}${statsB.streakType} Streak</div>
        <div class="h2h-form-dots">${last5B.map(m => {
          const won = m.winner.toLowerCase() === bL;
          return `<span class="h2h-form-dot ${won ? 'dot-win' : 'dot-loss'}" title="${escHtml(m.team1)} ${m.result} ${escHtml(m.team2)}">${won ? 'W' : 'L'}</span>`;
        }).join('')}</div>
      </div>
    </div>
  </div>`;

  // H2H Map breakdown
  if (Object.keys(h2hMapStats).length > 0) {
    html += `<div class="h2h-section">
      <div class="h2h-section-title">H2H Map Breakdown</div>
      <div class="table-wrapper" style="max-height:240px">
        <table class="h2h-map-table">
          <thead><tr><th>Map</th><th class="center">${escHtml(nameA)} Wins</th><th class="center">${escHtml(nameB)} Wins</th><th class="center">Total</th></tr></thead>
          <tbody>${Object.entries(h2hMapStats).sort((a, b) => b[1].total - a[1].total).map(([map, s]) => {
            const aWinCls = s.winsA > s.winsB ? 'pa-result-win' : s.winsA < s.winsB ? 'pa-result-loss' : '';
            const bWinCls = s.winsB > s.winsA ? 'pa-result-win' : s.winsB < s.winsA ? 'pa-result-loss' : '';
            return `<tr><td><span class="pa-map-badge">${escHtml(map)}</span></td><td class="center mono ${aWinCls}">${s.winsA}</td><td class="center mono ${bWinCls}">${s.winsB}</td><td class="center mono">${s.total}</td></tr>`;
          }).join('')}</tbody>
        </table>
      </div>
    </div>`;
  }

  // Map proficiency per team
  function renderMapProf(teamName, maps) {
    if (!maps.length) return `<div class="h2h-no-data">No map data for ${escHtml(teamName)}</div>`;
    return `<div class="h2h-map-prof">
      ${maps.map(m => {
        let cls = '';
        if (m.isBest) cls = 'pa-map-best';
        else if (m.isWorst) cls = 'pa-map-worst';
        return `<div class="h2h-map-chip ${cls}">
          <span class="h2h-map-chip-name">${escHtml(m.map)}</span>
          <span class="h2h-map-chip-stats">${m.wins}W-${m.losses}L</span>
          <span class="h2h-map-chip-wr">${m.wr.toFixed(0)}%</span>
        </div>`;
      }).join('')}
      <div class="h2h-map-legend"><span class="h2h-legend-best">● Best map</span> <span class="h2h-legend-worst">● Worst map</span></div>
    </div>`;
  }

  html += `<div class="h2h-section">
    <div class="h2h-section-title">Map Proficiency</div>
    <div class="h2h-map-prof-grid">
      <div>
        <div class="h2h-map-prof-team">${escHtml(nameA)}</div>
        ${renderMapProf(nameA, mapsA)}
      </div>
      <div>
        <div class="h2h-map-prof-team">${escHtml(nameB)}</div>
        ${renderMapProf(nameB, mapsB)}
      </div>
    </div>
  </div>`;

  // Last 5 matches per team
  function renderLast5(teamName, last5, teamLower) {
    if (!last5.length) return '<div class="h2h-no-data">No recent matches.</div>';
    return `<div class="table-wrapper" style="max-height:220px">
      <table class="h2h-recent-table">
        <thead><tr><th>Date</th><th>Opponent</th><th class="center">Result</th><th>Map</th><th class="center">W/L</th></tr></thead>
        <tbody>${last5.map(m => {
          const opp = m.team1.toLowerCase() === teamLower ? m.team2 : m.team1;
          const won = m.winner.toLowerCase() === teamLower;
          const dDisplay = m.dateObj ? m.dateObj.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' }) : m.dateStr;
          return `<tr>
            <td class="mono text-xs">${dDisplay}</td>
            <td>${escHtml(opp)}</td>
            <td class="center mono ${won ? 'pa-result-win' : 'pa-result-loss'}">${escHtml(m.result)}</td>
            <td><span class="pa-map-badge">${escHtml(m.map)}</span></td>
            <td class="center"><span class="${won ? 'pa-outcome-win' : 'pa-outcome-loss'}">${won ? 'W' : 'L'}</span></td>
          </tr>`;
        }).join('')}</tbody>
      </table>
    </div>`;
  }

  html += `<div class="h2h-section">
    <div class="h2h-section-title">Recent Matches (Last 5)</div>
    <div class="h2h-recent-grid">
      <div>
        <div class="h2h-recent-team">${escHtml(nameA)}</div>
        ${renderLast5(nameA, last5A, aL)}
      </div>
      <div>
        <div class="h2h-recent-team">${escHtml(nameB)}</div>
        ${renderLast5(nameB, last5B, bL)}
      </div>
    </div>
  </div>`;

  // H2H full history
  if (h2hMatches.length > 0) {
    html += `<div class="h2h-section">
      <div class="h2h-section-title">H2H Match History</div>
      <div class="table-wrapper" style="max-height:280px">
        <table class="h2h-history-table">
          <thead><tr><th>Date</th><th>Team 1</th><th class="center">Result</th><th>Team 2</th><th>Map</th><th>Event</th></tr></thead>
          <tbody>${h2hMatches.map(m => {
            const dDisplay = m.dateObj ? m.dateObj.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : m.dateStr;
            const t1Win = m.winner.toLowerCase() === m.team1.toLowerCase();
            return `<tr>
              <td class="mono text-xs">${dDisplay}</td>
              <td class="${t1Win ? 'pa-winner-cell' : ''}">${escHtml(m.team1)}</td>
              <td class="center mono">${escHtml(m.result)}</td>
              <td class="${!t1Win ? 'pa-winner-cell' : ''}">${escHtml(m.team2)}</td>
              <td><span class="pa-map-badge">${escHtml(m.map)}</span></td>
              <td class="text-xs">${escHtml(m.event)}</td>
            </tr>`;
          }).join('')}</tbody>
        </table>
      </div>
    </div>`;
  }

  // Verdict
  const probAColor = probA >= probB ? 'var(--accent)' : 'var(--muted2)';
  const probBColor = probB > probA ? 'var(--accent2, #a85d3a)' : 'var(--muted2)';
  html += `<div class="h2h-verdict">
    <div class="h2h-verdict-icon">🎯</div>
    <div class="h2h-verdict-title" style="color:${verdictColor}">Verdict</div>
    <div class="h2h-verdict-text">${escHtml(verdictText)}</div>
    <div class="h2h-win-prob" style="margin:12px 0 8px">
      <div style="display:flex;justify-content:space-between;font-family:var(--font-mono);font-size:12px;margin-bottom:4px">
        <span style="color:${probAColor};font-weight:700">${escHtml(nameA)} ${probA.toFixed(1)}%</span>
        <span style="color:var(--muted2);font-size:10px">WIN PROBABILITY</span>
        <span style="color:${probBColor};font-weight:700">${probB.toFixed(1)}% ${escHtml(nameB)}</span>
      </div>
      <div style="display:flex;height:8px;border-radius:4px;overflow:hidden;background:rgba(28,26,23,0.07)">
        <div style="width:${probA.toFixed(1)}%;background:var(--accent);transition:width 0.5s"></div>
        <div style="width:${probB.toFixed(1)}%;background:var(--accent2, #a85d3a);transition:width 0.5s"></div>
      </div>
    </div>
    <div class="h2h-verdict-factors">
      <span>H2H: ${h2hWinsA}-${h2hWinsB}</span>
      <span>WR: ${statsA.wr}% vs ${statsB.wr}%</span>
      <span>Form: ${statsA.streak}${statsA.streakType} vs ${statsB.streak}${statsB.streakType}</span>
    </div>
  </div>`;

  html += `</div>`;
  output.innerHTML = html;
}

/* ─── Comparison bar helper ─── */
function h2hCompBar(label, valA, valB, suffix, invertColors = false) {
  const total = valA + valB || 1;
  const pctA = (valA / total * 100).toFixed(1);
  const pctB = (valB / total * 100).toFixed(1);
  let clsA = 'h2h-bar-a', clsB = 'h2h-bar-b';
  if (invertColors) { clsA = 'h2h-bar-b'; clsB = 'h2h-bar-a'; } // losses: lower is better
  const winnerCls = valA > valB ? clsA + ' h2h-bar-lead' : valB > valA ? clsB + ' h2h-bar-lead' : '';
  return `<div class="h2h-bar-row">
    <div class="h2h-bar-label">${escHtml(label)}</div>
    <div class="h2h-bar-values">
      <span class="h2h-bar-val-a ${valA > valB && !invertColors ? 'h2h-val-win' : ''}">${valA}${suffix}</span>
      <div class="h2h-bar-track">
        <div class="${clsA}" style="width:${pctA}%"></div>
        <div class="${clsB}" style="width:${pctB}%"></div>
      </div>
      <span class="h2h-bar-val-b ${valB > valA && !invertColors ? 'h2h-val-win' : ''}">${valB}${suffix}</span>
    </div>
  </div>`;
}

/* ════════════════════════════════════════
   CSV EXPORT
════════════════════════════════════════ */
function toCsv(c) {
  const h = ['Position', 'Team', 'Points', 'Points Diff', 'Region', 'Tier', 'Position Delta', 'OldPos', 'NewPos'];
  const lines = [h.join(',')];
  for (const r of c) lines.push([r.newPos, `"${String(r.team).replaceAll('"', '""')}"`, Math.trunc(r.points), r.pointsDiffText, `"${String(r.region || '').replaceAll('"', '""')}"`, `"${String(r.tier || '').replaceAll('"', '""')}"`, r.deltaText, r.oldPos ?? '', r.newPos].join(','));
  return lines.join('\n');
}

function downloadText(fn, content, mime = 'text/plain') {
  const b = new Blob([content], { type: mime });
  const u = URL.createObjectURL(b);
  const a = document.createElement('a');
  a.href = u; a.download = fn;
  document.body.appendChild(a); a.click();
  a.remove(); URL.revokeObjectURL(u);
}

/* ════════════════════════════════════════
   CHANGE HISTORY
════════════════════════════════════════ */
function loadChangeHistory() {
  try {
    const saved = localStorage.getItem(CHANGE_HISTORY_KEY);
    changeHistory = saved ? JSON.parse(saved) : [];
  } catch (e) { changeHistory = []; }
}

function saveChangeHistory() {
  try { localStorage.setItem(CHANGE_HISTORY_KEY, JSON.stringify(changeHistory)); } catch (e) {}
}

function recordChange(fileName, lastModified, action = 'load') {
  const entry = {
    file: fileName,
    loadedAt: new Date().toISOString(),
    fileModified: lastModified ? new Date(lastModified).toISOString() : null,
    action
  };
  changeHistory.unshift(entry);
  // Keep max 50 entries
  if (changeHistory.length > 50) changeHistory.length = 50;
  saveChangeHistory();
}

function renderChangeHistory() {
  const card = document.getElementById('changeHistoryCard');
  const list = document.getElementById('changeHistoryList');
  if (!card || !list) return;

  if (changeHistory.length === 0) {
    card.style.display = 'none';
    return;
  }

  card.style.display = '';
  const fileIcons = {
    'old.xlsx': '📄',
    'ranking.xlsx': '📊',
    'history.xlsx': '📜'
  };
  function iconFor(file) {
    if (fileIcons[file]) return fileIcons[file];
    if (/^r\d+\.xlsx$/i.test(file)) return '🗄️';
    return '📁';
  }

  const now = new Date();
  function timeAgo(dateStr) {
    const d = new Date(dateStr);
    const diff = Math.floor((now - d) / 1000);
    if (diff < 60) return 'just now';
    if (diff < 3600) return Math.floor(diff / 60) + 'm ago';
    if (diff < 86400) return Math.floor(diff / 3600) + 'h ago';
    return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) + ' ' + d.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
  }

  function formatDate(dateStr) {
    if (!dateStr) return '—';
    const d = new Date(dateStr);
    return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) + ' ' + d.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
  }

  list.innerHTML = changeHistory.slice(0, 20).map(entry => {
    const icon = iconFor(entry.file);
    const actionLabel = entry.action === 'reload' ? 'reload' : entry.action === 'nav' ? 'time-travel' : 'load';
    const labelCls = entry.action === 'reload' ? 'reload' : entry.action === 'nav' ? 'nav' : '';
    const modifiedInfo = entry.fileModified ? `File modified: ${formatDate(entry.fileModified)}` : '';
    return `<div class="change-history-item" title="${modifiedInfo}">
      <span class="change-history-icon">${icon}</span>
      <span class="change-history-file">${entry.file}</span>
      <span class="change-history-label ${labelCls}">${actionLabel}</span>
      ${entry.fileModified ? `<span style="font-size:10px;color:var(--muted2)">mod: ${formatDate(entry.fileModified)}</span>` : ''}
      <span class="change-history-date">${timeAgo(entry.loadedAt)}</span>
    </div>`;
  }).join('');
}

// Load history from localStorage on init
loadChangeHistory();

// Clear history button
document.getElementById('btnClearHistory')?.addEventListener('click', () => {
  changeHistory = [];
  saveChangeHistory();
  renderChangeHistory();
});

// Render on load (will show or hide based on data)
renderChangeHistory();

/* ════════════════════════════════════════
   EVENT LISTENERS
════════════════════════════════════════ */
function snapNavToLive() {
  if (!navChain.length) return;
  navToIndex = navChain.length - 1;
  navFromIndex = navToIndex > 0 ? navToIndex - 1 : 0;
  syncNavSelectsToIndices(navFromIndex, navToIndex);
  updateNavUI();
}

$('#btnCompare').addEventListener('click', async () => {
  await loadAll(true);
  snapNavToLive();
});

$('#btnReloadOld').addEventListener('click', async () => {
  try {
    setStatus('Reloading old.xlsx…', 'loading');
    const d = await readExcel(OLD_URL + '?v=' + Date.now());
    oldRowsCache = d.json;
    recordChange('old.xlsx', d.lastModified, 'reload');
    computeVRS(); renderDashboard(); renderChangeHistory();
    snapNavToLive();
    setStatus('Old reloaded', 'ok');
  } catch (e) { setStatus('Error reloading old.xlsx', 'err'); }
});

$('#btnReloadNew').addEventListener('click', async () => {
  try {
    setStatus('Reloading ranking.xlsx…', 'loading');
    const d = await readExcel(NEW_URL + '?v=' + Date.now());
    newRowsCache = d.json;
    recordChange('ranking.xlsx', d.lastModified, 'reload');
    hltvHeaders = (d.raw[0] || []).map(h => (h ?? '').toString().trim());
    hltvRows = d.raw.slice(1).filter(r => r.some(c => c !== undefined && c !== ''));
    computeVRS(); renderDashboard(); renderHltvCharts(); renderChangeHistory();
    snapNavToLive();
    setStatus('New reloaded', 'ok');
  } catch (e) { setStatus('Error reloading ranking.xlsx', 'err'); }
});

$('#navPrevBtn')?.addEventListener('click', () => { if (navToIndex > 0) loadPairAt(navToIndex - 1); });
$('#navNextBtn')?.addEventListener('click', () => { if (navToIndex < navChain.length - 1) loadPairAt(navToIndex + 1); });
$('#navResetBtn')?.addEventListener('click', () => { if (navChain.length) loadPairAt(navChain.length - 1); });

$('#navCustomCompareBtn')?.addEventListener('click', () => {
  const fromSel = $('#navFromSelect'), toSel = $('#navToSelect');
  if (!fromSel || !toSel) return;
  loadCustomPair(parseInt(fromSel.value, 10), parseInt(toSel.value, 10));
});


['filterText', 'filterRegion', 'filterTier'].forEach(id => {
  const el = document.getElementById(id);
  if (el) {
    el.addEventListener('input', () => { if (lastCombined) renderGlobalTable(lastCombined); });
    el.addEventListener('change', () => { if (lastCombined) renderGlobalTable(lastCombined); });
  }
});

$('#btnClearFilters').addEventListener('click', () => {
  $('#filterText').value = ''; $('#filterRegion').value = ''; $('#filterTier').value = '';
  if (lastCombined) renderGlobalTable(lastCombined);
});

$('#btnDownloadCsv').addEventListener('click', () => { if (lastCombined) downloadText('nexus_ranking.csv', toCsv(lastCombined), 'text/csv;charset=utf-8'); });

$('#regionRankingSelect').addEventListener('change', recomputeRegionRanking);

['matchTeamA', 'matchTeamB', 'matchResult'].forEach(id => {
  document.getElementById(id).addEventListener('change', renderMatchPredictor);
});

$('#dashMetric').addEventListener('change', renderDashboard);
$('#btnDashSortOrder').addEventListener('click', () => {
  dashSortDesc = !dashSortDesc;
  renderDashboard();
});
$('#chartRegionFilter').addEventListener('change', renderHltvCharts);
$('#teamPieSelect').addEventListener('change', renderTeamPie);
$('#legendsRegionFilter').addEventListener('change', renderLegendsTable);

// Keyboard shortcut: 0 → menu
document.addEventListener('keydown', e => {
  if (e.key !== '0' && e.code !== 'Numpad0') return;
  const el = document.activeElement, tag = (el?.tagName || '').toLowerCase();
  if (['input', 'textarea', 'select'].includes(tag) || el?.isContentEditable) return;
  window.location.href = 'index.html';
});

/* ════════════════════════════════════════
   SEEDING GENERATOR
════════════════════════════════════════ */
let seedTeamData = []; // loaded from ranking.xlsx via currentNewTeams
let seedSelectedSet = new Set(); // lowercase keys of selected teams

function populateSeedTeamList(filter = '') {
  const list = $('#seedTeamList');
  if (!currentNewTeams || !currentNewTeams.length) {
    list.innerHTML = '<div class="empty-state" style="font-size:12px; padding:16px">No ranking data loaded yet.</div>';
    updateSeedCount();
    return;
  }

  // Cache sorted teams on first call
  if (!seedTeamData.length) {
    seedTeamData = [...currentNewTeams].sort((a, b) => b.points - a.points);
  }

  const q = filter.trim().toLowerCase();
  const filtered = q
    ? seedTeamData.filter(t => t.team.toLowerCase().includes(q))
    : seedTeamData;

  if (!filtered.length) {
    list.innerHTML = '<div class="empty-state" style="font-size:12px; padding:16px">No teams match the filter.</div>';
    return;
  }

  list.innerHTML = filtered.map((t, i) => {
    const key = t.team.toLowerCase();
    const checked = seedSelectedSet.has(key) ? 'checked' : '';
    const selCls = seedSelectedSet.has(key) ? ' selected' : '';
    const rank = seedTeamData.indexOf(t) + 1;
    return `<label class="seed-team-item${selCls}" data-key="${escHtml(key)}">
      <input type="checkbox" ${checked} data-team-key="${escHtml(key)}">
      <span class="seed-team-rank-label">#${rank}</span>
      <span class="seed-team-name-label">${escHtml(t.team)}</span>
      <span class="seed-team-pts-label">${Math.trunc(t.points).toLocaleString()} pts</span>
    </label>`;
  }).join('');

  // Attach change listeners
  list.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    cb.addEventListener('change', () => {
      const key = cb.dataset.teamKey;
      if (cb.checked) seedSelectedSet.add(key);
      else seedSelectedSet.delete(key);
      cb.closest('.seed-team-item').classList.toggle('selected', cb.checked);
      updateSeedCount();
    });
  });
}

function updateSeedCount() {
  const el = $('#seedSelectedCount');
  if (el) el.textContent = `(${seedSelectedSet.size} selected)`;
}

function seedQuickSelect() {
  const textarea = $('#seedQuickInput');
  const lines = textarea.value.split('\n').map(s => s.trim()).filter(Boolean);
  if (!lines.length) return;
  if (!seedTeamData.length) return;

  const teamMap = new Map();
  seedTeamData.forEach(t => teamMap.set(t.team.toLowerCase(), t));

  let matched = 0, notFound = [];
  for (const line of lines) {
    const key = line.toLowerCase();
    if (teamMap.has(key)) {
      seedSelectedSet.add(key);
      matched++;
    } else {
      // Try partial match
      const partial = seedTeamData.find(t => t.team.toLowerCase().includes(key));
      if (partial) {
        seedSelectedSet.add(partial.team.toLowerCase());
        matched++;
      } else {
        notFound.push(line);
      }
    }
  }

  // Re-render the list to update checkboxes
  populateSeedTeamList($('#seedTeamSearch').value);
  updateSeedCount();

  const info = $('#seedInfo');
  let msg = `<span style="color:var(--good)">✓ ${matched} team${matched !== 1 ? 's' : ''} matched</span>`;
  if (notFound.length) {
    msg += ` · <span style="color:var(--warn)">⚠ Not found: ${notFound.map(n => escHtml(n)).join(', ')}</span>`;
  }
  info.innerHTML = msg;
}

function generateSeeding() {
  const output = $('#seedOutput');
  const info = $('#seedInfo');

  if (seedSelectedSet.size < 2) {
    info.innerHTML = '<span style="color:var(--bad)">⚠ Select at least 2 teams.</span>';
    return;
  }

  // Build selected teams array from ranking data
  const teamMap = new Map();
  if (currentNewTeams) currentNewTeams.forEach(t => teamMap.set(t.team.toLowerCase(), t));

  const matched = [];
  for (const key of seedSelectedSet) {
    const data = teamMap.get(key);
    if (data) {
      matched.push({ name: data.team, points: data.points, pos: data.pos, region: data.region, tier: data.tier });
    }
  }

  if (matched.length < 2) {
    info.innerHTML = '<span style="color:var(--bad)">⚠ At least 2 valid teams required.</span>';
    return;
  }

  // Sort by points descending (strongest first)
  matched.sort((a, b) => b.points - a.points);

  // Info summary
  const total = matched.length;
  const hasOdd = total % 2 !== 0;
  let infoHtml = `<span style="color:var(--accent)">${total} team${total > 1 ? 's' : ''} selected</span>`;
  if (hasOdd) {
    infoHtml += ` · <span style="color:var(--warn)">Odd count → 1 BYE</span>`;
  }
  info.innerHTML = infoHtml;

  // Generate matchups: pair extremes (1st vs last, 2nd vs 2nd-last, etc.)
  const matches = [];
  const half = Math.floor(total / 2);
  for (let i = 0; i < half; i++) {
    const strong = matched[i];
    const weak = matched[total - 1 - i];
    matches.push({ teamA: strong, teamB: weak });
  }

  // Handle BYE for odd team count — the middle team gets a bye
  let byeTeam = null;
  if (hasOdd) {
    byeTeam = matched[half];
  }

  // Shuffle match order
  for (let i = matches.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [matches[i], matches[j]] = [matches[j], matches[i]];
  }

  // Randomize sides within each match & pick one to highlight (wine color)
  matches.forEach(m => {
    if (Math.random() < 0.5) {
      const tmp = m.teamA;
      m.teamA = m.teamB;
      m.teamB = tmp;
    }
    m.wineHighlight = Math.random() < 0.5 ? 'A' : 'B';
  });

  // Render output
  let html = '<div class="seed-bracket">';

  // Seeded order header
  html += '<div class="seed-order-header">';
  html += '<span class="seed-order-title">🏆 Seed Order (by Points)</span>';
  html += '<div class="seed-order-list">';
  matched.forEach((t, i) => {
    html += `<span class="seed-order-chip">${i + 1}. ${escHtml(t.name)} <span class="seed-order-pts">${Math.trunc(t.points).toLocaleString()} pts</span></span>`;
  });
  html += '</div></div>';

  // Matches
  html += '<div class="seed-matches-title">⚔️ Matchups <span style="color:var(--muted2); font-weight:400">(randomized order & sides)</span></div>';
  html += '<div class="seed-matches-grid">';
  matches.forEach((m, idx) => {
    const aWine = m.wineHighlight === 'A' ? ' seed-wine' : '';
    const bWine = m.wineHighlight === 'B' ? ' seed-wine' : '';
    html += `<div class="seed-match-card">`;
    html += `<div class="seed-match-num">Match ${idx + 1}</div>`;
    html += `<div class="seed-match-body">`;
    html += `<div class="seed-match-team${aWine}">`;
    html += `<span class="seed-team-name">${escHtml(m.teamA.name)}</span>`;
    html += `<span class="seed-team-pts">${Math.trunc(m.teamA.points).toLocaleString()} pts</span>`;
    html += `</div>`;
    html += `<div class="seed-vs">VS</div>`;
    html += `<div class="seed-match-team${bWine}">`;
    html += `<span class="seed-team-name">${escHtml(m.teamB.name)}</span>`;
    html += `<span class="seed-team-pts">${Math.trunc(m.teamB.points).toLocaleString()} pts</span>`;
    html += `</div>`;
    html += `</div></div>`;
  });

  // BYE card
  if (byeTeam) {
    html += `<div class="seed-match-card seed-bye-card">`;
    html += `<div class="seed-match-num">BYE</div>`;
    html += `<div class="seed-match-body seed-bye-body">`;
    html += `<div class="seed-match-team seed-bye-team">`;
    html += `<span class="seed-team-name">${escHtml(byeTeam.name)}</span>`;
    html += `<span class="seed-team-pts">${Math.trunc(byeTeam.points).toLocaleString()} pts</span>`;
    html += `</div>`;
    html += `<div class="seed-bye-label">⟶ Advances directly</div>`;
    html += `</div></div>`;
  }

  html += '</div></div>';
  output.innerHTML = html;
}

// Event listeners for seeding
$('#btnSeedGenerate').addEventListener('click', generateSeeding);

$('#btnSeedQuickSelect').addEventListener('click', seedQuickSelect);

$('#btnSeedSelectAll').addEventListener('click', () => {
  if (!seedTeamData.length) return;
  const q = $('#seedTeamSearch').value.trim().toLowerCase();
  const visible = q ? seedTeamData.filter(t => t.team.toLowerCase().includes(q)) : seedTeamData;
  visible.forEach(t => seedSelectedSet.add(t.team.toLowerCase()));
  populateSeedTeamList($('#seedTeamSearch').value);
  updateSeedCount();
});

$('#btnSeedClearSelection').addEventListener('click', () => {
  seedSelectedSet.clear();
  $('#seedQuickInput').value = '';
  $('#seedInfo').innerHTML = '';
  $('#seedOutput').innerHTML = '<div class="empty-state">Select teams and click <strong>Generate Seeding</strong> to create matchups.</div>';
  populateSeedTeamList($('#seedTeamSearch').value);
  updateSeedCount();
});

$('#seedTeamSearch').addEventListener('input', () => {
  populateSeedTeamList($('#seedTeamSearch').value);
});

/* ════════════════════════════════════════
   EVENTS TOOL — kz-events-backup.json
════════════════════════════════════════ */
let eventsData = [];
let evEditIndex = -1; // -1 = add mode, >=0 = edit mode

const EVENTS_JSON_URL = 'file/kz-events-backup.json';

async function loadEventsJson() {
  try {
    const res = await fetch(EVENTS_JSON_URL, { cache: 'no-store' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    eventsData = await res.json();
    if (!Array.isArray(eventsData)) eventsData = [];
  } catch (e) {
    console.warn('Events JSON not found or invalid, starting empty.', e);
    eventsData = [];
  }
  refreshEventsUI();
}

/* ─── Summary Stats ─── */
function updateEventsStats() {
  const total = eventsData.length;
  const live = eventsData.filter(e => (e.status || '').toLowerCase() === 'live').length;
  const majors = eventsData.filter(e => e.major === 'yes' || e.major === true).length;
  let totalPrize = 0;
  eventsData.forEach(e => {
    if (e.prizePool) {
      const n = parseFloat(String(e.prizePool).replace(/[^0-9.]/g, ''));
      if (!isNaN(n)) totalPrize += n;
    }
  });
  const el = id => document.getElementById(id);
  el('evTotalEvents').textContent = total;
  el('evLiveEvents').textContent = live;
  el('evMajors').textContent = majors;
  el('evTotalPrize').textContent = '$' + totalPrize.toLocaleString('en-US', { maximumFractionDigits: 0 });
}

/* ─── Medal awardee splitter ───
   VP (and occasionally other medals) can be shared by several players and is
   registered as a single string: "player1 - player2 - player3".
   Splits ONLY on a hyphen that has whitespace on both sides, so nicknames that
   contain a hyphen themselves (e.g. "t-3xy") are never broken apart.
   Spacing around the separator is free-form: "a - b" and "a   -   b" both work. */
function splitAwardees(field) {
  if (!field) return [];
  return String(field)
    .split(/\s+-\s+/)
    .map(n => n.trim())
    .filter(Boolean);
}

/* ─── Populate filter dropdowns ─── */
function populateEventsFilters() {
  // Player medal filter
  const playerSet = new Set();
  eventsData.forEach(e => {
    if ((e.status || '').toLowerCase() === 'finished') {
      ['mvp', 'evp', 'vp'].forEach(k => {
        splitAwardees(e[k]).forEach(n => playerSet.add(n));
      });
    }
  });
  const playerSel = document.getElementById('evFilterPlayer');
  const prevPlayer = playerSel.value;
  playerSel.innerHTML = '<option value="">All Players</option>';
  [...playerSet].sort((a, b) => a.localeCompare(b)).forEach(p => {
    playerSel.innerHTML += `<option value="${escHtml(p)}">${escHtml(p)}</option>`;
  });
  playerSel.value = prevPlayer;

  // Team wins filter (only winners / top1)
  const teamSet = new Set();
  eventsData.forEach(e => {
    if ((e.status || '').toLowerCase() === 'finished') {
      const w = (e.winner || '').trim();
      if (w) teamSet.add(w);
    }
  });
  const teamSel = document.getElementById('evFilterTeamWin');
  const prevTeam = teamSel.value;
  teamSel.innerHTML = '<option value="">All Teams</option>';
  [...teamSet].sort((a, b) => a.localeCompare(b)).forEach(t => {
    teamSel.innerHTML += `<option value="${escHtml(t)}">${escHtml(t)}</option>`;
  });
  teamSel.value = prevTeam;
}

/* ─── Filter events ─── */
function getFilteredEvents() {
  const nameQ = (document.getElementById('evFilterName').value || '').trim().toLowerCase();
  const playerQ = document.getElementById('evFilterPlayer').value;
  const teamQ = document.getElementById('evFilterTeamWin').value;
  const starsQ = document.getElementById('evFilterStars').value;
  const regionQ = document.getElementById('evFilterRegion').value;

  return eventsData.filter((e, idx) => {
    // Name filter
    if (nameQ && !(e.name || '').toLowerCase().includes(nameQ)) return false;
    // Stars filter
    if (starsQ && String(e.stars) !== starsQ) return false;
    // Region filter
    if (regionQ && (e.region || '') !== regionQ) return false;
    // Player medal filter
    if (playerQ) {
      const medalists = [e.mvp, e.evp, e.vp].reduce((acc, v) => acc.concat(splitAwardees(v)), []);
      if (!medalists.includes(playerQ)) return false;
    }
    // Team win filter
    if (teamQ && (e.winner || '').trim() !== teamQ) return false;
    return true;
  });
}

/* ─── Render events grid ─── */
function renderEventsGrid() {
  const grid = document.getElementById('eventsGrid');
  const filtered = getFilteredEvents();

  if (!filtered.length) {
    grid.innerHTML = '<div class="empty-state">No events match the current filters.</div>';
    return;
  }

  let html = '';
  filtered.forEach(e => {
    const idx = eventsData.indexOf(e);
    const isLive = (e.status || '').toLowerCase() === 'live';
    const isFinished = (e.status || '').toLowerCase() === 'finished';
    const starStr = '⭐'.repeat(Math.min(5, Math.max(1, parseInt(e.stars) || 0)));
    const isMajor = e.major === 'yes' || e.major === true;

    // Use the event logo as a subtle background watermark (CSS var consumed by .ev-card::before)
    const bg = e.logo ? ` style="--ev-bg:url(&quot;${escHtml(e.logo)}&quot;)"` : '';
    // Standardized visual types (status/major) as explicit classes for easier styling
    html += `<div class="ev-card ${isLive ? 'ev-card-live' : ''} ${isFinished ? 'ev-card-finished' : ''} ${isMajor ? 'ev-card-major' : ''}"${bg}>`;

    // Top section: logo + info
    html += `<div class="ev-card-top">`;
    if (e.logo) {
      html += `<div class="ev-card-logo-wrap"><img src="${escHtml(e.logo)}" alt="" onerror="this.style.display='none'"></div>`;
    }
    html += `<div class="ev-card-info">`;
    // Event name (no drive link here — link is in details view)
    html += `<div class="ev-card-name">${escHtml(e.name || 'Unnamed Event')}</div>`;
    // Location
    if (e.city || e.country) {
      html += `<div class="ev-card-location">`;
      if (e.flag) html += `<img class="ev-card-flag" src="${escHtml(e.flag)}" alt="" onerror="this.style.display='none'">`;
      html += `${escHtml(e.city || '')}${e.city && e.country ? ', ' : ''}${escHtml(e.country || '')}`;
      html += `</div>`;
    }
    // Meta badges
    html += `<div class="ev-card-meta">`;
    if (isMajor) html += `<span class="ev-badge ev-badge-major">Major</span>`;
    if (isLive) html += `<span class="ev-badge ev-badge-live">● Live</span>`;
    if (isFinished) html += `<span class="ev-badge ev-badge-finished">Finished</span>`;
    if (e.stars) html += `<span class="ev-badge ev-badge-stars">${starStr}</span>`;
    if (e.teamCount) html += `<span class="ev-badge ev-badge-teams">${e.teamCount} teams</span>`;
    if (e.region) {
      const regionEmoji = { global: '🌍', americas: '🌎', asia: '🌏', europe: '🌍' };
      const regionName = { global: 'Global', americas: 'Americas', asia: 'Asia', europe: 'Europe' };
      html += `<span class="ev-badge ev-badge-region ev-badge-region-${escHtml(e.region)}">${regionEmoji[e.region] || '🌐'} ${escHtml(regionName[e.region] || e.region)}</span>`;
    }
    if (e.prizePool) html += `<span class="ev-badge ev-badge-prize">${escHtml(e.prizePool)}</span>`;
    html += `</div>`;
    html += `</div></div>`;

    // Body — depends on status
    html += `<div class="ev-card-body">`;
    if (isFinished) {
      // Winner with trophy
      if (e.winner) {
        html += `<div class="ev-card-result">`;
        if (e.trophy) html += `<div class="ev-card-trophy-wrap"><img src="${escHtml(e.trophy)}" alt="🏆" onerror="this.textContent='🏆'"></div>`;
        else html += `<div class="ev-card-trophy-wrap">🏆</div>`;
        html += `<div><div class="ev-result-label">Winner</div><div class="ev-result-value gold-text">${escHtml(e.winner)}</div></div>`;
        html += `</div>`;
      }
      if (e.second) {
        html += `<div class="ev-card-result" style="border-color:rgba(154,160,171,0.15)">`;
        html += `<div class="ev-card-trophy-wrap">🥈</div>`;
        html += `<div><div class="ev-result-label">2nd Place</div><div class="ev-result-value silver-text">${escHtml(e.second)}</div></div>`;
        html += `</div>`;
      }
      if (e.third) {
        html += `<div class="ev-card-result" style="border-color:rgba(176,122,67,0.15)">`;
        html += `<div class="ev-card-trophy-wrap">🥉</div>`;
        html += `<div><div class="ev-result-label">3rd Place</div><div class="ev-result-value bronze-text">${escHtml(e.third)}</div></div>`;
        html += `</div>`;
      }
      if (e.fourth) {
        html += `<div class="ev-card-result" style="border-color:rgba(138,130,117,0.15)">`;
        html += `<div class="ev-card-trophy-wrap">4️⃣</div>`;
        html += `<div><div class="ev-result-label">4th Place</div><div class="ev-result-value steel-text">${escHtml(e.fourth)}</div></div>`;
        html += `</div>`;
      }
      // Medals
      const medals = [];
      if (e.mvp) medals.push({ label: 'MVP', icon: '🥇', val: e.mvp });
      if (e.evp) medals.push({ label: 'EVP', icon: '🥈', val: e.evp });
      if (e.vp) medals.push({ label: 'VP', icon: '🥉', val: e.vp });
      if (medals.length) {
        html += `<div class="ev-card-medals">`;
        medals.forEach(m => {
          html += `<div class="ev-medal-row"><span class="ev-medal-label">${m.icon} ${m.label}</span><span class="ev-medal-value">${escHtml(m.val)}</span></div>`;
        });
        html += `</div>`;
      }
    } else if (isLive) {
      html += `<div class="ev-card-live-info">`;
      if (e.stage) {
        html += `<div class="ev-live-stage"><div class="ev-live-stage-label">Stage</div><div class="ev-live-stage-value">${escHtml(e.stage)}</div></div>`;
      }
      const nextMatches = [e.nextMatch1, e.nextMatch2, e.nextMatch3].filter(Boolean);
      if (nextMatches.length) {
        html += `<div class="ev-live-matches">`;
        nextMatches.forEach(m => {
          html += `<div class="ev-live-match-item">${escHtml(m)}</div>`;
        });
        html += `</div>`;
      }
      html += `</div>`;
    }
    html += `</div>`;

    // Actions
    html += `<div class="ev-card-actions">`;
    html += `<button class="ev-btn-details" onclick="openEventDetails(${idx})">📋 Details</button>`;
    html += `<button class="ev-btn-edit" onclick="openEditEvent(${idx})">✎ Edit</button>`;
    html += `<button class="ev-btn-delete" onclick="deleteEvent(${idx})">🗑 Delete</button>`;
    html += `</div>`;
    html += `</div>`;
  });

  grid.innerHTML = html;
}

/* ════════════════════════════════════════
   EVENT DETAILS MODAL
════════════════════════════════════════ */

function openEventDetails(idx) {
  const e = eventsData[idx];
  if (!e) return;
  const overlay = document.getElementById('evDetailOverlay');

  // Header
  const logoEl = document.getElementById('evdLogo');
  if (e.logo) { logoEl.src = e.logo; logoEl.style.display = ''; }
  else { logoEl.style.display = 'none'; }
  document.getElementById('evdTitle').textContent = e.name || 'Unnamed Event';
  const isLive = (e.status || '').toLowerCase() === 'live';
  const isMajor = e.major === 'yes' || e.major === true;
  const starStr = '⭐'.repeat(Math.min(5, Math.max(1, parseInt(e.stars) || 0)));
  let sub = '';
  if (e.city || e.country) sub += `${e.city || ''}${e.city && e.country ? ', ' : ''}${e.country || ''}`;
  if (e.prizePool) sub += (sub ? ' · ' : '') + e.prizePool;
  if (starStr) sub += (sub ? ' · ' : '') + starStr;
  if (isMajor) sub += (sub ? ' · ' : '') + 'Major';
  if (e.region) {
    const rNames = { global: '🌍 Global', americas: '🌎 Americas', asia: '🌏 Asia', europe: '🌍 Europe' };
    sub += (sub ? ' · ' : '') + (rNames[e.region] || e.region);
  }
  document.getElementById('evdSubtitle').textContent = sub;

  let body = '';

  // Drive link
  if (e.driveLink) {
    body += `<div class="evd-section">
      <a class="evd-drive-link" href="${escHtml(e.driveLink)}" target="_blank" rel="noopener">
        <span>📁</span> Open in Google Drive
      </a>
    </div>`;
  }

  // Results section
  if ((e.status || '').toLowerCase() === 'finished' && (e.winner || e.second || e.third || e.fourth)) {
    body += `<div class="evd-section evd-results-row">`;
    if (e.winner) body += `<div class="evd-result-pill gold-text"><span class="evd-result-place">🏆 1st Place</span><span class="evd-result-team">${escHtml(e.winner)}</span></div>`;
    if (e.second) body += `<div class="evd-result-pill silver-text"><span class="evd-result-place">🥈 2nd Place</span><span class="evd-result-team">${escHtml(e.second)}</span></div>`;
    if (e.third) body += `<div class="evd-result-pill bronze-text"><span class="evd-result-place">🥉 3rd Place</span><span class="evd-result-team">${escHtml(e.third)}</span></div>`;
    if (e.fourth) body += `<div class="evd-result-pill steel-text"><span class="evd-result-place">4️⃣ 4th Place</span><span class="evd-result-team">${escHtml(e.fourth)}</span></div>`;
    body += `</div>`;
    // Player medals
    const medals = [];
    if (e.mvp) medals.push({ icon:'🥇', label:'MVP', val: e.mvp, prize: (e.playerPrizes||{}).mvp });
    if (e.evp) medals.push({ icon:'🥈', label:'EVP', val: e.evp, prize: (e.playerPrizes||{}).evp });
    if (e.vp)  medals.push({ icon:'🥉', label:'VP',  val: e.vp,  prize: (e.playerPrizes||{}).vp  });
    if (medals.length) {
      body += `<div class="evd-medals-row">`;
      medals.forEach(m => {
        body += `<div class="evd-medal-card"><div class="evd-medal-icon">${m.icon}</div><div class="evd-medal-label">${m.label}</div><div class="evd-medal-name">${escHtml(m.val)}</div>${m.prize ? `<div class="evd-medal-prize">${escHtml(m.prize)}</div>` : ''}</div>`;
      });
      body += `</div>`;
    }
  }

  // Prize distribution
  const pd = e.prizeDistribution || {};
  const hasPrizes = Object.values(pd).some(v => v && v.trim());
  if (hasPrizes) {
    const placements = [
      { key:'1st',    label:'🥇 1st Place',    color:'#FFD700' },
      { key:'2nd',    label:'🥈 2nd Place',    color:'#C0C8D8' },
      { key:'3rd-4th',label:'🥉 3rd–4th',      color:'#CD7F32' },
      { key:'5th-8th',label:'4th–8th Place',   color:'var(--muted)' },
      { key:'9th-11th',label:'9th–11th Place', color:'var(--muted)' },
      { key:'12th-16th',label:'12th–16th',     color:'var(--muted)' },
      { key:'17th-24th',label:'17th–24th',     color:'var(--muted)' },
    ];
    body += `<div class="evd-section"><div class="evd-section-title">💰 Prize Distribution</div><div class="evd-prize-grid">`;
    placements.forEach(p => {
      if (pd[p.key]) {
        body += `<div class="evd-prize-row"><span class="evd-prize-place" style="color:${p.color}">${p.label}</span><span class="evd-prize-amount">${escHtml(pd[p.key])}</span></div>`;
      }
    });
    body += `</div></div>`;
  }

  // Playoff section
  const teams = e.playoffTeams || [];
  if (teams.length === 8) {
    body += `<div class="evd-section"><div class="evd-section-title">🏆 Playoff Bracket</div>`;

    // Playoff teams with hover players
    const players = e.playoffPlayers || {};
    body += `<div class="evd-playoff-teams-list">`;
    teams.forEach((t, ti) => {
      const tName = t.name || `Team ${ti+1}`;
      const tPlayers = players[tName] || [];
      body += `<div class="evd-pt-chip" data-team-idx="${ti}">`;
      if (t.logo) body += `<img src="${escHtml(t.logo)}" class="evd-pt-logo" alt="" onerror="this.style.display='none'">`;
      body += `<span class="evd-pt-seed">Seed ${ti+1}</span>`;
      body += `<span class="evd-pt-name">${escHtml(tName)}</span>`;
      if (tPlayers.filter(Boolean).length) {
        body += `<div class="evd-pt-players">`;
        tPlayers.forEach((p, pi) => {
          if (p) body += `<div class="evd-pt-player">${escHtml(p)}</div>`;
        });
        body += `</div>`;
      }
      body += `</div>`;
    });
    body += `</div>`;

    // Bracket drawing
    body += renderPlayoffBracket(e);
    body += `</div>`;
  }

  document.getElementById('evDetailBody').innerHTML = body;
  overlay.classList.add('visible');
}

function closeEventDetails() {
  document.getElementById('evDetailOverlay').classList.remove('visible');
}

/* Determine winner of a bracket match given scores */
function bracketWinner(t1, t2, score1, score2) {
  const s1 = parseInt(score1) || 0;
  const s2 = parseInt(score2) || 0;
  if (s1 > s2) return t1;
  if (s2 > s1) return t2;
  return null; // no result yet
}

/* Render a static playoff bracket (view-only) from event data */
function renderPlayoffBracket(e) {
  const teams = e.playoffTeams || [];
  if (teams.length !== 8) return '';
  const bracket = e.playoffBracket || {};
  const qf = bracket.qf || [{},{},{},{}];
  const sf = bracket.sf || [{},{}];
  const fn = bracket.final || {};

  const tName = i => teams[i] ? (teams[i].name || `Team ${i+1}`) : `Team ${i+1}`;
  const tLogo = i => teams[i] ? teams[i].logo : '';

  // Build a map: team name → logo
  const logoMap = {};
  teams.forEach((t, i) => { if (t && t.name && t.logo) logoMap[t.name] = t.logo; });

  // QF: 0v7, 1v6, 2v5, 3v4
  const qfPairs = [[0,7],[1,6],[2,5],[3,4]];
  const qfWinners = qfPairs.map((p, i) => bracketWinner(tName(p[0]), tName(p[1]), (qf[i]||{}).score1, (qf[i]||{}).score2));

  // SF: use stored team names if available, otherwise use QF winners
  const sfPairs = [[0,1],[2,3]];
  const sfTeams = sfPairs.map((p, i) => {
    const s = sf[i] || {};
    return {
      team1: s.team1 || qfWinners[p[0]] || '?',
      team2: s.team2 || qfWinners[p[1]] || '?',
      score1: s.score1,
      score2: s.score2
    };
  });
  const sfWinners = sfTeams.map(s => bracketWinner(s.team1, s.team2, s.score1, s.score2));

  // Final: use stored team names if available, otherwise use SF winners
  const fnTeams = {
    team1: fn.team1 || sfWinners[0] || '?',
    team2: fn.team2 || sfWinners[1] || '?',
    score1: fn.score1,
    score2: fn.score2
  };
  const champion = bracketWinner(fnTeams.team1, fnTeams.team2, fnTeams.score1, fnTeams.score2);

  const mkTeam = (name, score, isWinner, logo) => {
    const cls = isWinner ? 'bk-team bk-team-winner' : 'bk-team';
    let img = logo ? `<img src="${escHtml(logo)}" class="bk-team-logo" alt="" onerror="this.style.display='none'">` : '';
    return `<div class="${cls}">${img}<span class="bk-team-name">${escHtml(name || '?')}</span><span class="bk-team-score">${score !== undefined && score !== '' ? escHtml(String(score)) : '—'}</span></div>`;
  };

  const mkMatch = (t1, t2, s1, s2, logo1='', logo2='') => {
    const w = bracketWinner(t1, t2, s1, s2);
    return `<div class="bk-match">${mkTeam(t1, s1, w===t1, logo1)}${mkTeam(t2, s2, w===t2, logo2)}</div>`;
  };

  let h = `<div class="bk-bracket">`;

  // Column: QF
  h += `<div class="bk-col"><div class="bk-col-title">Quarter-Finals</div>`;
  qfPairs.forEach((p,i) => {
    const q = qf[i] || {};
    h += mkMatch(tName(p[0]), tName(p[1]), q.score1, q.score2, tLogo(p[0]), tLogo(p[1]));
  });
  h += `</div>`;

  // Column: SF (with logos from map)
  h += `<div class="bk-col"><div class="bk-col-title">Semi-Finals</div>`;
  sfTeams.forEach((s, i) => {
    h += mkMatch(s.team1, s.team2, s.score1, s.score2, logoMap[s.team1]||'', logoMap[s.team2]||'');
  });
  h += `</div>`;

  // Column: Final + Champion (with logos from map)
  h += `<div class="bk-col"><div class="bk-col-title">Grand Final</div>`;
  h += mkMatch(fnTeams.team1, fnTeams.team2, fnTeams.score1, fnTeams.score2, logoMap[fnTeams.team1]||'', logoMap[fnTeams.team2]||'');
  h += `</div>`;

  // Champion (with logo)
  const champLogo = champion ? (logoMap[champion] || '') : '';
  h += `<div class="bk-col bk-col-champion"><div class="bk-col-title">Champion</div>`;
  h += `<div class="bk-champion">${champion ? `${champLogo ? `<img src="${escHtml(champLogo)}" class="bk-team-logo" style="width:28px;height:28px;" alt="" onerror="this.style.display='none'">` : ''}<div class="bk-champion-icon">🏆</div><div class="bk-champion-name">${escHtml(champion)}</div>` : '<div class="bk-champion-name muted">TBD</div>'}</div>`;
  h += `</div>`;

  h += `</div>`; // .bk-bracket
  return h;
}

/* ─── Refresh entire events UI ─── */
function refreshEventsUI() {
  updateEventsStats();
  populateEventsFilters();
  renderEventsGrid();
}

/* ─── Modal open/close ─── */
function openEventModal(editIdx = -1) {
  evEditIndex = editIdx;
  const overlay = document.getElementById('evModalOverlay');
  document.getElementById('evModalTitle').textContent = editIdx >= 0 ? 'Edit Event' : 'Add Event';
  clearEventModal();

  if (editIdx >= 0 && eventsData[editIdx]) {
    const e = eventsData[editIdx];
    document.getElementById('evmName').value = e.name || '';
    document.getElementById('evmDriveLink').value = e.driveLink || '';
    document.getElementById('evmPrizePool').value = e.prizePool || '';
    document.getElementById('evmLogo').value = e.logo || '';
    document.getElementById('evmTrophy').value = e.trophy || '';
    document.getElementById('evmCity').value = e.city || '';
    document.getElementById('evmCountry').value = e.country || '';
    document.getElementById('evmFlag').value = e.flag || '';
    document.getElementById('evmTeamCount').value = e.teamCount || '';
    document.getElementById('evmStars').value = String(e.stars || 3);
    document.getElementById('evmMajor').value = (e.major === 'yes' || e.major === true) ? 'yes' : 'no';
    document.getElementById('evmRegion').value = e.region || 'global';
    document.getElementById('evmStatus').value = (e.status || 'Finished');
    toggleStatusFields();
    if ((e.status || '').toLowerCase() === 'live') {
      document.getElementById('evmStage').value = e.stage || '';
      document.getElementById('evmNext1').value = e.nextMatch1 || '';
      document.getElementById('evmNext2').value = e.nextMatch2 || '';
      document.getElementById('evmNext3').value = e.nextMatch3 || '';
    } else {
      document.getElementById('evmWinner').value = e.winner || '';
      document.getElementById('evmSecond').value = e.second || '';
      document.getElementById('evmThird').value = e.third || '';
      document.getElementById('evmFourth').value = e.fourth || '';
      document.getElementById('evmMVP').value = e.mvp || '';
      document.getElementById('evmEVP').value = e.evp || '';
      document.getElementById('evmVP').value = e.vp || '';
    }

    // Prize distribution
    const pd = e.prizeDistribution || {};
    document.getElementById('evmPrize1st').value  = pd['1st'] || '';
    document.getElementById('evmPrize2nd').value  = pd['2nd'] || '';
    document.getElementById('evmPrize34').value   = pd['3rd-4th'] || '';
    document.getElementById('evmPrize58').value   = pd['5th-8th'] || '';
    document.getElementById('evmPrize911').value  = pd['9th-11th'] || '';
    document.getElementById('evmPrize1216').value = pd['12th-16th'] || '';
    document.getElementById('evmPrize1724').value = pd['17th-24th'] || '';

    // Player prizes
    const pp = e.playerPrizes || {};
    document.getElementById('evmPrizeMVP').value = pp.mvp || '';
    document.getElementById('evmPrizeEVP').value = pp.evp || '';
    document.getElementById('evmPrizeVP').value  = pp.vp  || '';

    // Playoff teams & bracket
    loadPlayoffTeamsEdit(e.playoffTeams || []);
    loadPlayoffBracketEdit(e.playoffTeams || [], e.playoffBracket || {}, e.playoffPlayers || {});
  } else {
    loadPlayoffTeamsEdit([]);
    loadPlayoffBracketEdit([], {}, {});
  }

  // Always reset file picker
  const f = document.getElementById('evmLogoFile');
  if (f) f.value = '';
  updateEventLogoPreview();

  overlay.classList.add('visible');
}

function closeEventModal() {
  document.getElementById('evModalOverlay').classList.remove('visible');
  evEditIndex = -1;
}

function clearEventModal() {
  ['evmName','evmDriveLink','evmPrizePool','evmLogo','evmTrophy','evmCity','evmCountry',
   'evmFlag','evmTeamCount','evmStage','evmNext1','evmNext2','evmNext3',
   'evmWinner','evmSecond','evmThird','evmFourth','evmMVP','evmEVP','evmVP',
   'evmPrize1st','evmPrize2nd','evmPrize34','evmPrize58','evmPrize911','evmPrize1216','evmPrize1724',
   'evmPrizeMVP','evmPrizeEVP','evmPrizeVP'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  const f = document.getElementById('evmLogoFile');
  if (f) f.value = '';
  updateEventLogoPreview();
  document.getElementById('evmStars').value = '3';
  document.getElementById('evmMajor').value = 'no';
  document.getElementById('evmRegion').value = 'global';
  document.getElementById('evmStatus').value = 'Finished';
  toggleStatusFields();
}

/* ─── Logo upload / preview helpers ─── */
function updateEventLogoPreview() {
  const url = (document.getElementById('evmLogo')?.value || '').trim();
  const img = document.getElementById('evmLogoPreview');
  if (!img) return;
  if (!url) {
    img.style.display = 'none';
    img.removeAttribute('src');
    return;
  }
  img.src = url;
  img.style.display = '';
}

/* ─── Inject 3rd/4th Place fields next to "2nd Place" in the Add/Edit modal ───
   The form markup itself lives in index.html (not included here), so instead
   of requiring an HTML edit, we clone the "2nd Place" field's structure at
   runtime and insert two more right after it. Safe to call multiple times —
   it's a no-op once the fields exist. */
function ensurePlacementFields() {
  if (document.getElementById('evmThird')) return;
  const secondInput = document.getElementById('evmSecond');
  if (!secondInput) return;
  const secondGroup = secondInput.closest('.form-group') || secondInput.parentElement;
  if (!secondGroup) return;

  const makeGroup = (id, labelText) => {
    const group = document.createElement('div');
    group.className = secondGroup.className || 'form-group';
    const label = document.createElement('label');
    label.className = 'form-label';
    label.textContent = labelText;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'form-control';
    input.id = id;
    input.placeholder = 'Team name';
    group.appendChild(label);
    group.appendChild(input);
    return group;
  };

  const thirdGroup = makeGroup('evmThird', '3rd Place');
  const fourthGroup = makeGroup('evmFourth', '4th Place');
  secondGroup.insertAdjacentElement('afterend', thirdGroup);
  thirdGroup.insertAdjacentElement('afterend', fourthGroup);
}
ensurePlacementFields();

function toggleStatusFields() {
  const status = document.getElementById('evmStatus').value;
  document.getElementById('evmLiveFields').style.display = (status === 'Live') ? 'block' : 'none';
  document.getElementById('evmFinishedFields').style.display = (status === 'Finished') ? 'block' : 'none';
}

/* ─── Save event ─── */
function saveEvent() {
  const status = document.getElementById('evmStatus').value;
  const ev = {
    name: document.getElementById('evmName').value.trim(),
    driveLink: document.getElementById('evmDriveLink').value.trim(),
    prizePool: document.getElementById('evmPrizePool').value.trim(),
    logo: document.getElementById('evmLogo').value.trim(),
    trophy: document.getElementById('evmTrophy').value.trim(),
    city: document.getElementById('evmCity').value.trim(),
    country: document.getElementById('evmCountry').value.trim(),
    flag: document.getElementById('evmFlag').value.trim(),
    teamCount: parseInt(document.getElementById('evmTeamCount').value) || 0,
    stars: parseInt(document.getElementById('evmStars').value) || 3,
    major: document.getElementById('evmMajor').value,
    region: document.getElementById('evmRegion').value,
    status: status
  };

  if (status === 'Live') {
    ev.stage = document.getElementById('evmStage').value.trim();
    ev.nextMatch1 = document.getElementById('evmNext1').value.trim();
    ev.nextMatch2 = document.getElementById('evmNext2').value.trim();
    ev.nextMatch3 = document.getElementById('evmNext3').value.trim();
    ev.winner = ''; ev.second = ''; ev.third = ''; ev.fourth = ''; ev.mvp = ''; ev.evp = ''; ev.vp = '';
  } else {
    ev.winner = document.getElementById('evmWinner').value.trim();
    ev.second = document.getElementById('evmSecond').value.trim();
    ev.third = document.getElementById('evmThird').value.trim();
    ev.fourth = document.getElementById('evmFourth').value.trim();
    ev.mvp = document.getElementById('evmMVP').value.trim();
    ev.evp = document.getElementById('evmEVP').value.trim();
    ev.vp = document.getElementById('evmVP').value.trim();
    ev.stage = ''; ev.nextMatch1 = ''; ev.nextMatch2 = ''; ev.nextMatch3 = '';
  }

  // Prize distribution
  ev.prizeDistribution = {
    '1st':       document.getElementById('evmPrize1st').value.trim(),
    '2nd':       document.getElementById('evmPrize2nd').value.trim(),
    '3rd-4th':   document.getElementById('evmPrize34').value.trim(),
    '5th-8th':   document.getElementById('evmPrize58').value.trim(),
    '9th-11th':  document.getElementById('evmPrize911').value.trim(),
    '12th-16th': document.getElementById('evmPrize1216').value.trim(),
    '17th-24th': document.getElementById('evmPrize1724').value.trim(),
  };

  // Player prizes
  ev.playerPrizes = {
    mvp: document.getElementById('evmPrizeMVP').value.trim(),
    evp: document.getElementById('evmPrizeEVP').value.trim(),
    vp:  document.getElementById('evmPrizeVP').value.trim(),
  };

  // Playoff teams
  ev.playoffTeams = collectPlayoffTeams();

  // Playoff bracket
  ev.playoffBracket = collectPlayoffBracket();

  // Playoff players
  ev.playoffPlayers = collectPlayoffPlayers(ev.playoffTeams);

  if (!ev.name) {
    alert('Event name is required.');
    return;
  }

  if (evEditIndex >= 0) {
    eventsData[evEditIndex] = ev;
  } else {
    eventsData.push(ev);
  }

  closeEventModal();
  refreshEventsUI();
}

/* ─── Playoff edit helpers ─── */
function loadPlayoffTeamsEdit(teams) {
  const grid = document.getElementById('evmPlayoffTeamsGrid');
  if (!grid) return;
  let html = '';
  for (let i = 0; i < 8; i++) {
    const t = teams[i] || {};
    const hasLogo = !!(t.logo);
    html += `<div class="ev-playoff-team-row">
      <span class="ev-playoff-seed">Seed ${i+1}</span>
      <input class="form-control ev-pt-name" type="text" placeholder="Team name" value="${escHtml(t.name||'')}" data-pt-idx="${i}" data-pt-field="name">
      <div class="ev-pt-logo-wrap">
        <input class="form-control ev-pt-logo" type="text" placeholder="Logo URL" value="${escHtml(t.logo||'')}" data-pt-idx="${i}" data-pt-field="logo">
        <label class="ev-pt-upload-btn" title="Upload image from computer">
          📁
          <input type="file" accept="image/*" data-pt-upload="${i}" style="display:none;">
        </label>
        <button type="button" class="ev-pt-clear-btn" data-pt-clear="${i}" title="Clear logo" style="${hasLogo ? '' : 'display:none;'}">✕</button>
        <img class="ev-pt-logo-preview" data-pt-preview="${i}" src="${hasLogo ? escHtml(t.logo) : ''}" alt="" style="${hasLogo ? '' : 'display:none;'}" onerror="this.style.display='none'">
      </div>
    </div>`;
  }
  grid.innerHTML = html;
  // Update bracket labels whenever team names change
  grid.querySelectorAll('.ev-pt-name').forEach(inp => {
    inp.addEventListener('input', () => refreshBracketEditLabels());
  });
  // File upload handlers (delegated)
  grid.querySelectorAll('[data-pt-upload]').forEach(fileInp => {
    fileInp.addEventListener('change', () => {
      const idx = fileInp.getAttribute('data-pt-upload');
      const file = fileInp.files && fileInp.files[0];
      if (!file) return;
      if (!file.type || !file.type.startsWith('image/')) {
        alert('Please select an image file.');
        fileInp.value = '';
        return;
      }
      const reader = new FileReader();
      reader.onload = () => {
        const logoInput = grid.querySelector(`[data-pt-idx="${idx}"][data-pt-field="logo"]`);
        if (logoInput) logoInput.value = String(reader.result || '');
        updatePtLogoPreview(idx);
      };
      reader.readAsDataURL(file);
    });
  });
  // Clear button handlers
  grid.querySelectorAll('[data-pt-clear]').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = btn.getAttribute('data-pt-clear');
      const logoInput = grid.querySelector(`[data-pt-idx="${idx}"][data-pt-field="logo"]`);
      if (logoInput) logoInput.value = '';
      const fileInp = grid.querySelector(`[data-pt-upload="${idx}"]`);
      if (fileInp) fileInp.value = '';
      updatePtLogoPreview(idx);
    });
  });
  // URL input change → update preview
  grid.querySelectorAll('.ev-pt-logo').forEach(inp => {
    inp.addEventListener('input', () => {
      const idx = inp.getAttribute('data-pt-idx');
      updatePtLogoPreview(idx);
    });
  });
}

function updatePtLogoPreview(idx) {
  const grid = document.getElementById('evmPlayoffTeamsGrid');
  if (!grid) return;
  const logoInput = grid.querySelector(`[data-pt-idx="${idx}"][data-pt-field="logo"]`);
  const preview = grid.querySelector(`[data-pt-preview="${idx}"]`);
  const clearBtn = grid.querySelector(`[data-pt-clear="${idx}"]`);
  const val = logoInput ? logoInput.value.trim() : '';
  if (preview) {
    if (val) { preview.src = val; preview.style.display = ''; }
    else { preview.src = ''; preview.style.display = 'none'; }
  }
  if (clearBtn) {
    clearBtn.style.display = val ? '' : 'none';
  }
}

function getPlayoffTeamNames() {
  const names = [];
  for (let i = 0; i < 8; i++) {
    const inp = document.querySelector(`[data-pt-idx="${i}"][data-pt-field="name"]`);
    names.push(inp ? inp.value.trim() || `Team ${i+1}` : `Team ${i+1}`);
  }
  return names;
}

function loadPlayoffBracketEdit(teams, bracket, players) {
  const qf = bracket.qf || [{},{},{},{}];
  const sf = bracket.sf || [{},{}];
  const fn = bracket.final || {};
  const names = teams.length === 8 ? teams.map(t => t ? (t.name||'?') : '?') : Array.from({length:8}, (_,i)=>`Team ${i+1}`);

  // QF matches: 0v7, 1v6, 2v5, 3v4
  const qfPairs = [[0,7],[1,6],[2,5],[3,4]];
  let qfHtml = '';
  qfPairs.forEach((p, i) => {
    const q = qf[i] || {};
    qfHtml += `<div class="ev-bk-match-row">
      <span class="ev-bk-team-label bk-label-${i}-0">${escHtml(names[p[0]])}</span>
      <input class="form-control ev-bk-score" type="text" placeholder="0" maxlength="2" value="${escHtml(String(q.score1||''))}" data-round="qf" data-match="${i}" data-pos="score1">
      <span class="ev-bk-vs">–</span>
      <input class="form-control ev-bk-score" type="text" placeholder="0" maxlength="2" value="${escHtml(String(q.score2||''))}" data-round="qf" data-match="${i}" data-pos="score2">
      <span class="ev-bk-team-label bk-label-${i}-1">${escHtml(names[p[1]])}</span>
    </div>`;
  });
  document.getElementById('evmQFMatches').innerHTML = qfHtml;

  // SF — editable team name inputs
  let sfHtml = '';
  [[0,1],[2,3]].forEach((p,i) => {
    const s = sf[i] || {};
    const defaultT1 = s.team1 || (i===0 ? 'QF1 Winner' : 'QF3 Winner');
    const defaultT2 = s.team2 || (i===0 ? 'QF2 Winner' : 'QF4 Winner');
    sfHtml += `<div class="ev-bk-match-row">
      <input class="form-control ev-bk-team-input" type="text" placeholder="${i===0?'QF1 Winner':'QF3 Winner'}" value="${escHtml(defaultT1!==`QF${i===0?1:3} Winner`?defaultT1:'')}" data-round="sf" data-match="${i}" data-pos="team1">
      <input class="form-control ev-bk-score" type="text" placeholder="0" maxlength="2" value="${escHtml(String(s.score1||''))}" data-round="sf" data-match="${i}" data-pos="score1">
      <span class="ev-bk-vs">–</span>
      <input class="form-control ev-bk-score" type="text" placeholder="0" maxlength="2" value="${escHtml(String(s.score2||''))}" data-round="sf" data-match="${i}" data-pos="score2">
      <input class="form-control ev-bk-team-input" type="text" placeholder="${i===0?'QF2 Winner':'QF4 Winner'}" value="${escHtml(defaultT2!==`QF${i===0?2:4} Winner`?defaultT2:'')}" data-round="sf" data-match="${i}" data-pos="team2">
    </div>`;
  });
  document.getElementById('evmSFMatches').innerHTML = sfHtml;

  // Final — editable team name inputs
  const fDefaultT1 = fn.team1 || '';
  const fDefaultT2 = fn.team2 || '';
  const fHtml = `<div class="ev-bk-match-row">
    <input class="form-control ev-bk-team-input" type="text" placeholder="SF1 Winner" value="${escHtml(fDefaultT1)}" data-round="final" data-match="0" data-pos="team1">
    <input class="form-control ev-bk-score" type="text" placeholder="0" maxlength="2" value="${escHtml(String(fn.score1||''))}" data-round="final" data-match="0" data-pos="score1">
    <span class="ev-bk-vs">–</span>
    <input class="form-control ev-bk-score" type="text" placeholder="0" maxlength="2" value="${escHtml(String(fn.score2||''))}" data-round="final" data-match="0" data-pos="score2">
    <input class="form-control ev-bk-team-input" type="text" placeholder="SF2 Winner" value="${escHtml(fDefaultT2)}" data-round="final" data-match="0" data-pos="team2">
  </div>`;
  document.getElementById('evmFinalMatch').innerHTML = fHtml;

  // Players
  loadPlayoffPlayersEdit(teams, players);
}

function refreshBracketEditLabels() {
  const names = getPlayoffTeamNames();
  const qfPairs = [[0,7],[1,6],[2,5],[3,4]];
  qfPairs.forEach((p, i) => {
    const l0 = document.querySelector(`.bk-label-${i}-0`);
    const l1 = document.querySelector(`.bk-label-${i}-1`);
    if (l0) l0.textContent = names[p[0]];
    if (l1) l1.textContent = names[p[1]];
  });
}

function loadPlayoffPlayersEdit(teams, players) {
  const sec = document.getElementById('evmPlayersSection');
  if (!sec) return;
  let html = '';
  for (let i = 0; i < 8; i++) {
    const t = teams[i] || {};
    const tName = t.name || `Team ${i+1}`;
    const tPlayers = players[tName] || [];
    html += `<div class="ev-players-team-block">
      <div class="ev-players-team-name" data-players-idx="${i}">Seed ${i+1}: <span class="ev-players-team-ref">${escHtml(tName)}</span></div>
      <div class="ev-players-inputs">`;
    for (let j = 0; j < 5; j++) {
      html += `<input class="form-control ev-player-inp" type="text" placeholder="Player ${j+1}" value="${escHtml((tPlayers[j]||''))}" data-team-idx="${i}" data-player-idx="${j}">`;
    }
    html += `</div></div>`;
  }
  sec.innerHTML = html;
}

function collectPlayoffTeams() {
  const teams = [];
  for (let i = 0; i < 8; i++) {
    const nameEl = document.querySelector(`[data-pt-idx="${i}"][data-pt-field="name"]`);
    const logoEl = document.querySelector(`[data-pt-idx="${i}"][data-pt-field="logo"]`);
    teams.push({
      name: nameEl ? nameEl.value.trim() : '',
      logo: logoEl ? logoEl.value.trim() : ''
    });
  }
  return teams;
}

function collectPlayoffBracket() {
  const bracket = { qf: [], sf: [], final: {} };
  // QF
  for (let i = 0; i < 4; i++) {
    const s1 = document.querySelector(`[data-round="qf"][data-match="${i}"][data-pos="score1"]`);
    const s2 = document.querySelector(`[data-round="qf"][data-match="${i}"][data-pos="score2"]`);
    bracket.qf.push({ score1: s1 ? s1.value.trim() : '', score2: s2 ? s2.value.trim() : '' });
  }
  // SF (with team names)
  for (let i = 0; i < 2; i++) {
    const t1 = document.querySelector(`[data-round="sf"][data-match="${i}"][data-pos="team1"]`);
    const t2 = document.querySelector(`[data-round="sf"][data-match="${i}"][data-pos="team2"]`);
    const s1 = document.querySelector(`[data-round="sf"][data-match="${i}"][data-pos="score1"]`);
    const s2 = document.querySelector(`[data-round="sf"][data-match="${i}"][data-pos="score2"]`);
    bracket.sf.push({
      team1: t1 ? t1.value.trim() : '',
      team2: t2 ? t2.value.trim() : '',
      score1: s1 ? s1.value.trim() : '',
      score2: s2 ? s2.value.trim() : ''
    });
  }
  // Final (with team names)
  const ft1 = document.querySelector(`[data-round="final"][data-match="0"][data-pos="team1"]`);
  const ft2 = document.querySelector(`[data-round="final"][data-match="0"][data-pos="team2"]`);
  const fs1 = document.querySelector(`[data-round="final"][data-match="0"][data-pos="score1"]`);
  const fs2 = document.querySelector(`[data-round="final"][data-match="0"][data-pos="score2"]`);
  bracket.final = {
    team1: ft1 ? ft1.value.trim() : '',
    team2: ft2 ? ft2.value.trim() : '',
    score1: fs1 ? fs1.value.trim() : '',
    score2: fs2 ? fs2.value.trim() : ''
  };
  return bracket;
}

function collectPlayoffPlayers(playoffTeams) {
  const result = {};
  for (let i = 0; i < 8; i++) {
    const tName = (playoffTeams[i] && playoffTeams[i].name) || `Team ${i+1}`;
    const ps = [];
    for (let j = 0; j < 5; j++) {
      const inp = document.querySelector(`[data-team-idx="${i}"][data-player-idx="${j}"]`);
      ps.push(inp ? inp.value.trim() : '');
    }
    result[tName] = ps;
  }
  return result;
}

/* ─── Event Details listeners ─── */
document.getElementById('evDetailClose').addEventListener('click', closeEventDetails);
document.getElementById('evDetailOverlay').addEventListener('click', (e) => {
  if (e.target === document.getElementById('evDetailOverlay')) closeEventDetails();
});

/* ─── Edit / Delete ─── */
function openEditEvent(idx) {
  openEventModal(idx);
}

function deleteEvent(idx) {
  if (!confirm('Delete this event?')) return;
  eventsData.splice(idx, 1);
  refreshEventsUI();
}

/* ─── Download JSON ─── */
function downloadEventsJson() {
  const blob = new Blob([JSON.stringify(eventsData, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'kz-events-backup.json';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* ─── Event listeners ─── */
document.getElementById('btnAddEvent').addEventListener('click', () => openEventModal(-1));
document.getElementById('evModalClose').addEventListener('click', closeEventModal);
document.getElementById('evModalCancel').addEventListener('click', closeEventModal);
document.getElementById('evModalSave').addEventListener('click', saveEvent);
document.getElementById('evmStatus').addEventListener('change', toggleStatusFields);
document.getElementById('btnDownloadEventsJson').addEventListener('click', downloadEventsJson);

// Logo input + upload
const evmLogoInput = document.getElementById('evmLogo');
if (evmLogoInput) evmLogoInput.addEventListener('input', updateEventLogoPreview);

const evmLogoFile = document.getElementById('evmLogoFile');
if (evmLogoFile) {
  evmLogoFile.addEventListener('change', () => {
    const file = evmLogoFile.files && evmLogoFile.files[0];
    if (!file) return;
    if (!file.type || !file.type.startsWith('image/')) {
      alert('Please select an image file.');
      evmLogoFile.value = '';
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      document.getElementById('evmLogo').value = String(reader.result || '');
      updateEventLogoPreview();
    };
    reader.readAsDataURL(file);
  });
}

const evmLogoClear = document.getElementById('evmLogoClear');
if (evmLogoClear) {
  evmLogoClear.addEventListener('click', () => {
    const t = document.getElementById('evmLogo');
    if (t) t.value = '';
    const f = document.getElementById('evmLogoFile');
    if (f) f.value = '';
    updateEventLogoPreview();
  });
}

// Close modal on overlay click
document.getElementById('evModalOverlay').addEventListener('click', (e) => {
  if (e.target === document.getElementById('evModalOverlay')) closeEventModal();
});

// Filters
['evFilterName', 'evFilterPlayer', 'evFilterTeamWin', 'evFilterStars', 'evFilterRegion'].forEach(id => {
  const el = document.getElementById(id);
  if (el) el.addEventListener(el.tagName === 'INPUT' ? 'input' : 'change', renderEventsGrid);
});

document.getElementById('btnClearEvFilters').addEventListener('click', () => {
  document.getElementById('evFilterName').value = '';
  document.getElementById('evFilterPlayer').value = '';
  document.getElementById('evFilterTeamWin').value = '';
  document.getElementById('evFilterStars').value = '';
  document.getElementById('evFilterRegion').value = '';
  renderEventsGrid();
});

/* ════════════════════════════════════════
   AUTO-LOAD
════════════════════════════════════════ */
(async () => {
  try { await loadTeamLogos(); }
  catch (e) { console.warn('Team logos load failed:', e); }
  try { await loadAll(false); }
  catch (e) { console.warn('Auto-load failed:', e); setStatus('Auto-load failed. Ensure /file/*.xlsx exists.', 'err'); }
  try { await loadArchiveManifest(); }
  catch (e) { console.warn('Archive manifest load failed:', e); }
  try { await loadEventsJson(); }
  catch (e) { console.warn('Events load failed:', e); }
})();

/* ════════════════════════════════════════
   SIDEBAR FULL-HIDE (full-screen content mode)
   « button hides the menu entirely; a small floating »
   button in the top-left corner (or Ctrl+B) brings it back.
   State persists across reloads via localStorage.
════════════════════════════════════════ */
(() => {
  const HIDE_KEY = 'nexus_sidebar_hidden';
  const hideBtn = $('#sidebarHideBtn');
  const showBtn = $('#sidebarShowBtn');
  if (!hideBtn || !showBtn || !appShell) return;

  function resizeCharts() {
    // Same pattern as the compact toggle: wait for the 0.3s
    // grid transition to finish, then let Chart.js re-measure.
    setTimeout(() => {
      if (typeof charts === 'object' && charts) {
        Object.values(charts).forEach(c => { if (c) c.resize(); });
      }
    }, 350);
  }

  function setSidebarHidden(hidden, persist = true) {
    appShell.classList.toggle('sidebar-hidden', hidden);
    document.body.classList.toggle('sidebar-hidden-mode', hidden);
    if (persist) {
      try { localStorage.setItem(HIDE_KEY, hidden ? '1' : '0'); } catch (e) {}
    }
    resizeCharts();
  }

  hideBtn.addEventListener('click', () => setSidebarHidden(true));
  showBtn.addEventListener('click', () => setSidebarHidden(false));

  // Ctrl+B / Cmd+B toggles, VS Code style — but never while the
  // user is typing in an input, select or textarea.
  document.addEventListener('keydown', (e) => {
    if (!(e.ctrlKey || e.metaKey) || e.key.toLowerCase() !== 'b') return;
    const tag = (document.activeElement?.tagName || '').toLowerCase();
    if (tag === 'input' || tag === 'textarea' || tag === 'select') return;
    e.preventDefault();
    setSidebarHidden(!appShell.classList.contains('sidebar-hidden'));
  });

  // Restore previous state on load (no animation flash: the class
  // lands before first paint since this runs synchronously).
  try {
    if (localStorage.getItem(HIDE_KEY) === '1') setSidebarHidden(true, false);
  } catch (e) {}
})();
