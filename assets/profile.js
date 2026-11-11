/* ═══════════════════════════════════════════════════════════════
   NEXUS · Team page
   Read-only view. Sources (same files the other pages already use):
     file/ranking.xlsx           position, points, record, tier, region
     file/players_updated.xlsx   roster, roles, country, market value, medals
     file/history_updated.xlsx   TeamHistory + PlayerHistory (achievements)
     file/history.xlsx           match results, one row per map (optional)
     file/old.xlsx               previous ranking, for movement (optional)
     file/archives/              older ranking snapshots, for the chart (optional)
     file/kz-events-backup.json  event logos, places and full standings (optional, large)
     assets/team_logo.json       team logos (optional)
     assets/awards.json          trophy / medal icons (optional, large)
     assets/players/<Name>.png   player photos (optional)
═══════════════════════════════════════════════════════════════ */
const $ = sel => document.querySelector(sel);
const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');

const URLS = {
  ranking: 'file/ranking.xlsx',
  old:     'file/old.xlsx',
  archiveDir: 'file/archives/',
  events:  'file/kz-events-backup.json',
  players: 'file/players_updated.xlsx',
  history: 'file/history_updated.xlsx',
  matches: 'file/history.xlsx',
  logos:   'assets/team_logo.json',
  awards:  'assets/awards.json'
};
const MATCH_LIMIT = 15;
const REGION_LABELS = { 'AM': 'Americas', 'EU': 'Europe', 'AS/SIS/ESEA': 'Asia' };
const ROLE_LABELS = { 'igl': 'IGL', 'rifler': 'Rifler', 'entry fragger': 'Entry Fragger', 'star player': 'Star Player', 'awper': 'AWPer', 'anchor': 'Anchor' };
const COUNTRY_CODES = {"brazil":"br","brasil":"br","united states":"us","usa":"us","us":"us","canada":"ca","argentina":"ar","chile":"cl","colombia":"co","mexico":"mx","peru":"pe","uruguay":"uy","venezuela":"ve","ecuador":"ec","bolivia":"bo","paraguay":"py","france":"fr","germany":"de","sweden":"se","denmark":"dk","finland":"fi","norway":"no","ukraine":"ua","poland":"pl","czech republic":"cz","czechia":"cz","slovakia":"sk","hungary":"hu","portugal":"pt","spain":"es","italy":"it","netherlands":"nl","belgium":"be","austria":"at","switzerland":"ch","uk":"gb","united kingdom":"gb","england":"gb","latvia":"lv","lithuania":"lt","estonia":"ee","serbia":"rs","croatia":"hr","romania":"ro","bulgaria":"bg","greece":"gr","turkey":"tr","turkiye":"tr","iceland":"is","ireland":"ie","georgia":"ge","moldova":"md","north macedonia":"mk","slovenia":"si","bosnia":"ba","bosnia and herzegovina":"ba","albania":"al","montenegro":"me","belarus":"by","kosovo":"xk","luxembourg":"lu","south africa":"za","israel":"il","jordan":"jo","lebanon":"lb","palestine":"ps","russia":"ru","kazakhstan":"kz","uzbekistan":"uz","china":"cn","south korea":"kr","korea":"kr","japan":"jp","australia":"au","new zealand":"nz","indonesia":"id","malaysia":"my","singapore":"sg","thailand":"th","vietnam":"vn","philippines":"ph","india":"in","pakistan":"pk","mongolia":"mn","taiwan":"tw","hong kong":"hk","iraq":"iq","uae":"ae","united arab emirates":"ae","guatemala":"gt","united kingom":"gb"};
const PHOTO_EXTS = ['png', 'jpg', 'jpeg', 'webp'];

const state = {
  teamParam: new URLSearchParams(location.search).get('team') || '',
  ranking: [], oldRanking: null, team: null, player: null, trajectory: [],
  series: null, events: null, transactions: [],
  players: [], playersLoaded: false,
  teamHistory: [], playerHistory: [], historyLoaded: false,
  matches: null,            // null = file not available
  logos: {}, awards: null, historyIcons: {},
  openPlayer: null
};
const sameTeam = (a, b) => String(a).trim().toLowerCase() === String(b).trim().toLowerCase();

/* ── formatting ── */
const fmtMoney = v => {
  v = Number(v) || 0;
  const a = Math.abs(v), s = v < 0 ? '-' : '';
  return s + (a >= 1000 ? '$' + (a / 1000).toFixed(1) + 'k' : '$' + a.toLocaleString('en-US', { maximumFractionDigits: 0 }));
};
const fmtPts = v => (Number(v) || 0).toLocaleString('en-US', { maximumFractionDigits: 0 });

function setStatus(msg, cls) {
  $('#statusText').textContent = msg;
  $('#statusDot').className = 'status-dot ' + (cls || 'loading');
}

/* ── loading ── */
async function fetchWorkbook(url) {
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error(url + ' → HTTP ' + res.status);
  return XLSX.read(await res.arrayBuffer(), { type: 'array' });
}
const sheetRows = (wb, name) => {
  const ws = name ? wb.Sheets[name] : wb.Sheets[wb.SheetNames[0]];
  return ws ? XLSX.utils.sheet_to_json(ws, { defval: '', raw: true }) : [];
};
/* Case-insensitive column lookup, so small header differences don't break the page */
function col(row, ...names) {
  for (const n of names) if (row[n] !== undefined && row[n] !== '') return row[n];
  const keys = Object.keys(row);
  for (const n of names) {
    const k = keys.find(k => k.trim().toLowerCase() === n.toLowerCase());
    if (k !== undefined && row[k] !== '') return row[k];
  }
  return '';
}
const num = v => { const n = Number(String(v).replace(',', '.')); return Number.isFinite(n) ? n : 0; };

/* Same ordering rule as the Global Ranking: points, then name. Adds global + regional position. */
function parseRanking(rows) {
  const teams = rows.map(r => ({
    team: String(col(r, 'Team')).trim(),
    points: num(col(r, 'Points')),
    victories: num(col(r, 'Victories')), loses: num(col(r, 'Loses')), streak: num(col(r, 'Streak')),
    majors: num(col(r, 'Majors')),
    tier: String(col(r, 'Tier')).trim(), region: String(col(r, 'Region')).trim()
  })).filter(t => t.team);
  teams.sort((a, b) => b.points !== a.points ? b.points - a.points : a.team.localeCompare(b.team, 'en'));
  const regionCount = {};
  teams.forEach((t, i) => { t.pos = i + 1; t.regionPos = regionCount[t.region] = (regionCount[t.region] || 0) + 1; });
  return teams;
}
async function loadRanking() { state.ranking = parseRanking(sheetRows(await fetchWorkbook(URLS.ranking))); }
async function loadOldRanking() { state.oldRanking = parseRanking(sheetRows(await fetchWorkbook(URLS.old))); }

/* Position over time: archived snapshots (file/archives/manifest.json → r0.xlsx, r1.xlsx …),
   then old.xlsx, then the live ranking. Missing archive → just old + live. */
async function loadTrajectory() {
  let manifest = { archives: [], old: { date: null } };
  try {
    const res = await fetch(URLS.archiveDir + 'manifest.json', { cache: 'no-store' });
    if (res.ok) manifest = await res.json();
  } catch (e) {}
  const entries = (manifest.archives || []).filter(a => a && a.file)
    .sort((a, b) => String(a.date || '').localeCompare(String(b.date || '')));
  const loaded = await Promise.allSettled(entries.map(a => fetchWorkbook(URLS.archiveDir + a.file).then(wb => parseRanking(sheetRows(wb)))));
  const points = [];
  const add = (teams, label) => {
    const t = teams && teams.find(x => sameTeam(x.team, state.team.team));
    if (t) points.push({ label, pos: t.pos, regionPos: t.regionPos, points: t.points });
  };
  const shortDate = d => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(d || '')); return m ? `${m[3]}/${m[2]}` : ''; };
  loaded.forEach((r, i) => { if (r.status === 'fulfilled') add(r.value, shortDate(entries[i].date) || entries[i].file.replace(/\.xlsx$/i, '')); });
  add(state.oldRanking, shortDate(manifest.old && manifest.old.date) || 'Previous');
  add(state.ranking, 'Now');
  state.trajectory = points;
}

async function loadPlayers() {
  const rows = sheetRows(await fetchWorkbook(URLS.players));
  state.players = rows.map(r => ({
    name: String(col(r, 'PlayerName')).trim(),
    team: String(col(r, 'TeamName')).trim(),
    roster: String(col(r, 'Roster')).trim(),
    roles: [...new Set(String(col(r, 'Role')).split('|').map(s => s.trim().toLowerCase()).filter(Boolean))],
    country: String(col(r, 'Country')).trim(),
    marketValue: num(col(r, 'Market Value', 'MarketValue')),
    mvp: num(col(r, 'MVP')), evp: num(col(r, 'EVP')), vp: num(col(r, 'VP')),
    titles: num(col(r, 'Titles')), majorWinner: num(col(r, 'Major Winner'))
  })).filter(p => p.name);
  state.playersLoaded = true;
}

const normScope = v => { const s = String(v ?? '').trim(); if (!s) return 'Global'; return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase(); };
async function loadHistory() {
  const wb = await fetchWorkbook(URLS.history);
  state.teamHistory = sheetRows(wb, 'TeamHistory').filter(r => r.event).map(r => ({
    id: String(r.id), team: String(r.team || '').trim(), event: String(r.event), tier: Number(r.tier) || 1,
    scope: normScope(r.scope), compRegion: String(r.compRegion || '').trim(), position: String(r.position || 'Winner').trim(),
    prizePool: Number(r.prizePool) || 0, wonAmount: Number(r.wonAmount) || 0, month: Number(r.month) || 1,
    kind: String(r.kind || 'EVENT').trim(), awardKey: String(r.awardKey || '').trim() || null, note: String(r.note || '').trim()
  }));
  state.playerHistory = sheetRows(wb, 'PlayerHistory').filter(r => r.player && r.event).map(r => ({
    id: String(r.id), refId: r.refId === '' || r.refId == null ? null : String(r.refId),
    player: String(r.player).trim(), team: String(r.team || '').trim(), kind: String(r.kind || 'EVENT').trim(),
    event: String(r.event), tier: Number(r.tier) || 1, scope: normScope(r.scope), compRegion: String(r.compRegion || '').trim(),
    position: String(r.position || 'Winner').trim(), month: Number(r.month) || 1,
    awardKey: String(r.awardKey || '').trim() || null, note: String(r.note || '').trim()
  }));
  state.transactions = sheetRows(wb, 'Transactions').filter(r => r.type && r.playerName).map(r => ({
    id: Number(r.id) || 0, month: Number(r.month) || 1, type: String(r.type).trim(), player: String(r.playerName).trim(),
    from: String(r.from || '').trim(), to: String(r.to || '').trim(), amount: Number(r.amount) || 0, note: String(r.note || '').trim()
  }));
  state.historyLoaded = true;
}

async function loadMatches() {
  const rows = sheetRows(await fetchWorkbook(URLS.matches));
  const out = [];
  let rowIdx = 0;
  for (const r of rows) {
    const team1 = String(col(r, 'Team 1', 'team1')).trim(), team2 = String(col(r, 'Team 2', 'team2')).trim();
    const result = String(col(r, 'Result')).trim();
    if (!team1 || !team2 || !result) continue;
    const dateRaw = col(r, 'Date');
    let dateObj = null;
    if (dateRaw !== '') {
      dateObj = typeof dateRaw === 'number' ? new Date((dateRaw - 25569) * 86400000) : new Date(dateRaw);
      if (isNaN(dateObj.getTime())) dateObj = null;
    }
    const parts = result.toLowerCase().split('x').map(s => parseInt(s.trim(), 10));
    out.push({
      idx: rowIdx++, team1, team2, result, score1: parts[0] || 0, score2: parts[1] || 0,
      map: String(col(r, 'Map')).trim(), event: String(col(r, 'Event')).trim(), details: String(col(r, 'Details')).trim(),
      dateObj, dateStr: dateObj ? dateObj.toISOString().split('T')[0] : String(dateRaw).trim()
    });
  }
  out.sort((a, b) => (a.dateObj && b.dateObj) ? b.dateObj - a.dateObj : b.dateStr.localeCompare(a.dateStr));
  state.matches = out;
  state.series = buildSeries(out);
}

/* Event names differ slightly between files ("DreamHack Winter 2026" vs
   "DreamHack Winter Major 2026", double spaces) — compare on a normalised key. */
const evKey = name => String(name || '').toLowerCase().replace(/\bmajor\b/g, ' ').replace(/[^a-z0-9]+/g, ' ').trim();

/* Maps → matches. One match = same two teams, same day, same event and stage.
   The score is simply maps won, so BO1, BO3 and BO5 all come out right. */
function buildSeries(maps) {
  const bag = new Map();
  for (const m of maps) {
    const pair = [m.team1.toLowerCase(), m.team2.toLowerCase()].sort().join('|');
    const key = [m.dateStr, pair, evKey(m.event), m.details.toLowerCase()].join('~');
    let s = bag.get(key);
    if (!s) { s = { dateStr: m.dateStr, dateObj: m.dateObj, event: m.event, details: m.details, teamA: m.team1, teamB: m.team2, winsA: 0, winsB: 0, maps: [], idx: m.idx }; bag.set(key, s); }
    const aIsT1 = sameTeam(m.team1, s.teamA);
    const a = aIsT1 ? m.score1 : m.score2, b = aIsT1 ? m.score2 : m.score1;
    if (a > b) s.winsA++; else if (b > a) s.winsB++;
    s.maps.push({ map: m.map, a, b, idx: m.idx });
    s.idx = Math.max(s.idx, m.idx);
  }
  const list = [...bag.values()];
  list.forEach(s => s.maps.sort((x, y) => x.idx - y.idx));
  // newest first; inside one day the later rows of the sheet are the later matches
  list.sort((x, y) => (y.dateStr || '').localeCompare(x.dateStr || '') || y.idx - x.idx);
  return list;
}

async function loadEvents() {
  const res = await fetch(URLS.events);   // large file: let the browser cache it
  if (!res.ok) throw new Error('kz-events-backup.json → HTTP ' + res.status);
  const arr = await res.json();
  state.events = (Array.isArray(arr) ? arr : []).filter(e => e && e.name).map(e => ({
    key: evKey(e.name), name: String(e.name).replace(/\s+/g, ' ').trim(), logo: e.logo || '', stars: Number(e.stars) || 0,
    major: String(e.major || '').toLowerCase() === 'yes', city: e.city || '', country: e.country || '', prizePool: e.prizePool || '',
    standings: (e.standings || []).map(s => ({ label: String(s.label || '').trim(),
      names: (s.teams || []).map(t => String(t.name || '').trim()).filter(Boolean),
      teams: (s.teams || []).map(t => String(t.name || '').trim().toLowerCase()) })),
    // lineups the event recorded (playoff teams only); empty placeholder slots are dropped
    lineups: Object.fromEntries(Object.entries(e.playoffPlayers || {}).map(([team, list]) =>
      [String(team).trim().toLowerCase(), (list || []).map(n => String(n || '').trim()).filter(Boolean)]).filter(x => x[1].length))
  }));
}

async function loadLogos() {
  const res = await fetch(URLS.logos, { cache: 'no-store' });
  if (!res.ok) return;
  const arr = await res.json();
  if (Array.isArray(arr)) arr.forEach(e => { if (e.name && e.logo) state.logos[e.name.trim().toLowerCase()] = e.logo.trim(); });
}

async function loadAwards() {
  const res = await fetch(URLS.awards);   // large file: let the browser cache it
  if (!res.ok) throw new Error('awards.json → HTTP ' + res.status);
  const data = await res.json();
  state.awards = data;
  state.historyIcons = data.__historyIcons || {};
}

/* ── small HTML helpers ── */
function logoImg(team, cls) {
  const url = state.logos[String(team).trim().toLowerCase()];
  return url ? `<img class="team-logo ${cls || 'team-logo-sm'}" src="${esc(url)}" alt="" onerror="this.style.display='none'">` : '';
}
function flagImg(country) {
  const code = COUNTRY_CODES[String(country || '').toLowerCase().trim()];
  return code ? `<img class="tp-flag" src="https://flagcdn.com/w40/${code}.png" alt="${esc(country)}" title="${esc(country)}" onerror="this.style.display='none'">` : '';
}
/* Looks for assets/players/<Name>.png|jpg|jpeg|webp; falls back to the initial letter */
function photoHtml(name) {
  const base = 'assets/players/' + encodeURIComponent(String(name).trim());
  const initial = esc(String(name).trim().charAt(0).toUpperCase() || '?');
  return `<img src="${base}.${PHOTO_EXTS[0]}" data-base="${esc(base)}" data-idx="0" data-initial="${initial}" alt="" onerror="photoFallback(this)">`;
}
function photoFallback(img) {
  const next = (parseInt(img.dataset.idx, 10) || 0) + 1;
  if (next < PHOTO_EXTS.length) { img.dataset.idx = next; img.src = img.dataset.base + '.' + PHOTO_EXTS[next]; }
  else img.outerHTML = `<span class="tp-initial">${img.dataset.initial}</span>`;
}
const roleChips = roles => roles.map(r => `<span class="tp-role">${esc(ROLE_LABELS[r] || r)}</span>`).join('');

/* ── achievements (same matching rules as the Finances page) ── */
function computeAwardKey(kind, tier, scope, compRegion, position) {
  tier = Number(tier) || 1;
  let comp;
  if (tier >= 5) comp = 'major';
  else if (scope === 'Global') comp = 'global';
  else if (scope === 'Regional') {
    comp = { Europe: 'regional_eu', America: 'regional_am', Asia: 'regional_as' }[compRegion];
    if (!comp) return null;
  } else return null;
  if (kind === 'EVENT') {
    const sub = { Winner: 'winner', Second: 'second', Third: 'third', Fourth: 'fourth' }[position];
    return sub ? `${comp}_${sub}` : null;
  }
  if (kind === 'MVP' || kind === 'EVP' || kind === 'VP') return `${comp}_${kind.toLowerCase()}`;
  return null;
}
const positionLabel = p => ({ Winner: '1st place', Second: '2nd place', Third: '3rd place', Fourth: '4th–8th' }[p] || p || '');
const starsHtml = tier => {
  tier = Number(tier) || 1;
  return tier >= 5 ? '<span class="ach-stars">★★★★★</span> Major' : `<span class="ach-stars">${'★'.repeat(Math.max(1, Math.min(4, tier)))}</span>`;
};

/* Turns one TeamHistory / PlayerHistory row into a display model */
function describeEntry(e, isPlayerRow) {
  const isBadge = e.kind === 'BADGE';
  const awardKey = e.awardKey || (isBadge ? null : computeAwardKey(e.kind, e.tier, e.scope, e.compRegion, e.position));
  const award = awardKey && state.awards ? state.awards[awardKey] : null;
  const isMedal = e.kind === 'MVP' || e.kind === 'EVP' || e.kind === 'VP';
  const fallbackTier = isMedal ? { MVP: 'gold', EVP: 'silver', VP: 'bronze' }[e.kind]
    : ({ Winner: 'gold', Second: 'silver', Third: 'bronze' }[e.position] || 'none');
  const tier = (award && award.tierColor) || fallbackTier;
  // Icon priority: icon baked into this specific entry → award registry icon → plain glyph
  const custom = isPlayerRow
    ? (e.refId != null ? state.historyIcons[e.refId] : state.historyIcons[e.id])
    : state.historyIcons[e.id];
  const icon = custom || (award && award.icon) || '';
  const label = isBadge ? (award ? String(award.label).split('—').pop().trim() : 'Badge')
    : (isMedal ? e.kind : positionLabel(e.position));
  const glyph = isBadge ? '★' : isMedal ? e.kind : ({ Winner: '1st', Second: '2nd', Third: '3rd', Fourth: 'Top 8' }[e.position] || '•');
  return { tier, icon, label, glyph, isBadge, isMedal, award };
}

function achItemHtml(e, opts) {
  opts = opts || {};
  const d = describeEntry(e, !!opts.playerRow);
  const iconHtml = d.icon
    ? `<img src="${d.icon}" alt="">`
    : `<span class="ach-glyph t-${d.tier}">${esc(d.glyph)}</span>`;
  const meta = [];
  if (d.isBadge) { meta.push(esc(d.award ? d.award.category : 'Badge')); if (e.note) meta.push(esc(e.note)); }
  else meta.push(starsHtml(e.tier), esc(e.scope + (e.compRegion ? ' ' + e.compRegion : '')));
  if (opts.showTeam && e.team && e.team.toLowerCase() !== 'none') meta.push('with ' + esc(e.team));
  if (opts.showPrize && e.wonAmount) meta.push('Prize ' + fmtMoney(e.wonAmount));
  meta.push('Month ' + e.month);
  const who = opts.showPlayer ? `<span class="ach-who">${esc(e.player)}</span> · ` : '';
  return `<div class="ach-item tier-${d.tier}">
    <div class="ach-icon">${iconHtml}</div>
    <div class="ach-text">
      <div class="ach-line1">${who}${esc(e.event)}</div>
      <div class="ach-line2"><span class="ach-pos t-${d.tier}">${esc(d.label)}</span> · ${meta.join(' · ')}</div>
    </div>
  </div>`;
}
const byNewest = (a, b) => b.month - a.month;   // stable sort keeps the file's newest-first order inside a month

/* ── view preference (icons ↔ detailed), remembered between visits ── */
const VIEW_KEYS = { team: 'nexus_team_ach_view', player: 'nexus_player_ach_view' };
function getView(which) { try { return localStorage.getItem(VIEW_KEYS[which]) === 'detailed' ? 'detailed' : 'icons'; } catch (e) { return 'icons'; } }
function setView(which, v) { try { localStorage.setItem(VIEW_KEYS[which], v); } catch (e) {} }
const toggleLabel = v => v === 'detailed' ? 'Icon view ▴' : 'Expand details ▾';

/* ── hover details: one floating tip, content looked up by index ── */
let tips = [];
function tipAttr(html) { tips.push(html); return `data-tip="${tips.length - 1}"`; }
const isPodium = e => e.kind === 'EVENT' && (e.position === 'Winner' || e.position === 'Second' || e.position === 'Third');
const scopeText = e => e.scope + (e.compRegion ? ' ' + e.compRegion : '');
const starsText = tier => { tier = Number(tier) || 1; return tier >= 5 ? '★★★★★ Major' : '★'.repeat(Math.max(1, Math.min(4, tier))); };

function iconHtml(d) {
  return d.icon ? `<img src="${d.icon}" alt="">` : `<span class="ach-glyph t-${d.tier}">${esc(d.glyph)}</span>`;
}
/* One trophy = one icon. Hover: event, placement, stars, scope, team, month. */
function trophyIconHtml(e, opts) {
  opts = opts || {};
  const d = describeEntry(e, !!opts.playerRow);
  const line2 = [opts.playerRow && e.team && e.team.toLowerCase() !== 'none' ? 'with ' + esc(e.team) : '',
    opts.showPrize && e.wonAmount ? 'Prize ' + fmtMoney(e.wonAmount) : '', 'Month ' + e.month].filter(Boolean).join(' · ');
  const tip = `<b>${esc(e.event)}</b><span class="tip-meta"><em>${esc(d.label)}</em> · ${starsText(e.tier)} · ${esc(scopeText(e))}<br>${line2}</span>`;
  return `<div class="shelf-i" ${opts.focusable ? 'tabindex="0"' : ''} ${tipAttr(tip)} aria-label="${esc(e.event + ' — ' + d.label)}">
    <div class="shelf-pic">${iconHtml(d)}</div></div>`;
}
/* Medals (MVP / EVP / VP / badges) that share the same icon collapse into one,
   with the count underneath. Hover lists every event it was won at. */
function groupMedals(rows) {
  const groups = new Map();
  for (const h of rows) {
    const d = describeEntry(h, true);
    const key = d.icon ? 'i:' + d.icon : 'k:' + (h.awardKey || computeAwardKey(h.kind, h.tier, h.scope, h.compRegion, h.position) || h.kind + '|' + scopeText(h));
    if (!groups.has(key)) groups.set(key, { d, first: h, rows: [] });
    groups.get(key).rows.push(h);
  }
  return [...groups.values()];
}
function medalTitle(g) {
  if (g.d.isBadge) return g.d.award ? String(g.d.award.label).replace(/\s*—\s*/, ' · ') : 'Badge';
  const h = g.first;
  return `${h.kind} · ${h.tier >= 5 ? 'Major' : scopeText(h)}`;
}
function medalIconHtml(g, opts) {
  opts = opts || {};
  const n = g.rows.length;
  const tip = `<b>${esc(medalTitle(g))}${n > 1 ? ' ×' + n : ''}</b><ul>${g.rows.map(h => `<li>${esc(h.event)}</li>`).join('')}</ul>`;
  return `<div class="shelf-i" ${opts.focusable ? 'tabindex="0"' : ''} ${tipAttr(tip)} aria-label="${esc(medalTitle(g) + (n > 1 ? ', ' + n + ' times' : ''))}">
    <div class="shelf-pic">${iconHtml(g.d)}</div><div class="shelf-n">${n > 1 ? n : ''}</div></div>`;
}

(() => {
  const tip = $('#tip');
  let cur = null;
  function show(el) {
    cur = el;
    tip.innerHTML = tips[Number(el.dataset.tip)] || '';
    tip.classList.add('on');
    const r = el.getBoundingClientRect(), tw = tip.offsetWidth, th = tip.offsetHeight;
    let left = Math.max(8, Math.min(r.left + r.width / 2 - tw / 2, window.innerWidth - tw - 8));
    let top = r.bottom + 8;
    if (top + th > window.innerHeight - 8) top = Math.max(8, r.top - th - 8);
    tip.style.left = left + 'px'; tip.style.top = top + 'px';
  }
  function hide() { cur = null; tip.classList.remove('on'); }
  document.addEventListener('mouseover', e => { const el = e.target.closest('[data-tip]'); if (el) { if (el !== cur) show(el); } else if (cur) hide(); });
  document.addEventListener('focusin', e => { const el = e.target.closest('[data-tip]'); if (el) show(el); });
  document.addEventListener('focusout', hide);
  window.addEventListener('scroll', hide, true);
  window.hideTip = hide;
})();

/* ── ranking movement (current ranking vs old.xlsx, like the Global Ranking) ── */
function movement() {
  const t = state.team;
  const o = state.oldRanking && state.oldRanking.find(x => sameTeam(x.team, t.team));
  if (!state.oldRanking) return null;
  if (!o) return { isNew: true };
  return { old: o, pos: o.pos - t.pos, regionPos: o.regionPos - t.regionPos, points: Math.trunc(t.points) - Math.trunc(o.points) };
}
/* positive = improved. `what` is appended for screen readers / clarity */
function deltaChip(d, opts) {
  opts = opts || {};
  if (d === 'new') return '<span class="dl new">NEW</span>';
  if (d == null) return '';
  if (d === 0) return `<span class="dl flat">—${opts.flat === undefined ? ' no change' : (opts.flat ? ' ' + opts.flat : '')}</span>`;
  const n = Math.abs(d).toLocaleString('en-US');
  return d > 0 ? `<span class="dl up">▲ ${opts.signed ? '+' : ''}${n}${opts.unit || ''}</span>` : `<span class="dl down">▼ ${opts.signed ? '−' : ''}${n}${opts.unit || ''}</span>`;
}

/* ── matches as series, from this team's point of view ── */
function teamSeries(t) {
  if (!state.series) return [];
  t = t || state.team.team;
  return state.series.filter(s => sameTeam(s.teamA, t) || sameTeam(s.teamB, t)).map(s => {
    const isA = sameTeam(s.teamA, t);
    const mine = isA ? s.winsA : s.winsB, theirs = isA ? s.winsB : s.winsA;
    return { s, isA, opp: isA ? s.teamB : s.teamA, mine, theirs, res: mine > theirs ? 'w' : mine < theirs ? 'l' : 'd' };
  });
}
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function fmtDate(iso, withYear) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || '');
  return m ? `${m[3]} ${MONTHS[Number(m[2]) - 1]}${withYear ? ' ' + m[1] : ''}` : (iso || '—');
}

/* ── render: hero ── */
function renderHero() {
  const t = state.team;
  document.title = `${t.team} · NEXUS`;
  const logoUrl = state.logos[t.team.toLowerCase()];
  const initial = esc(t.team.charAt(0).toUpperCase());
  const logo = logoUrl
    ? `<img src="${esc(logoUrl)}" alt="" onerror="this.outerHTML='<span>${initial}</span>'">`
    : `<span>${initial}</span>`;
  const region = REGION_LABELS[t.region] || t.region || '';
  const played = t.victories + t.loses;
  const wr = played ? Math.round(t.victories / played * 100) + '%' : '—';
  const th = state.teamHistory.filter(e => sameTeam(e.team, t.team) && e.kind === 'EVENT');
  const titles = th.filter(e => e.position === 'Winner').length;
  const podiums = th.filter(isPodium).length;
  const majors = Math.max(t.majors, th.filter(e => e.position === 'Winner' && e.tier >= 5).length);
  const form = teamSeries().slice(0, 5);
  const mv = movement();
  const dash = v => state.historyLoaded ? v : '—';
  const stat = (l, v, cls, extra) => `<div><div class="hero-stat-l">${l}</div><div class="hero-stat-v${cls || ''}">${v}</div>${extra ? `<div class="hero-stat-d">${extra}</div>` : ''}</div>`;
  const gChip = mv ? (mv.isNew ? deltaChip('new') : deltaChip(mv.pos, { flat: 'holding' })) : '';
  const rChip = mv && !mv.isNew ? deltaChip(mv.regionPos, { flat: '' }) : '';
  const pChip = mv && !mv.isNew ? deltaChip(mv.points, { signed: true, flat: 'no change' }) : '';

  $('#tpHero').innerHTML = `
    <div class="hero-ghost" aria-hidden="true">${t.pos}</div>
    <div class="hero-main">
      <div class="hero-logo">${logo}</div>
      <div class="hero-id">
        <div class="hero-kicker"><span>Team profile</span>
          ${region ? `<span class="hero-chip">${esc(region)}</span>` : ''}${t.tier ? `<span class="hero-chip accent">${esc(t.tier)}</span>` : ''}</div>
        <h1 class="hero-name">${esc(t.team)}</h1>
        ${form.length ? `<div class="hero-form"><span>Last ${form.length} matches</span><div class="hero-form-dots">${form.map(f =>
          `<span class="${f.res}" title="${esc(f.mine + '–' + f.theirs + ' vs ' + f.opp)}">${f.res.toUpperCase()}</span>`).join('')}</div></div>` : ''}
      </div>
      <div class="hero-rank">
        <div class="hero-rank-label">World ranking</div>
        <div class="hero-rank-row"><div class="hero-rank-n${t.pos <= 3 ? ' r' + t.pos : ''}"><sup>#</sup>${t.pos}</div>${gChip ? `<div class="hero-rank-d">${gChip}</div>` : ''}</div>
        <div class="hero-rank-sub"><b>#${t.regionPos}</b> in ${esc(region || 'region')} ${rChip}</div>
      </div>
    </div>
    <div class="hero-stats">
      ${stat('Points', fmtPts(t.points), '', pChip)}
      ${stat('Record', `${t.victories}–${t.loses}`)}
      ${stat('Win rate', wr)}
      ${stat('Win streak', t.streak, t.streak >= 3 ? ' hot' : '')}
      ${stat('Titles', dash(titles), titles ? ' hot' : '')}
      ${stat('Podiums', dash(podiums))}
      ${stat('Majors', majors, majors ? ' hot' : '')}
    </div>`;

}

/* ── render: trophy cabinet (icons ↔ detailed) — team results only ── */
function tallyHtml(items) {
  return '<div class="ach-tally">' + items.map(i =>
    `<div><div class="ach-tally-n ${i[1] === 0 ? 't-none' : (i[2] || '')}">${i[1]}</div><div class="ach-tally-l">${i[0]}</div></div>`).join('') + '</div>';
}
function renderAchievements() {
  const box = $('#tpAch'), btn = $('#tpAchToggle');
  const view = getView('team');
  btn.textContent = toggleLabel(view);
  btn.setAttribute('aria-pressed', view === 'detailed');
  if (!state.historyLoaded) { box.innerHTML = '<div class="shelf-none">Achievements file (file/history_updated.xlsx) could not be loaded.</div>'; return; }
  const t = state.team.team;
  const events = state.teamHistory.filter(e => sameTeam(e.team, t)).sort(byNewest);
  const results = events.filter(e => e.kind === 'EVENT');
  const podium = events.filter(e => isPodium(e) || e.kind === 'BADGE');
  const others = results.length - results.filter(isPodium).length;
  $('#tpCabCount').textContent = podium.length ? `${podium.length} on the shelf` : '';

  if (view === 'icons') {
    box.innerHTML = (podium.length
        ? `<div class="shelf lg">${podium.map(e => trophyIconHtml(e, { showPrize: true, focusable: true })).join('')}</div>`
        : '<div class="shelf-none">No top-3 finishes yet.</div>')
      + (others ? `<div class="cab-foot">${others} top 4–8 finish${others > 1 ? 'es' : ''} in the expanded view</div>` : '');
    return;
  }
  if (!events.length) { box.innerHTML = '<div class="shelf-none">No achievements recorded for this team yet.</div>'; return; }
  const count = pos => results.filter(e => e.position === pos).length;
  box.innerHTML = tallyHtml([
    ['Titles', count('Winner'), 't-gold'], ['Runner-up', count('Second'), 't-silver'], ['3rd place', count('Third'), 't-bronze'],
    ['Top 4–8', count('Fourth'), 't-none'], ['Events placed', results.length], ['Prize money', fmtMoney(results.reduce((s, e) => s + e.wonAmount, 0))]
  ]) + '<div class="ach-grid">' + events.map(e => achItemHtml(e, { showPrize: true })).join('') + '</div>';
}

/* ── render: ranking trend (Overview tab) ── */
function trendSvg(pts) {
  const W = 760, H = 240, L = 70, R = 34, T = 34, B = 44;
  const all = pts.flatMap(p => [p.pos, p.regionPos]);
  let lo = Math.min(...all), hi = Math.max(...all);
  if (hi - lo < 4) { hi = lo + 4; }
  lo = Math.max(1, lo - 1); hi = hi + 1;
  const x = i => pts.length === 1 ? (L + W - R) / 2 : L + i * (W - L - R) / (pts.length - 1);
  const y = v => T + (v - lo) * (H - T - B) / (hi - lo);          // position 1 at the top
  const ticks = [...new Set([lo, Math.round((lo + hi) / 2), hi])];
  const line = key => pts.map((p, i) => `${x(i).toFixed(1)},${y(p[key]).toFixed(1)}`).join(' ');
  const every = Math.ceil(pts.length / 10);
  const showLbl = i => i === 0 || i === pts.length - 1 || i % every === 0;
  const dots = (key, cls) => pts.map((p, i) => {
    // the line that is higher at this point gets its label above, the other one below, so they never collide
    const other = key === 'pos' ? p.regionPos : p.pos;
    const up = p[key] < other || (p[key] === other && key === 'pos');
    const dy = up ? -11 : 19;
    return `<circle class="${cls}" cx="${x(i).toFixed(1)}" cy="${y(p[key]).toFixed(1)}" r="4.5"><title>${esc(p.label)}: #${p[key]}</title></circle>` +
      (showLbl(i) ? `<text class="tr-val ${cls}" x="${x(i).toFixed(1)}" y="${(y(p[key]) + dy).toFixed(1)}" text-anchor="middle">#${p[key]}</text>` : '');
  }).join('');
  return `<svg class="tr-svg" viewBox="0 0 ${W} ${H}" role="img" aria-label="Ranking position over time, global and regional">
    ${ticks.map(v => `<line class="tr-grid" x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}"/><text class="tr-axis" x="${L - 10}" y="${y(v) + 4}" text-anchor="end">#${v}</text>`).join('')}
    <polyline class="tr-line reg" points="${line('regionPos')}"/><polyline class="tr-line glo" points="${line('pos')}"/>
    ${dots('regionPos', 'reg')}${dots('pos', 'glo')}
    ${pts.map((p, i) => showLbl(i) ? `<text class="tr-axis" x="${x(i).toFixed(1)}" y="${H - 12}" text-anchor="middle">${esc(p.label)}</text>` : '').join('')}
  </svg>`;
}
function renderTrend() {
  const box = $('#tpTrend');
  const t = state.team, mv = movement();
  const region = REGION_LABELS[t.region] || t.region || 'Region';
  const tile = (label, now, before, chip) => `<div class="mv-tile"><div class="hero-stat-l">${label}</div>
    <div class="mv-now">${now}</div><div class="mv-was">${before != null ? 'was ' + before : ''} ${chip || ''}</div></div>`;
  let tiles;
  if (!mv) tiles = tile('Global rank', '#' + t.pos) + tile(esc(region) + ' rank', '#' + t.regionPos) + tile('Points', fmtPts(t.points));
  else if (mv.isNew) tiles = tile('Global rank', '#' + t.pos, null, deltaChip('new')) + tile(esc(region) + ' rank', '#' + t.regionPos, null, deltaChip('new')) + tile('Points', fmtPts(t.points), null, deltaChip('new'));
  else tiles = tile('Global rank', '#' + t.pos, '#' + mv.old.pos, deltaChip(mv.pos, { unit: mv.pos && Math.abs(mv.pos) === 1 ? ' place' : ' places' }))
    + tile(esc(region) + ' rank', '#' + t.regionPos, '#' + mv.old.regionPos, deltaChip(mv.regionPos, { unit: Math.abs(mv.regionPos) === 1 ? ' place' : ' places' }))
    + tile('Points', fmtPts(t.points), fmtPts(mv.old.points), deltaChip(mv.points, { signed: true }));
  const pts = state.trajectory;
  $('#tpTrendNote').textContent = pts.length > 1 ? `${pts.length} ranking snapshots` : '';
  box.innerHTML = `<div class="mv-grid"><div class="mv-tiles">${tiles}</div>
    <div class="mv-chart">${pts.length > 1
      ? `<div class="tr-legend"><span class="glo">● Global</span><span class="reg">● ${esc(region)}</span><span>higher on the chart = better position</span></div>${trendSvg(pts)}`
      : '<div class="tp-empty">The chart needs at least two ranking snapshots (file/old.xlsx or file/archives/).</div>'}</div></div>`;
}

/* ── render: roster ── */
const CARD_TROPHY_LIMIT = 4;   // newest trophies stay on the card; older ones drop off as new ones arrive
function teamPlayers() {
  return state.players.filter(p => sameTeam(p.team, state.team.team))
    .sort((a, b) => (a.roster === 'Starter' ? 0 : 1) - (b.roster === 'Starter' ? 0 : 1) || b.marketValue - a.marketValue);
}
function playerAwards(name) {
  const hist = state.playerHistory.filter(h => h.player === name).sort(byNewest);
  return { hist, trophies: hist.filter(isPodium), placements: hist.filter(h => h.kind === 'EVENT'), medals: hist.filter(h => h.kind !== 'EVENT') };
}
function renderRoster() {
  const box = $('#tpRoster');
  if (!state.playersLoaded) { box.innerHTML = '<div class="tp-empty">Roster file (file/players_updated.xlsx) could not be loaded.</div>'; return; }
  const list = teamPlayers();
  if (!list.length) { box.innerHTML = '<div class="tp-empty">No players registered for this team.</div>'; return; }
  $('#tpRosterNote').textContent = 'Squad value ' + fmtMoney(list.reduce((s, p) => s + p.marketValue, 0));
  box.innerHTML = '<div class="tp-roster">' + list.map(p => {
    const a = playerAwards(p.name);
    const shown = a.trophies.slice(0, CARD_TROPHY_LIMIT), hidden = a.trophies.length - shown.length;
    const groups = groupMedals(a.medals);
    const medals = groups.length ? `<div class="shelf xs tp-cardmedals">${groups.map(g => medalIconHtml(g)).join('')}</div>` : '';
    const trophies = shown.length
      ? `<div class="tp-awards-l"><span>Latest trophies</span>${hidden > 0 ? `<span>+${hidden} older</span>` : ''}</div><div class="shelf cardrow">${shown.map(h => trophyIconHtml(h, { playerRow: true })).join('')}</div>`
      : `<div class="shelf-none">${state.historyLoaded ? 'No trophies yet' : '—'}</div>`;
    return `<button type="button" class="tp-player" data-player="${esc(p.name)}">
      <div class="tp-photo">${photoHtml(p.name)}${p.roster && p.roster !== 'Starter' ? `<span class="tp-bench-tag">${esc(p.roster)}</span>` : ''}</div>
      <div class="tp-pbody">
        <div class="tp-phead"><div class="tp-pname">${flagImg(p.country)}<span>${esc(p.name)}</span></div>${medals}</div>
        <div class="tp-roles">${roleChips(p.roles)}</div>
        <div class="tp-awards">${trophies}</div>
      </div>
    </button>`;
  }).join('') + '</div>';
}

/* ── History tab: tournaments + matches ── */
const ordinal = n => { n = Number(n); const s = ['th', 'st', 'nd', 'rd'], v = n % 100; return n + (s[(v - 20) % 10] || s[v] || s[0]); };
function prettyPlace(label) {
  const nums = String(label).match(/\d+/g);
  if (!nums) return { text: label, tier: 'none' };
  const text = nums.length > 1 ? `${ordinal(nums[0])}–${ordinal(nums[1])}` : ordinal(nums[0]);
  return { text, tier: nums.length === 1 ? ({ 1: 'gold', 2: 'silver', 3: 'bronze' }[nums[0]] || 'none') : 'none' };
}
function buildTournaments(t) {
  t = t || state.team.team;
  const tl = t.toLowerCase();
  const bag = new Map();
  const get = (key, name) => { if (!bag.has(key)) bag.set(key, { key, name, series: [], th: null, ev: null }); return bag.get(key); };
  teamSeries(t).forEach(v => get(evKey(v.s.event), v.s.event).series.push(v));
  state.teamHistory.filter(e => sameTeam(e.team, t) && e.kind === 'EVENT').forEach(e => { get(evKey(e.event), e.event).th = e; });
  (state.events || []).forEach(ev => {
    const st = ev.standings.find(s => s.teams.includes(tl));
    if (st || bag.has(ev.key)) { const o = get(ev.key, ev.name); o.ev = ev; o.standing = st ? st.label : null; }
  });
  const list = [...bag.values()];
  list.forEach(o => {
    if (o.ev) o.name = o.ev.name;
    o.series.sort((a, b) => (a.s.dateStr || '').localeCompare(b.s.dateStr || '') || a.s.idx - b.s.idx);   // path: oldest → newest
    o.first = o.series.length ? o.series[0].s.dateStr : ''; o.last = o.series.length ? o.series[o.series.length - 1].s.dateStr : '';
    o.won = o.series.filter(v => v.res === 'w').length; o.lost = o.series.filter(v => v.res === 'l').length;
    if (o.standing) o.place = prettyPlace(o.standing);
    else if (o.th) o.place = prettyPlace({ Winner: '1', Second: '2', Third: '3', Fourth: '4-8' }[o.th.position] || o.th.position);
    else if (o.series.length) {
      const lastS = o.series[o.series.length - 1];
      o.place = { text: lastS.res === 'l' ? 'Out · ' + (lastS.s.details || 'eliminated') : 'Played', tier: 'none', soft: true };
    } else o.place = { text: '—', tier: 'none', soft: true };
    o.stars = (o.ev && o.ev.stars) || (o.th && o.th.tier) || 0;
  });
  list.sort((a, b) => (b.last || '').localeCompare(a.last || '') || ((b.th ? b.th.month : 0) - (a.th ? a.th.month : 0)));
  return list;
}
function seriesRowHtml(v, opts) {
  opts = opts || {};
  const s = v.s;
  const opp = teamLink(v.opp, true);
  const maps = s.maps.map((m, i) => {
    const mine = v.isA ? m.a : m.b, theirs = v.isA ? m.b : m.a;
    return `<div class="sr-map"><span class="sr-mapn">Map ${i + 1}</span><span class="pa-map-badge">${esc(m.map || '—')}</span>
      <span class="sr-mapscore ${mine > theirs ? 'w' : mine < theirs ? 'l' : ''}">${mine}–${theirs}</span></div>`;
  }).join('');
  return `<details class="sr">
    <summary>
      <span class="sr-date">${esc(fmtDate(s.dateStr, opts.year))}</span>
      <span class="tp-res ${v.res}">${v.res.toUpperCase()}</span>
      <span class="sr-score">${v.mine}–${v.theirs}</span>
      <span class="sr-opp"><span class="sr-vs">vs</span>${opp}</span>
      <span class="sr-stage">${esc(s.details || '—')}${opts.showEvent && s.event ? `<small>${esc(s.event)}</small>` : ''}</span>
      <span class="sr-chev" aria-hidden="true">▾</span>
    </summary>
    <div class="sr-maps">${maps}</div>
  </details>`;
}
const MATCH_PAGE = 20;
function renderMatches() {
  const box = $('#tpMatches');
  if (state.series === null) { box.innerHTML = '<div class="tp-empty">Match history (file/history.xlsx) is not available.</div>'; return; }
  const all = teamSeries();
  if (!all.length) { box.innerHTML = '<div class="tp-empty">No matches recorded for this team.</div>'; return; }
  const w = all.filter(v => v.res === 'w').length, l = all.filter(v => v.res === 'l').length;
  const mapsW = all.reduce((s, v) => s + v.mine, 0), mapsL = all.reduce((s, v) => s + v.theirs, 0);
  $('#tpMatchNote').textContent = `${all.length} matches · ${w}–${l} · maps ${mapsW}–${mapsL}`;
  const shown = all.slice(0, state.matchLimit);
  box.innerHTML = shown.map(v => seriesRowHtml(v, { showEvent: true, year: true })).join('')
    + (all.length > shown.length ? `<button type="button" class="btn btn-ghost" id="tpMoreMatches" style="margin-top:14px">Show ${Math.min(MATCH_PAGE, all.length - shown.length)} more · ${all.length - shown.length} left</button>` : '');
}

/* ── tabs ── */
function setTab(name, push) {
  name = name === 'history' ? 'history' : 'overview';
  ['overview', 'history'].forEach(n => {
    $('#tab-' + n).hidden = n !== name;
    const b = $(`.tp-tab[data-tab="${n}"]`); b.classList.toggle('active', n === name); b.setAttribute('aria-selected', n === name);
  });
  document.querySelectorAll('.mh-ribbon [data-tab]').forEach(a => a.classList.toggle('active', a.dataset.tab === name));
  if (push) history.replaceState(null, '', name === 'history' ? '#history' : location.pathname + location.search);
}

/* ── player panel ── */
let lastFocus = null;
function openPlayer(name) {
  const p = state.players.find(x => x.name === name);
  if (!p) return;
  state.openPlayer = name;
  lastFocus = document.activeElement;
  hideTip();
  renderPlayerPanel(true);
  $('#ppOverlay').classList.add('open');
  $('#ppPanel').classList.add('open');
  document.body.style.overflow = 'hidden';
  $('#ppPanel').focus();
}
function closePlayer() {
  state.openPlayer = null;
  hideTip();
  $('#ppOverlay').classList.remove('open');
  $('#ppPanel').classList.remove('open');
  document.body.style.overflow = '';
  if (lastFocus && lastFocus.focus) lastFocus.focus();
}
const titleCase = s => String(s).toLowerCase().replace(/\b[a-z]/g, c => c.toUpperCase()).replace(/\bFa\b/, 'Free-agent');
function careerHtml(p, noTitle) {
  const seen = new Set();
  const moves = state.transactions.filter(x => x.player.toLowerCase() === p.name.toLowerCase()).sort((a, b) => a.id - b.id)
    .filter(x => { const k = [x.month, x.type, x.from, x.to].join('|'); if (seen.has(k)) return false; seen.add(k); return true; });
  const hasTeam = p.team && p.team.toLowerCase() !== 'none';
  const teamCell = n => /^free agen/i.test(n) ? '<span class="cr-fa">Free agency</span>' : `<span class="cr-team">${teamLink(n, true)}</span>`;
  let rows = moves.map(x => `<li><span class="cr-month">Month ${x.month}${fmtLogged(x.id) ? `<small>${fmtLogged(x.id)}</small>` : ''}</span><span class="cr-type">${esc(titleCase(x.type))}</span>
      <span class="cr-move">${x.from === x.to ? teamCell(x.to) : teamCell(x.from) + '<span class="cr-arrow">→</span>' + teamCell(x.to)}</span>
      ${x.amount > 0 ? `<span class="cr-fee">Fee ${fmtMoney(x.amount)}</span>` : ''}</li>`).join('');
  rows += `<li class="cr-now"><span class="cr-month">Now</span><span class="cr-type">${hasTeam ? esc(p.roster || 'Player') : 'Unsigned'}</span><span class="cr-move">${hasTeam ? teamCell(p.team) : teamCell('Free agency')}</span></li>`;
  // teams the player collected results with, even without a logged transfer
  const others = [...new Set(state.playerHistory.filter(h => h.player === p.name && h.team && h.team.toLowerCase() !== 'none' && !sameTeam(h.team, p.team)).map(h => h.team))];
  return `${noTitle ? '' : '<div class="ach-sub">Career</div>'}<ul class="cr-list">${rows}</ul>`
    + (others.length ? `<div class="cab-foot">Also has results with ${others.map(esc).join(', ')}</div>` : '')
    + (!moves.length ? '<div class="cab-foot">No transfers on record · dates are the day each move was logged</div>' : '<div class="cab-foot">Dates are the day each move was logged</div>');
}
function renderPlayerPanel(resetScroll) {
  const p = state.players.find(x => x.name === state.openPlayer);
  if (!p) return;
  const hasTeam = p.team && p.team.toLowerCase() !== 'none';
  const a = playerAwards(p.name);
  const majors = Math.max(p.majorWinner, a.placements.filter(h => h.tier >= 5 && h.position === 'Winner').length);
  const view = getView('player');
  const groups = groupMedals(a.medals);
  const meta = [
    hasTeam ? `<span class="pp-team">${logoImg(p.team)}${esc(p.team)}</span>` : '<span class="pp-team">Free agent</span>',
    p.roster && hasTeam ? esc(p.roster) : '',
    p.country ? esc(p.country) : ''
  ].filter(Boolean).join('<span class="sep">|</span>');

  $('#ppHead').innerHTML = `
    <div class="pp-photo">${photoHtml(p.name)}</div>
    <div class="pp-id">
      <div class="pp-kicker">Player</div>
      <div class="pp-name" id="ppName">${flagImg(p.country)}<span>${esc(p.name)}</span></div>
      <div class="pp-meta">${meta}</div>
      <div class="tp-roles" style="margin-top:10px;min-height:0">${roleChips(p.roles) || '<span class="tp-role">No role assigned</span>'}</div>
    </div>
    ${groups.length ? `<div class="pp-medals"><div class="tp-awards-l"><span>Medals</span></div><div class="shelf pm">${groups.map(g => medalIconHtml(g, { focusable: true })).join('')}</div></div>` : ''}
    <button type="button" class="pp-close" id="ppClose" aria-label="Close player panel">×</button>`;

  let body = `
    <div class="pp-mv"><span class="tp-mv-label">Market value</span><span class="pp-mv-value">${fmtMoney(p.marketValue)}</span></div>
    <div class="pp-counts">
      ${[['Titles', p.titles, 't-gold'], ['Majors', majors, 't-gold'], ['MVP', p.mvp, 't-gold'], ['EVP', p.evp, 't-silver'], ['VP', p.vp, 't-bronze']]
        .map(c => `<div><div class="ach-tally-n ${c[1] ? c[2] : 't-none'}">${c[1]}</div><div class="ach-tally-l">${c[0]}</div></div>`).join('')}
    </div>
    <a class="btn btn-primary pp-open" href="${playerUrl(p.name)}">Open player profile →</a>
    <div class="pp-achhead"><div class="pp-achhead-t">Trophy Cabinet</div>
      ${a.hist.length ? `<button type="button" class="btn btn-ghost view-toggle" id="ppAchToggle" aria-pressed="${view === 'detailed'}">${toggleLabel(view)}</button>` : ''}</div>`;

  if (!state.historyLoaded) body += '<div class="tp-empty" style="margin-top:12px">Achievements file could not be loaded.</div>';
  else if (!a.hist.length) body += '<div class="tp-empty" style="margin-top:12px">No achievements recorded for this player yet.</div>';
  else if (view === 'icons') {
    const others = a.placements.length - a.trophies.length;
    body += `<div class="pp-shelfbox" style="margin-top:12px">${a.trophies.length
      ? `<div class="shelf md">${a.trophies.map(h => trophyIconHtml(h, { playerRow: true, focusable: true })).join('')}</div>`
      : '<div class="shelf-none">No top-3 finishes yet.</div>'}</div>
      <div class="cab-foot">Hover an icon for details${others ? ` · ${others} top 4–8 finish${others > 1 ? 'es' : ''} in the expanded view` : ''}</div>`;
  } else {
    if (a.medals.length) body += `<div class="ach-sub">Individual awards · ${a.medals.length}</div><div class="ach-grid">${a.medals.map(h => achItemHtml(h, { playerRow: true, showTeam: true })).join('')}</div>`;
    if (a.placements.length) body += `<div class="ach-sub">Trophies &amp; placements · ${a.placements.length}</div><div class="ach-grid">${a.placements.map(h => achItemHtml(h, { playerRow: true, showTeam: true })).join('')}</div>`;
  }
  if (state.historyLoaded) body += careerHtml(p);
  $('#ppBody').innerHTML = body;
  if (resetScroll) $('#ppBody').scrollTop = 0;
}

document.addEventListener('click', e => {
  const tab = e.target.closest('[data-tab]');
  if (tab) { e.preventDefault(); setTab(tab.dataset.tab, true); if (tab.closest('.mh-ribbon')) $('#tpTabs').scrollIntoView({ block: 'start' }); return; }
  const card = e.target.closest('.tp-player');
  if (card) { openPlayer(card.dataset.player); return; }
  if (e.target.closest('#tpMoreMatches')) { state.matchLimit += MATCH_PAGE; renderMatches(); return; }
  if (e.target.closest('#tpAchToggle')) { setView('team', getView('team') === 'detailed' ? 'icons' : 'detailed'); hideTip(); renderAchievements(); return; }
  if (e.target.closest('#ppAchToggle') && !state.player) { setView('player', getView('player') === 'detailed' ? 'icons' : 'detailed'); hideTip(); renderPlayerPanel(false); $('#ppAchToggle').focus(); return; }
  if (e.target.closest('#ppClose') || e.target.id === 'ppOverlay') closePlayer();
});
document.addEventListener('keydown', e => { if (e.key === 'Escape' && state.openPlayer) closePlayer(); });


/* ══════════════════════════════════════════════════════════════
   v4 — lineups per event, former players, player events,
        roster hover, search, player profile page
══════════════════════════════════════════════════════════════ */
const playerUrl = name => 'player.html?player=' + encodeURIComponent(name);
const teamUrl = name => 'team.html?team=' + encodeURIComponent(name);
const findPlayer = name => state.players.find(p => p.name.toLowerCase() === String(name).trim().toLowerCase());
function playerLink(name) {
  const p = findPlayer(name);
  return p ? `<a class="pl-link" href="${playerUrl(p.name)}">${esc(p.name)}</a>` : `<span class="pl-link off" title="Not in the players file">${esc(name)}</span>`;
}
function teamLink(name, withLogo) {
  const ranked = state.rankedSet && state.rankedSet.has(String(name).toLowerCase());
  const inner = (withLogo ? logoImg(name) : '') + esc(name);
  return ranked ? `<a class="tp-opp" data-team="${esc(name)}" href="${teamUrl(name)}">${inner}</a>` : `<span class="tp-opp-plain">${inner}</span>`;
}
function fmtLogged(ms) {
  const d = new Date(Number(ms));
  if (!ms || isNaN(d.getTime()) || d.getFullYear() < 2000 || d.getFullYear() > 2100) return '';
  return `${String(d.getDate()).padStart(2, '0')} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
}

/* ── tournaments per team, cached (the player page asks for many teams) ── */
let tourCache = new Map();
function tournamentsOf(team) {
  const k = team.toLowerCase();
  if (!tourCache.has(k)) tourCache.set(k, buildTournaments(team));
  return tourCache.get(k);
}
/* first / last match day of each event, across all teams */
function eventDates() {
  if (state._evDates) return state._evDates;
  const m = new Map();
  (state.series || []).forEach(s => {
    const k = evKey(s.event); if (!s.dateStr) return;
    const o = m.get(k) || { first: s.dateStr, last: s.dateStr, teams: new Map() };
    if (s.dateStr < o.first) o.first = s.dateStr; if (s.dateStr > o.last) o.last = s.dateStr;
    o.teams.set(s.teamA.toLowerCase(), s.teamA); o.teams.set(s.teamB.toLowerCase(), s.teamB);
    m.set(k, o);
  });
  return (state._evDates = m);
}

/* ── lineups ──
   1) If the events backup recorded the lineup for that team at that event, use it.
   2) Otherwise rebuild it: start from today's roster and undo every transfer that was
      logged after the event started. Marked as "estimated" on the page. */
function rosterAt(team, ms) {
  const names = state.players.filter(p => sameTeam(p.team, team))
    .sort((a, b) => (a.roster === 'Starter' ? 0 : 1) - (b.roster === 'Starter' ? 0 : 1) || b.marketValue - a.marketValue).map(p => p.name);
  if (ms == null) return names;
  const has = n => names.findIndex(x => x.toLowerCase() === n.toLowerCase());
  state.transactions.filter(x => x.id > ms).sort((a, b) => b.id - a.id).forEach(x => {
    const into = sameTeam(x.to, team), outOf = sameTeam(x.from, team);
    if (into && !outOf) { const i = has(x.player); if (i >= 0) names.splice(i, 1); }       // joined later → was not here yet
    else if (outOf && !into && has(x.player) < 0) names.push(x.player);                      // left later → was still here
  });
  return names;
}
function lineupFor(team, key) {
  const ev = (state.events || []).find(e => e.key === key);
  const rec = ev && ev.lineups[team.toLowerCase()];
  if (rec && rec.length) return { players: rec, source: 'recorded' };
  if (!state.playersLoaded) return { players: [], source: 'none' };
  const d = eventDates().get(key);
  const ms = d ? new Date(d.first + 'T00:00:00').getTime() : null;
  return { players: rosterAt(team, ms), source: 'estimated' };
}
function lineupHtml(team, key) {
  const lu = lineupFor(team, key);
  if (!lu.players.length) return '';
  return `<div class="lu"><span class="lu-l">Lineup</span><span class="lu-p">${lu.players.map(playerLink).join('')}</span>
    <span class="lu-src ${lu.source}">${lu.source === 'recorded' ? 'recorded at the event' : 'estimated · today’s roster with later transfers undone'}</span></div>`;
}

/* ── former players of a team: logged departures + names seen in old lineups ── */
function formerPlayers(team) {
  const now = new Set(state.players.filter(p => sameTeam(p.team, team)).map(p => p.name.toLowerCase()));
  const out = new Map();
  state.transactions.filter(x => sameTeam(x.from, team) && !sameTeam(x.to, team)).sort((a, b) => a.id - b.id).forEach(x => {
    if (now.has(x.player.toLowerCase())) return;                       // came back
    out.set(x.player.toLowerCase(), { name: x.player, tx: x, events: [] });
  });
  tournamentsOf(team).forEach(o => {
    const lu = lineupFor(team, o.key);
    if (lu.source !== 'recorded') return;
    lu.players.forEach(n => {
      const k = n.toLowerCase(); if (now.has(k)) return;
      if (!out.has(k)) out.set(k, { name: n, tx: null, events: [] });
      out.get(k).events.push(o.name);
    });
  });
  return [...out.values()].sort((a, b) => (b.tx ? b.tx.id : 0) - (a.tx ? a.tx.id : 0) || a.name.localeCompare(b.name));
}
function renderFormer() {
  const box = $('#tpFormer'); if (!box) return;
  if (!state.historyLoaded || !state.playersLoaded) { box.innerHTML = '<div class="tp-empty">Needs the players and history files.</div>'; return; }
  const list = formerPlayers(state.team.team);
  $('#tpFormerNote').textContent = list.length ? `${list.length} player${list.length > 1 ? 's' : ''}` : '';
  if (!list.length) { box.innerHTML = '<div class="tp-empty">No departures on record for this team.</div>'; return; }
  box.innerHTML = '<ul class="fm-list">' + list.map(f => {
    const p = findPlayer(f.name);
    const where = p ? (p.team && p.team.toLowerCase() !== 'none' ? teamLink(p.team, true) : '<span class="cr-fa">Free agent</span>') : '<span class="cr-fa">not in the players file</span>';
    const when = f.tx ? `<span class="fm-when"><b>${esc(titleCase(f.tx.type))}</b> · Month ${f.tx.month}${fmtLogged(f.tx.id) ? ' · logged ' + fmtLogged(f.tx.id) : ''}${f.tx.amount > 0 ? ' · fee ' + fmtMoney(f.tx.amount) : ''}</span>`
      : '<span class="fm-when"><b>Left</b> · no transfer logged</span>';
    const seen = f.events.length ? `<span class="fm-seen">Played ${f.events.map(esc).join(', ')}</span>` : '';
    return `<li><span class="fm-name">${p ? flagImg(p.country) : ''}${playerLink(f.name)}</span>${when}<span class="fm-now"><span class="sr-vs">now</span>${where}</span>${seen}</li>`;
  }).join('') + '</ul>';
}

/* ── events a player took part in (any team), with that team's result ── */
function playerEvents(name) {
  const nl = name.toLowerCase();
  const dates = eventDates();
  const keys = new Set([...(state.events || []).map(e => e.key), ...dates.keys()]);
  const found = new Map();
  keys.forEach(key => {
    const teams = new Map(dates.get(key) ? dates.get(key).teams : []);
    const ev = (state.events || []).find(e => e.key === key);
    if (ev) ev.standings.forEach(s => s.names.forEach(n => { if (!teams.has(n.toLowerCase())) teams.set(n.toLowerCase(), n); }));
    teams.forEach(team => {
      const lu = lineupFor(team, key);
      if (lu.players.some(p => p.toLowerCase() === nl)) found.set(key + '|' + team.toLowerCase(), { key, team, source: lu.source });
    });
  });
  // results stored on the player always count, whatever the lineups say
  state.playerHistory.filter(h => h.player.toLowerCase() === nl && h.kind === 'EVENT' && h.team && h.team.toLowerCase() !== 'none').forEach(h => {
    const key = evKey(h.event), id = key + '|' + h.team.toLowerCase();
    const sameEvent = [...found.keys()].filter(k => k.startsWith(key + '|'));
    sameEvent.forEach(k => { if (k !== id && found.get(k).source === 'estimated') found.delete(k); });   // the recorded result wins over a guess
    if (!found.has(id)) found.set(id, { key, team: h.team, source: 'recorded' });
  });
  const list = [...found.values()].map(f => {
    const canon = state.ranking.find(t => sameTeam(t.team, f.team));
    const team = canon ? canon.team : f.team;
    return { ...f, team, o: tournamentsOf(team).find(o => o.key === f.key) };
  }).filter(f => f.o);
  list.sort((a, b) => (b.o.last || '').localeCompare(a.o.last || ''));
  return list;
}

/* ── one tournament row (team page and player page share it) ── */
function tourRowHtml(o, team, opts) {
  opts = opts || {};
  const where = o.ev ? [o.ev.city, o.ev.country].filter(Boolean).filter((x, i, a) => a.indexOf(x) === i).join(', ') : '';
  const dates = o.first ? (o.first === o.last ? fmtDate(o.first, true) : `${fmtDate(o.first)} – ${fmtDate(o.last, true)}`) : (o.th ? 'Month ' + o.th.month : '');
  const meta = [o.stars ? `<span class="ach-stars">${'★'.repeat(Math.min(5, o.stars))}</span>${o.stars >= 5 ? ' Major' : ''}` : '', esc(where), esc(dates),
    o.ev && o.ev.prizePool ? 'Prize pool ' + esc(o.ev.prizePool) : ''].filter(Boolean).join(' · ');
  const logo = o.ev && o.ev.logo ? `<img src="${o.ev.logo}" alt="">` : `<span>${esc(o.name.charAt(0))}</span>`;
  const withTeam = opts.showTeam ? `<span class="tour-with">with ${logoImg(team)}<b>${esc(team)}</b>${opts.source === 'estimated' ? ' <i title="No lineup was recorded for this team at this event. Rebuilt from today’s roster and the logged transfers.">estimated</i>' : ''}</span>` : '';
  return `<details class="tour">
    <summary>
      <span class="tour-logo">${logo}</span>
      <span class="tour-id"><span class="tour-name">${esc(o.name)}</span>${withTeam}<span class="tour-meta">${meta}</span></span>
      <span class="tour-rec">${o.series.length ? `<b>${o.won}–${o.lost}</b><small>matches</small>` : ''}</span>
      <span class="tour-place t-${o.place.tier}${o.place.soft ? ' soft' : ''}">${esc(o.place.text)}</span>
      <span class="sr-chev" aria-hidden="true">▾</span>
    </summary>
    <div class="tour-path">${lineupHtml(team, o.key)}${o.series.length
      ? `<div class="tour-path-l">${opts.showTeam ? esc(team) + '’s path' : 'Path'} through the event · first match at the top</div>${o.series.map(v => seriesRowHtml(v)).join('')}`
      : '<div class="shelf-none">No matches logged for this event.</div>'}</div>
  </details>`;
}
function renderTournaments() {
  const box = $('#tpTours'); if (!box) return;
  if (state.series === null && !state.historyLoaded) { box.innerHTML = '<div class="tp-empty">Match history and achievements files are not available.</div>'; return; }
  const list = tournamentsOf(state.team.team);
  $('#tabHistCount').textContent = list.length || '';
  $('#tpTourNote').textContent = list.length ? `${list.length} tournament${list.length > 1 ? 's' : ''}` : '';
  box.innerHTML = list.length ? list.map(o => tourRowHtml(o, state.team.team)).join('') : '<div class="tp-empty">No tournaments on record for this team.</div>';
}

/* ── roster preview when hovering any team link on these pages ── */
(() => {
  const card = document.createElement('div');
  card.className = 'roster-hover'; card.setAttribute('role', 'tooltip');
  document.body.appendChild(card);
  let cur = null;
  function html(team) {
    const list = state.players.filter(p => sameTeam(p.team, team)).sort((a, b) => (a.roster === 'Starter' ? 0 : 1) - (b.roster === 'Starter' ? 0 : 1) || a.name.localeCompare(b.name));
    const r = state.ranking.find(t => sameTeam(t.team, team));
    const body = !state.playersLoaded ? '<div class="rh-empty">Loading roster…</div>'
      : !list.length ? '<div class="rh-empty">No players registered for this team.</div>'
      : '<ul class="rh-list">' + list.map(p => {
          const code = COUNTRY_CODES[p.country.toLowerCase().trim()];
          const flag = code ? `<img class="rh-flag" src="https://flagcdn.com/w20/${code}.png" alt="" onerror="this.style.visibility='hidden'">` : '<span class="rh-flag rh-flag-empty"></span>';
          return `<li>${flag}<span class="rh-name">${esc(p.name)}${p.roster && p.roster !== 'Starter' ? ` <span class="rh-bench">${esc(p.roster)}</span>` : ''}</span><span class="rh-role">${esc(p.roles.map(x => ROLE_LABELS[x] || x).join(' · '))}</span></li>`;
        }).join('') + '</ul>';
    return `<div class="rh-head">${logoImg(team, 'team-logo-md')}<div><div class="rh-team">${esc(team)}</div><div class="rh-kicker">${r ? '#' + r.pos + ' · ' : ''}Roster</div></div></div>${body}<div class="rh-foot">Click to open team page →</div>`;
  }
  function show(a) {
    cur = a; card.innerHTML = html(a.dataset.team); card.classList.add('visible');
    const r = a.getBoundingClientRect(), cw = card.offsetWidth, ch = card.offsetHeight;
    let left = r.right + 12; if (left + cw > window.innerWidth - 8) left = Math.max(8, r.left - cw - 12);
    card.style.left = left + 'px';
    card.style.top = Math.max(8, Math.min(r.top + r.height / 2 - ch / 2, window.innerHeight - ch - 8)) + 'px';
  }
  const hide = () => { cur = null; card.classList.remove('visible'); };
  document.addEventListener('mouseover', e => { const a = e.target.closest('a[data-team]'); if (a) { if (a !== cur) show(a); } else if (cur) hide(); });
  document.addEventListener('focusin', e => { const a = e.target.closest('a[data-team]'); if (a) show(a); else if (cur) hide(); });
  window.addEventListener('scroll', hide, true);
})();

/* ── search: teams and every player, including those without a team ── */
(() => {
  const input = $('#nxSearch'), list = $('#nxSearchList');
  if (!input) return;
  let items = [], active = -1;
  const norm = s => String(s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  function search(q) {
    q = norm(q.trim()); if (!q) return [];
    const score = n => { const x = norm(n); return x === q ? 0 : x.startsWith(q) ? 1 : x.includes(q) ? 2 : -1; };
    const teams = state.ranking.map(t => ({ kind: 'Team', name: t.team, sub: `#${t.pos} · ${REGION_LABELS[t.region] || t.region || ''}`, url: teamUrl(t.team), s: score(t.team) }));
    const players = state.players.map(p => ({ kind: 'Player', name: p.name, sub: p.team && p.team.toLowerCase() !== 'none' ? p.team : 'Free agent', url: playerUrl(p.name), s: score(p.name), country: p.country }));
    return [...teams, ...players].filter(x => x.s >= 0).sort((a, b) => a.s - b.s || a.name.length - b.name.length || a.name.localeCompare(b.name)).slice(0, 12);
  }
  function draw() {
    if (!input.value.trim()) { list.hidden = true; return; }
    list.hidden = false;
    list.innerHTML = items.length ? items.map((x, i) => `<a class="nx-res${i === active ? ' on' : ''}" href="${x.url}" role="option" aria-selected="${i === active}">
        <span class="nx-kind ${x.kind.toLowerCase()}">${x.kind}</span>
        <span class="nx-name">${x.kind === 'Team' ? logoImg(x.name) : flagImg(x.country)}${esc(x.name)}</span><span class="nx-sub">${esc(x.sub)}</span></a>`).join('')
      : `<div class="nx-none">${state.playersLoaded ? 'No team or player matches that.' : 'Still loading the player list…'}</div>`;
  }
  input.addEventListener('input', () => { items = search(input.value); active = items.length ? 0 : -1; draw(); });
  input.addEventListener('focus', () => { if (input.value.trim()) { items = search(input.value); draw(); } });
  input.addEventListener('keydown', e => {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); if (items.length) { active = (active + (e.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length; draw(); } }
    else if (e.key === 'Enter' && items[active]) { location.href = items[active].url; }
    else if (e.key === 'Escape') { list.hidden = true; input.blur(); e.stopPropagation(); }
  });
  document.addEventListener('click', e => { if (!e.target.closest('.nx-search')) list.hidden = true; });
  document.addEventListener('keydown', e => { if (e.key === '/' && !/^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement.tagName)) { e.preventDefault(); input.focus(); } });
})();

/* ══════════════════════════════════════════════════════════════
   PLAYER PAGE
══════════════════════════════════════════════════════════════ */
function renderPlayerHero() {
  const p = state.player;
  document.title = `${p.name} · NEXUS`;
  const hasTeam = p.team && p.team.toLowerCase() !== 'none';
  const a = playerAwards(p.name);
  const groups = groupMedals(a.medals);
  const majors = Math.max(p.majorWinner, a.placements.filter(h => h.tier >= 5 && h.position === 'Winner').length);
  const evs = state._plEvents || [];
  const stat = (l, v, cls) => `<div><div class="hero-stat-l">${l}</div><div class="hero-stat-v${cls || ''}">${v}</div></div>`;
  $('#plHero').innerHTML = `
    <div class="hero-main">
      <div class="hero-logo photo">${photoHtml(p.name)}</div>
      <div class="hero-id">
        <div class="hero-kicker"><span>Player profile</span>
          ${p.country ? `<span class="hero-chip">${flagImg(p.country)} ${esc(p.country)}</span>` : ''}
          <span class="hero-chip accent">${hasTeam ? esc(p.roster || 'Player') : 'Free agent'}</span></div>
        <h1 class="hero-name">${esc(p.name)}</h1>
        <div class="hero-form"><span class="pl-team">${hasTeam ? teamLink(p.team, true) : 'No team'}</span>
          <span class="tp-roles" style="min-height:0 !important">${roleChips(p.roles)}</span></div>
      </div>
      <div class="hero-rank pl-medals">
        <div class="hero-rank-label">Medals</div>
        ${groups.length ? `<div class="shelf hm">${groups.map(g => medalIconHtml(g, { focusable: true })).join('')}</div>` : '<div class="hero-rank-sub">No MVP, EVP or VP yet</div>'}
      </div>
    </div>
    <div class="hero-stats">
      ${stat('Market value', fmtMoney(p.marketValue))}
      ${stat('Titles', p.titles, p.titles ? ' hot' : '')}
      ${stat('Majors', majors, majors ? ' hot' : '')}
      ${stat('MVP', p.mvp, p.mvp ? ' hot' : '')}
      ${stat('EVP', p.evp)}
      ${stat('VP', p.vp)}
      ${stat('Events', evs.length || '—')}
    </div>`;
}
function renderPlayerCabinet() {
  const p = state.player, box = $('#plAch'), btn = $('#ppAchToggle');
  const a = playerAwards(p.name), view = getView('player');
  btn.textContent = toggleLabel(view); btn.setAttribute('aria-pressed', view === 'detailed'); btn.hidden = !a.hist.length;
  $('#plCabCount').textContent = a.trophies.length ? `${a.trophies.length} on the shelf` : '';
  if (!state.historyLoaded) { box.innerHTML = '<div class="shelf-none">Achievements file (file/history_updated.xlsx) could not be loaded.</div>'; return; }
  if (!a.hist.length) { box.innerHTML = '<div class="shelf-none">No achievements recorded for this player yet.</div>'; return; }
  if (view === 'icons') {
    const others = a.placements.length - a.trophies.length;
    box.innerHTML = (a.trophies.length ? `<div class="shelf lg">${a.trophies.map(h => trophyIconHtml(h, { playerRow: true, focusable: true })).join('')}</div>` : '<div class="shelf-none">No top-3 finishes yet.</div>')
      + (others ? `<div class="cab-foot">${others} top 4–8 finish${others > 1 ? 'es' : ''} in the expanded view</div>` : '');
    return;
  }
  let html = '';
  if (a.medals.length) html += `<div class="ach-sub">Individual awards · ${a.medals.length}</div><div class="ach-grid">${a.medals.map(h => achItemHtml(h, { playerRow: true, showTeam: true })).join('')}</div>`;
  if (a.placements.length) html += `<div class="ach-sub">Trophies &amp; placements · ${a.placements.length}</div><div class="ach-grid">${a.placements.map(h => achItemHtml(h, { playerRow: true, showTeam: true })).join('')}</div>`;
  box.innerHTML = html;
}
function renderPlayerOverview() {
  const p = state.player;
  $('#plCareer').innerHTML = state.historyLoaded ? careerHtml(p, true) : '<div class="tp-empty">History file could not be loaded.</div>';
  const hasTeam = p.team && p.team.toLowerCase() !== 'none';
  const box = $('#plTeam');
  if (!hasTeam) { box.innerHTML = '<div class="tp-empty">Currently without a team.</div>'; return; }
  const t = state.ranking.find(x => sameTeam(x.team, p.team));
  const mates = state.players.filter(x => sameTeam(x.team, p.team) && x.name !== p.name);
  box.innerHTML = `<div class="pl-teamcard">
      <a class="pl-teamhead" href="${teamUrl(p.team)}">${logoImg(p.team, 'team-logo-lg')}<span><b>${esc(p.team)}</b>
        <small>${t ? `#${t.pos} world · #${t.regionPos} ${esc(REGION_LABELS[t.region] || t.region)} · ${fmtPts(t.points)} pts` : 'not in the current ranking'}</small></span><i>Open team page →</i></a>
      <div class="pl-mates"><span class="lu-l">Teammates</span>${mates.length ? mates.map(m => `<span class="pl-mate">${flagImg(m.country)}${playerLink(m.name)}<small>${esc(m.roles.map(r => ROLE_LABELS[r] || r).join(' · '))}</small></span>`).join('') : '<span class="cr-fa">none listed</span>'}</div>
    </div>`;
}
function renderPlayerEvents() {
  const box = $('#plEvents');
  const list = state._plEvents = playerEvents(state.player.name);
  $('#tabHistCount').textContent = list.length || '';
  $('#plEventsNote').textContent = list.length ? `${list.length} event${list.length > 1 ? 's' : ''}` : '';
  if (!list.length) { box.innerHTML = '<div class="tp-empty">No events on record for this player.</div>'; return; }
  box.innerHTML = list.map(f => tourRowHtml(f.o, f.team, { showTeam: true, source: f.source })).join('')
    + (list.some(f => f.source === 'estimated') ? '<div class="cab-foot">“Estimated” = no lineup was recorded for that team at that event, so it was rebuilt from today’s roster and the logged transfers.</div>' : '');
}
function renderPlayerAll() { tips = []; tourCache = new Map(); state._evDates = null; renderPlayerEvents(); renderPlayerHero(); renderPlayerCabinet(); renderPlayerOverview(); }

/* ── boot ── */
function showNotFound(msg, what) {
  document.title = 'Not found · NEXUS';
  $('#tpContent').innerHTML = `<div class="card"><div class="tp-notfound">
    <div class="card-title">${esc(msg)}</div>
    <p style="margin:14px 0 20px" class="text-muted">${what || 'Pick a team from the ranking, or use the search at the top.'}</p>
    <a class="btn btn-primary" href="index.html#global" style="text-decoration:none">← Global Ranking</a></div></div>`;
}
function renderAll() {
  tips = []; tourCache = new Map(); state._evDates = null;
  renderHero(); renderAchievements(); renderTrend(); renderRoster(); renderFormer(); renderTournaments(); renderMatches();
  if (state.openPlayer) renderPlayerPanel(false);
}

async function bootTeam() {
  state.matchLimit = MATCH_PAGE;
  const [rankRes] = await Promise.allSettled([loadRanking(), loadLogos()]);
  if (rankRes.status === 'rejected') { console.error(rankRes.reason); showNotFound('Ranking file (file/ranking.xlsx) could not be loaded.'); return; }
  state.rankedSet = new Set(state.ranking.map(r => r.team.toLowerCase()));
  if (!state.teamParam) { showNotFound('No team selected.'); loadPlayers().catch(() => {}); return; }
  state.team = state.ranking.find(t => sameTeam(t.team, state.teamParam));
  if (!state.team) { showNotFound(`“${state.teamParam}” is not in the current ranking.`); loadPlayers().catch(() => {}); return; }
  setTab(location.hash === '#history' ? 'history' : 'overview', false);
  renderHero();

  const [pl, hi, ma, ol] = await Promise.allSettled([loadPlayers(), loadHistory(), loadMatches(), loadOldRanking()]);
  [pl, hi, ma, ol].forEach(r => { if (r.status === 'rejected') console.warn(r.reason); });
  state.trajectory = [];
  renderAll();
  const missing = [pl.status === 'rejected' && 'players_updated.xlsx', hi.status === 'rejected' && 'history_updated.xlsx'].filter(Boolean);
  setStatus(missing.length ? 'Loaded with missing files: ' + missing.join(', ') : 'Team data loaded · loading ranking archive, trophy icons and event details…', missing.length ? 'err' : 'loading');

  // Slower extras, each optional: the page is already usable without them
  const extras = await Promise.allSettled([
    loadTrajectory().then(renderTrend),
    loadAwards().then(() => { tips = []; renderAchievements(); renderRoster(); if (state.openPlayer) renderPlayerPanel(false); }),
    loadEvents().then(() => { tourCache = new Map(); renderTournaments(); renderFormer(); })
  ]);
  extras.forEach(r => { if (r.status === 'rejected') console.warn(r.reason); });
  const off = [extras[1].status === 'rejected' && 'trophy icons (assets/awards.json)', extras[2].status === 'rejected' && 'event details (file/kz-events-backup.json)'].filter(Boolean);
  if (!missing.length) setStatus(off.length ? 'Loaded · not available: ' + off.join(', ') : 'All data loaded', 'ok');
}

async function bootPlayer() {
  const [rankRes, , pl] = await Promise.allSettled([loadRanking(), loadLogos(), loadPlayers()]);
  if (rankRes.status === 'fulfilled') state.rankedSet = new Set(state.ranking.map(r => r.team.toLowerCase())); else state.rankedSet = new Set();
  if (pl.status === 'rejected') { console.error(pl.reason); showNotFound('Players file (file/players_updated.xlsx) could not be loaded.'); return; }
  const want = new URLSearchParams(location.search).get('player') || '';
  if (!want) { showNotFound('No player selected.', 'Use the search at the top to find a player.'); return; }
  state.player = findPlayer(want);
  if (!state.player) { showNotFound(`“${want}” is not in the players file.`, 'Use the search at the top to find a player.'); return; }
  setTab(location.hash === '#history' ? 'history' : 'overview', false);
  state._plEvents = [];
  renderPlayerHero();

  const [hi, ma] = await Promise.allSettled([loadHistory(), loadMatches()]);
  [hi, ma].forEach(r => { if (r.status === 'rejected') console.warn(r.reason); });
  renderPlayerAll();
  setStatus(hi.status === 'rejected' ? 'Loaded with missing file: history_updated.xlsx' : 'Player data loaded · loading trophy icons and event details…', hi.status === 'rejected' ? 'err' : 'loading');
  const extras = await Promise.allSettled([
    loadAwards().then(() => { tips = []; renderPlayerHero(); renderPlayerCabinet(); }),
    loadEvents().then(renderPlayerAll)
  ]);
  extras.forEach(r => { if (r.status === 'rejected') console.warn(r.reason); });
  const off = [extras[0].status === 'rejected' && 'trophy icons (assets/awards.json)', extras[1].status === 'rejected' && 'event details (file/kz-events-backup.json)'].filter(Boolean);
  if (hi.status !== 'rejected') setStatus(off.length ? 'Loaded · not available: ' + off.join(', ') : 'All data loaded', 'ok');
}

document.addEventListener('click', e => {
  if (e.target.closest('#ppAchToggle') && state.player) { setView('player', getView('player') === 'detailed' ? 'icons' : 'detailed'); hideTip(); renderPlayerCabinet(); }
});
(document.body.dataset.page === 'player' ? bootPlayer : bootTeam)();
