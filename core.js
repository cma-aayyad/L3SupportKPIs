/* L2/L3 Support KPI dashboard — core logic.
 * DOM-free: loaded by index.html with a plain <script> tag (window.TrackerCore)
 * and by Node with require('./core.js') for tests and the publish script.
 * Sections: CONFIG · PARSING · NORMALIZATION · AGGREGATION · PERIODS
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TrackerCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ===================================================================== */
  /* CONFIG                                                                */
  /* ===================================================================== */
  const CONFIG = {
    // ServiceNow instance, e.g. 'https://cmacgm.service-now.com'. Empty = ticket IDs stay plain text.
    serviceNowBaseUrl: 'https://cma.service-now.com',
    // {base} and {id} are replaced. text_search_exact_match works for every table below.
    ticketUrlTemplate: '{base}/text_search_exact_match.do?sysparm_search={id}',
    ticketPatterns: [/^INC\d+$/i, /^RITM\d+$/i, /^CHG\d+$/i, /^SCTASK\d+$/i],

    // Status → bucket. First match wins, so the "not done" rule must stay above Done.
    statusBuckets: [
      { bucket: 'Pending', re: /\bincomplete\b|\bnot\s+(?:yet\s+)?(?:completed?|done|resolved)\b/i },
      { bucket: 'Done', re: /\b(?:completed?|done|closed|resolved|delivered|deployed)\b/i },
      { bucket: 'In Progress', re: /\b(?:in[\s-]?progress|ongoing|wip|working)\b/i },
      { bucket: 'Pending', re: /\b(?:pending|on[\s-]?hold|waiting|blocked)\b/i },
      { bucket: 'Requested', re: /\b(?:requested|new|open|to[\s-]?do|not\s+started)\b/i },
    ],
    bucketOrder: ['Requested', 'Pending', 'In Progress', 'Done', 'Unknown'],
    doneBucket: 'Done',

    // Applied to Type; to TASK only when Type is blank. \b keeps "tissue" out.
    issueRe: /\b(?:incidents?|issues?|bugs?|errors?|failures?|outages?|problems?|defects?)\b/i,
    deploymentRe: /\b(?:deploy(?:ment|ments|ed)?|change\s*requests?|releases?|CR|CHG)\b/i,

    agingBuckets: [
      { label: '0–7 d', max: 7 },
      { label: '8–14 d', max: 14 },
      { label: '15–30 d', max: 30 },
      { label: '30+ d', max: Infinity },
    ],
    ageHighlightDays: 14,

    // Delta colour: 'up' = increase is good, 'down' = decrease is good, 'neutral' = no colour.
    kpiDirection: {
      total: 'neutral', done: 'up', open: 'down', completionRate: 'up', issues: 'down',
      openIssues: 'down', deployments: 'up', pipeline: 'neutral', avgAge: 'down', activeApps: 'neutral',
    },

    unassigned: 'Unassigned',
    noApp: '(no application)',
  };

  /* ===================================================================== */
  /* PARSING                                                               */
  /* ===================================================================== */
  const pad = n => String(n).padStart(2, '0');

  // SheetJS date cells → 'yyyy-mm-dd'. +12h makes the result immune to UTC/local-midnight offsets.
  function dateToYMD(d) {
    const t = new Date(d.getTime() + 12 * 3600 * 1000);
    return `${t.getFullYear()}-${pad(t.getMonth() + 1)}-${pad(t.getDate())}`;
  }

  /** Read every sheet into plain rows; real date cells become 'yyyy-mm-dd' strings. */
  function sheetsFromWorkbook(wb, XLSX) {
    const out = {};
    for (const name of wb.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '', raw: true });
      out[name] = rows.map(r => {
        const o = {};
        for (const k of Object.keys(r)) {
          const v = r[k];
          o[String(k).trim()] = v instanceof Date ? dateToYMD(v) : v;
        }
        return o;
      });
    }
    return out;
  }
  const READ_OPTS = { type: 'array', cellDates: true };

  const normKey = s => String(s == null ? '' : s).trim().replace(/\s+/g, ' ').toLowerCase();

  function headersOf(rows) {
    const set = new Set();
    (rows || []).slice(0, 50).forEach(r => Object.keys(r).forEach(k => set.add(k)));
    return [...set];
  }
  // Column lookup tolerant to case, spaces and extra columns.
  function colFinder(rows) {
    const map = new Map(headersOf(rows).map(h => [normKey(h).replace(/[\s_]/g, ''), h]));
    return (...names) => {
      for (const n of names) {
        const h = map.get(normKey(n).replace(/[\s_]/g, ''));
        if (h !== undefined) return h;
      }
      return null;
    };
  }

  /** Find TaskList / AppDomain / Team sheets without relying on exact headers. */
  function detectSheets(sheets) {
    let tasks = null, mapping = null, team = null;
    const names = Object.keys(sheets || {});
    const info = names.map(name => {
      const col = colFinder(sheets[name]);
      return { name, col, hasTask: !!col('TASK'), hasApp: !!col('Application'), hasDomain: !!col('Domain'),
        hasName: !!col('Name'), hasAliases: !!col('Aliases') };
    });
    tasks = info.find(i => normKey(i.name) === 'tasklist' && i.hasTask) || info.find(i => i.hasTask) || null;
    mapping = info.find(i => normKey(i.name) === 'appdomain') ||
      info.find(i => i.hasApp && i.hasDomain && !i.hasTask) || null;
    team = info.find(i => normKey(i.name) === 'team' && i.hasName) ||
      info.find(i => i.hasName && i.hasAliases && !i.hasTask) || null;
    return {
      tasksSheet: tasks && tasks.name, mappingSheet: mapping && mapping.name, teamSheet: team && team.name,
      tasksRows: tasks ? sheets[tasks.name] : [], mapRows: mapping ? sheets[mapping.name] : [],
      teamRows: team ? sheets[team.name] : [],
    };
  }

  /* ===================================================================== */
  /* NORMALIZATION                                                         */
  /* ===================================================================== */

  /** Split multi-value cells on ; , newline and " / " — never on a bare "/" inside an ID. */
  function splitMulti(v) {
    return String(v == null ? '' : v).split(/\s*[;,\n]\s*|\s+\/\s+/).map(s => s.trim()).filter(Boolean);
  }

  function validYMD(y, m, d) {
    if (y < 1990 || y > 2100 || m < 1 || m > 12 || d < 1) return null;
    const dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
    return `${y}-${pad(m)}-${pad(d)}`;
  }
  /** Returns 'yyyy-mm-dd', '' for blank, or null when it can't be parsed safely. */
  function parseDate(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date) return isNaN(v) ? null : dateToYMD(v);
    if (typeof v === 'number') { // Excel serial in the Date column (1900 date system)
      if (v < 1 || v > 2958465) return null;
      const dt = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000);
      return validYMD(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
    }
    const s = String(v).trim();
    if (!s) return '';
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
    if (m) return validYMD(+m[1], +m[2], +m[3]);
    m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/); // day-first
    if (m) return validYMD(+m[3], +m[2], +m[1]);
    return null; // never new Date(string)
  }
  const ymdToUTC = s => { const [y, m, d] = s.split('-').map(Number); return Date.UTC(y, m - 1, d); };
  const daysBetween = (a, b) => Math.round((ymdToUTC(b) - ymdToUTC(a)) / 86400000);

  function bucketOf(status) {
    const s = String(status == null ? '' : status).trim();
    if (!s) return 'Unknown';
    for (const r of CONFIG.statusBuckets) if (r.re.test(s)) return r.bucket;
    return 'Unknown';
  }

  const tokenKey = name => normKey(name).replace(/[^\p{L}\p{N}\s]/gu, ' ').split(/\s+/).filter(Boolean).sort().join(' ');

  function buildPeople(teamRows) {
    const col = colFinder(teamRows);
    const nameCol = col('Name'), aliasCol = col('Aliases', 'Alias');
    const alias = new Map(); const members = [];
    for (const r of teamRows || []) {
      const name = String(r[nameCol] || '').trim();
      if (!name) continue;
      members.push(name);
      alias.set(tokenKey(name), name);
      for (const a of String((aliasCol && r[aliasCol]) || '').split(';').map(s => s.trim()).filter(Boolean)) alias.set(tokenKey(a), name);
    }
    const variants = new Map(); // key → Map(rawSpelling → count)
    return {
      hasTeam: members.length > 0, members,
      resolve(raw) {
        const k = tokenKey(raw);
        if (!k) return null;
        const teamName = alias.get(k);
        const key = teamName ? 'team:' + tokenKey(teamName) : 'p:' + k;
        if (!variants.has(key)) variants.set(key, new Map());
        const vm = variants.get(key); vm.set(raw.trim(), (vm.get(raw.trim()) || 0) + 1);
        return { key, isTeam: !!teamName, teamName: teamName || null };
      },
      displayName(key) {
        if (key.startsWith('team:')) return alias.get(key.slice(5)) || key.slice(5);
        const vm = variants.get(key); if (!vm) return key.slice(2);
        return [...vm.entries()].sort((a, b) => b[1] - a[1])[0][0];
      },
      variants,
    };
  }

  /**
   * Turn raw sheets into normalized tasks + quality findings.
   * @returns {{tasks, mapping, people, quality, meta}}
   */
  function normalize(sheets) {
    const det = detectSheets(sheets);
    const mcol = colFinder(det.mapRows);
    const mApp = mcol('Application', 'App'), mDom = mcol('Domain');
    const domainCanon = new Map(); // normKey(domain) → display spelling
    const mapping = new Map();     // normKey(app) → {app, domain}
    for (const r of det.mapRows) {
      const app = String((mApp && r[mApp]) || '').trim(), dom = String((mDom && r[mDom]) || '').trim();
      if (!app) continue;
      const dk = normKey(dom);
      if (dom && !domainCanon.has(dk)) domainCanon.set(dk, dom);
      if (!mapping.has(normKey(app))) mapping.set(normKey(app), { app, domain: dom ? domainCanon.get(dk) : CONFIG.unassigned });
    }
    const people = buildPeople(det.teamRows);

    const c = colFinder(det.tasksRows);
    const C = {
      task: c('TASK', 'Task'), type: c('Type'), app: c('Application', 'App'), date: c('Date'), poc: c('POC'),
      shadow: c('Shadowing', 'Shadow'), status: c('Status'), domain: c('Domain'), ticket: c('Ticket ID', 'TicketID', 'Ticket'),
    };
    const get = (r, k) => (C[k] ? r[C[k]] : '');
    const appCanon = new Map(); // for unmapped apps: first spelling seen

    const Q = { noTicket: [], unmappedApp: [], domainMismatch: [], unassigned: [], badDate: [], blankStatus: [], nameVariants: [], dupTickets: [] };
    const tasks = [];
    det.tasksRows.forEach((r, i) => {
      const taskText = String(get(r, 'task') || '').trim();
      const rawApp = String(get(r, 'app') || '').trim();
      // skip fully empty rows
      if (!taskText && !rawApp && !String(get(r, 'status') || '').trim() && !String(get(r, 'ticket') || '').trim()) return;
      const t = { idx: tasks.length, row: i + 2, task: taskText, type: String(get(r, 'type') || '').trim(), flags: [] };

      // application + domain (mapping is the source of truth)
      const ak = normKey(rawApp), colDom = String(get(r, 'domain') || '').trim();
      const mapped = ak ? mapping.get(ak) : null;
      if (mapped) {
        t.app = mapped.app; t.domain = mapped.domain; t.mapped = true;
        if (colDom && normKey(colDom) !== normKey(mapped.domain)) { t.flags.push('domainMismatch'); t.domainColumn = colDom; }
      } else {
        if (ak && !appCanon.has(ak)) appCanon.set(ak, rawApp);
        t.app = ak ? appCanon.get(ak) : CONFIG.noApp; t.mapped = false;
        if (ak) t.flags.push('unmappedApp');
        const dk = normKey(colDom);
        if (dk && !domainCanon.has(dk)) domainCanon.set(dk, colDom);
        t.domain = dk ? domainCanon.get(dk) : CONFIG.unassigned;
      }
      t.appKey = ak; t.domainColumn = t.domainColumn || colDom;
      if (t.domain === CONFIG.unassigned) t.flags.push('unassigned');

      // date
      const rawDate = get(r, 'date');
      const d = parseDate(rawDate);
      t.date = d || ''; t.dateRaw = rawDate instanceof Date ? dateToYMD(rawDate) : String(rawDate == null ? '' : rawDate);
      if (d === null || d === '') t.flags.push('badDate');

      // status
      t.status = String(get(r, 'status') || '').trim();
      t.bucket = bucketOf(t.status);
      if (!t.status) t.flags.push('blankStatus');
      t.isDone = t.bucket === CONFIG.doneBucket; t.isOpen = !t.isDone;

      // type flags
      t.isIssue = t.type ? CONFIG.issueRe.test(t.type) : CONFIG.issueRe.test(t.task);
      t.isDeployment = CONFIG.deploymentRe.test(t.type);

      // people
      t.pocRaw = String(get(r, 'poc') || '').trim(); t.shadowRaw = String(get(r, 'shadow') || '').trim();
      t.poc = splitMulti(t.pocRaw).map(n => people.resolve(n)).filter(Boolean);
      t.shadow = splitMulti(t.shadowRaw).map(n => people.resolve(n)).filter(Boolean);

      // tickets
      t.ticketRaw = String(get(r, 'ticket') == null ? '' : get(r, 'ticket')).trim();
      t.tickets = splitMulti(t.ticketRaw);
      if (!t.tickets.length) t.flags.push('noTicket');
      tasks.push(t);
    });

    // duplicate ticket IDs across rows
    const byTicket = new Map();
    tasks.forEach(t => new Set(t.tickets.map(x => x.toUpperCase())).forEach(id => {
      if (!byTicket.has(id)) byTicket.set(id, []); byTicket.get(id).push(t.idx);
    }));
    const dupIds = [];
    for (const [id, idxs] of byTicket) if (idxs.length > 1) { dupIds.push(id); idxs.forEach(i => { if (!tasks[i].flags.includes('dupTicket')) tasks[i].flags.push('dupTicket'); }); }

    // merged name variants
    const variantKeys = new Set();
    const variantList = [];
    for (const [key, vm] of people.variants) if (vm.size > 1) { variantKeys.add(key); variantList.push({ name: people.displayName(key), variants: [...vm.keys()] }); }
    tasks.forEach(t => { if ([...t.poc, ...t.shadow].some(p => variantKeys.has(p.key))) t.flags.push('nameVariant'); });

    for (const t of tasks) for (const f of t.flags) {
      const k = f === 'dupTicket' ? 'dupTickets' : f === 'nameVariant' ? 'nameVariants' : f;
      if (Q[k]) Q[k].push(t.idx);
    }
    const quality = {
      checks: [
        { id: 'noTicket', label: 'Tasks with no ticket ID', idx: Q.noTicket },
        { id: 'unmappedApp', label: 'Apps missing from AppDomain', idx: Q.unmappedApp,
          detail: [...new Set(Q.unmappedApp.map(i => tasks[i].app))] },
        { id: 'domainMismatch', label: 'Domain column differs from mapping', idx: Q.domainMismatch },
        { id: 'unassigned', label: 'Tasks with no domain (Unassigned)', idx: Q.unassigned },
        { id: 'badDate', label: 'Missing or unparseable dates', idx: Q.badDate },
        { id: 'blankStatus', label: 'Blank status', idx: Q.blankStatus },
        { id: 'nameVariants', label: 'Tasks with merged name variants', idx: Q.nameVariants,
          detail: variantList.map(v => `${v.name} ← ${v.variants.join(' | ')}`) },
        { id: 'dupTickets', label: 'Ticket IDs on more than one row', idx: Q.dupTickets, detail: dupIds, info: true },
      ],
    };
    quality.warningCount = quality.checks.filter(q => !q.info).reduce((s, q) => s + q.idx.length, 0);

    return {
      tasks, people, quality,
      mapping: { size: mapping.size, apps: [...mapping.values()], keys: new Set(mapping.keys()),
        domains: [...new Set([...mapping.values()].map(m => m.domain))] },
      meta: { tasksSheet: det.tasksSheet, mappingSheet: det.mappingSheet, teamSheet: det.teamSheet, hasTeam: people.hasTeam, columns: C },
    };
  }

  /* ===================================================================== */
  /* AGGREGATION                                                           */
  /* ===================================================================== */

  /** Reference date for aging: today for the current (or a future) month, the month's last day for past months. */
  function refDateFor(period, todayYMD) {
    if (!period) return todayYMD;
    if (period >= todayYMD.slice(0, 7)) return todayYMD;
    const [y, m] = period.split('-').map(Number);
    const last = new Date(Date.UTC(y, m, 0));
    return `${last.getUTCFullYear()}-${pad(last.getUTCMonth() + 1)}-${pad(last.getUTCDate())}`;
  }
  function applyAges(tasks, refYMD) {
    for (const t of tasks) t.age = (t.isOpen && t.date) ? Math.max(0, daysBetween(t.date, refYMD)) : null;
    return tasks;
  }
  const ageBucket = age => CONFIG.agingBuckets.findIndex(b => age <= b.max);

  /** filters: {domain, app, status ('Open' | bucket), type, person (key), q, qc (quality check id)} */
  function applyFilters(tasks, f, ds) {
    f = f || {};
    const qcSet = f.qc && ds ? new Set((ds.quality.checks.find(c => c.id === f.qc) || { idx: [] }).idx) : null;
    const q = normKey(f.q);
    return tasks.filter(t =>
      (!f.domain || normKey(t.domain) === normKey(f.domain)) &&
      (!f.app || normKey(t.app) === normKey(f.app)) &&
      (!f.status || (f.status === 'Open' ? t.isOpen : t.bucket === f.status)) &&
      (!f.type || normKey(t.type) === normKey(f.type)) &&
      (!f.person || t.poc.some(p => p.key === f.person) || t.shadow.some(p => p.key === f.person)) &&
      (!qcSet || qcSet.has(t.idx)) &&
      (!q || normKey([t.task, t.app, t.domain, t.ticketRaw, t.pocRaw, t.shadowRaw, t.type, t.status].join(' ')).includes(q)));
  }

  function countBy(list, fn) {
    const m = new Map();
    for (const x of list) { const k = fn(x); m.set(k, (m.get(k) || 0) + 1); }
    return m;
  }
  function stats(list) {
    const total = list.length, done = list.filter(t => t.isDone).length;
    const aged = list.filter(t => t.age != null);
    return {
      total, done, open: total - done,
      completionRate: total ? (done / total) * 100 : 0,
      issues: list.filter(t => t.isIssue).length,
      openIssues: list.filter(t => t.isIssue && t.isOpen).length,
      deployments: list.filter(t => t.isDeployment && t.isDone).length,
      pipeline: list.filter(t => t.isDeployment && !t.isDone).length,
      avgAge: aged.length ? aged.reduce((s, t) => s + t.age, 0) / aged.length : null,
    };
  }

  /** KPIs and breakdowns for an (already filtered) task list. */
  function aggregate(list, ds) {
    const k = stats(list);
    // Active applications: distinct mapped apps with at least one task, out of the mapping size.
    const activeKeys = new Set(list.filter(t => t.appKey).map(t => t.appKey));
    if (ds.mapping.size) {
      k.activeApps = [...activeKeys].filter(a => ds.mapping.keys.has(a)).length;
      k.totalApps = ds.mapping.size;
    } else { k.activeApps = activeKeys.size; k.totalApps = activeKeys.size; }
    k.unmappedActive = [...activeKeys].filter(a => !ds.mapping.keys.has(a)).length;

    const domains = new Map();
    for (const t of list) {
      if (!domains.has(t.domain)) domains.set(t.domain, new Map());
      const apps = domains.get(t.domain);
      if (!apps.has(t.app)) apps.set(t.app, []);
      apps.get(t.app).push(t);
    }
    const byDomain = [...domains.entries()].map(([domain, apps]) => {
      const all = [...apps.values()].flat();
      return { domain, ...stats(all), apps: [...apps.entries()].map(([app, ts]) => ({ app, domain, ...stats(ts) })).sort((a, b) => b.total - a.total) };
    }).sort((a, b) => (a.domain === CONFIG.unassigned) - (b.domain === CONFIG.unassigned) || b.total - a.total);
    const byApp = byDomain.flatMap(d => d.apps).sort((a, b) => b.total - a.total);

    const byBucket = CONFIG.bucketOrder.map(b => ({ bucket: b, count: list.filter(t => t.bucket === b).length })).filter(x => x.count || x.bucket !== 'Unknown');
    const byType = [...countBy(list, t => t.type || '(blank)').entries()].map(([type, count]) => ({ type, count })).sort((a, b) => b.count - a.count);

    const aging = CONFIG.agingBuckets.map(b => ({ label: b.label, count: 0 }));
    list.forEach(t => { if (t.age != null) aging[ageBucket(t.age)].count++; });

    const topOpenIssues = list.filter(t => t.isIssue && t.isOpen).sort((a, b) => (b.age ?? -1) - (a.age ?? -1)).slice(0, 5);

    // completion trend within the month (by task date)
    const byDate = new Map();
    list.filter(t => t.date).forEach(t => { const e = byDate.get(t.date) || { total: 0, done: 0 }; e.total++; if (t.isDone) e.done++; byDate.set(t.date, e); });
    let ct = 0, cd = 0;
    const daily = [...byDate.keys()].sort().map(d => { ct += byDate.get(d).total; cd += byDate.get(d).done; return { date: d, total: ct, done: cd }; });

    return { kpis: k, byDomain, byApp, byBucket, byType, aging, topOpenIssues, daily, team: teamWorkload(list, ds) };
  }

  /** Team workload: with a Team sheet, a task counts for every member in POC or Shadowing (once each). */
  function teamWorkload(list, ds) {
    const P = ds.people, hasTeam = P.hasTeam;
    const team = new Map(), req = new Map();
    const row = key => ({ key, name: P.displayName(key), poc: 0, shadow: 0, total: 0, open: 0, ages: [], aging: CONFIG.agingBuckets.map(() => 0) });
    for (const t of list) {
      const roles = new Map();
      t.poc.forEach(p => roles.set(p.key, { p, role: 'poc' }));
      if (hasTeam) t.shadow.forEach(p => { if (!roles.has(p.key)) roles.set(p.key, { p, role: 'shadow' }); });
      for (const { p, role } of roles.values()) {
        const target = !hasTeam || p.isTeam ? team : (role === 'poc' ? req : null);
        if (!target) continue;
        if (!target.has(p.key)) target.set(p.key, row(p.key));
        const e = target.get(p.key);
        e[role]++; e.total++;
        if (t.isOpen) { e.open++; if (t.age != null) { e.ages.push(t.age); e.aging[ageBucket(t.age)]++; } }
      }
      const covered = hasTeam ? [...roles.values()].some(r => r.p.isTeam) : t.poc.length > 0;
      if (!covered) {
        const key = 'p:unassigned';
        if (!team.has(key)) team.set(key, { ...row(key), name: CONFIG.unassigned });
        const e = team.get(key); e.poc++; e.total++; if (t.isOpen) { e.open++; if (t.age != null) { e.ages.push(t.age); e.aging[ageBucket(t.age)]++; } }
      }
    }
    const fin = m => [...m.values()].map(e => ({ ...e, avgAge: e.ages.length ? e.ages.reduce((a, b) => a + b, 0) / e.ages.length : null })).sort((a, b) => b.total - a.total);
    return { hasTeam, members: fin(team), requesters: fin(req) };
  }

  /* ===================================================================== */
  /* PERIODS                                                               */
  /* ===================================================================== */
  const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  function periodFromFileName(name) {
    const s = String(name || '').toLowerCase();
    let m = s.match(/(20\d{2})[-_ .](0[1-9]|1[0-2])(?!\d)/);
    if (m) return `${m[1]}-${m[2]}`;
    m = s.match(/\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)[\s_.-]*(20\d{2})\b/);
    if (m) return `${m[2]}-${pad(MONTHS.indexOf(m[1].slice(0, 3)) + 1)}`;
    return null;
  }
  function periodFromTasks(tasks) {
    const ds = tasks.map(t => t.date).filter(Boolean).sort();
    return ds.length ? ds[ds.length - 1].slice(0, 7) : null;
  }
  function detectPeriod(fileName, sheets) {
    return periodFromFileName(fileName) || periodFromTasks(normalize(sheets).tasks);
  }
  function periodLabel(p) {
    if (!p) return '';
    const [y, m] = p.split('-').map(Number);
    return ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][m - 1] + ' ' + y;
  }

  // Snapshot file text (shared by the publish button and scripts/publish.js)
  function snapshotJs(snap) {
    return `/* Generated by L2/L3 KPI publish — do not edit by hand. */\nwindow.TRACKER_SNAPSHOT = ${JSON.stringify(snap)};\n`;
  }
  function historyJs(snap) {
    return `/* Generated by L2/L3 KPI publish — do not edit by hand. */\nwindow.TRACKER_HISTORY = window.TRACKER_HISTORY || {};\nwindow.TRACKER_HISTORY[${JSON.stringify(snap.period)}] = ${JSON.stringify(snap)};\n`;
  }
  function historyIndexJs(list) {
    const sorted = [...list].sort((a, b) => b.period.localeCompare(a.period));
    return `/* Generated by L2/L3 KPI publish — available periods, newest first. */\nwindow.TRACKER_HISTORY_INDEX = ${JSON.stringify(sorted, null, 2)};\n`;
  }
  function mergeIndex(list, entry) {
    return [...(list || []).filter(e => e.period !== entry.period), entry].sort((a, b) => b.period.localeCompare(a.period));
  }

  function ticketUrl(id) {
    if (!CONFIG.serviceNowBaseUrl || !CONFIG.ticketPatterns.some(re => re.test(id))) return null;
    return CONFIG.ticketUrlTemplate.replace('{base}', CONFIG.serviceNowBaseUrl.replace(/\/+$/, '')).replace('{id}', encodeURIComponent(id));
  }

  return {
    CONFIG, READ_OPTS, sheetsFromWorkbook, detectSheets, normalize, splitMulti, parseDate, bucketOf, tokenKey,
    refDateFor, applyAges, applyFilters, aggregate, stats, daysBetween, normKey,
    periodFromFileName, periodFromTasks, detectPeriod, periodLabel, snapshotJs, historyJs, historyIndexJs, mergeIndex, ticketUrl,
  };
});
