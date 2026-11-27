import ExcelJS from 'exceljs';
import type {
  AbnormalItem,
  CardRecord,
  ConsolidatedAbnormalRow,
  ConsolidatedCardOnlyRow,
  ConsolidatedDay,
  ConsolidatedMatchedRow,
  GpsReportData,
  MatchedEntry,
} from '../types/fuel';
import {
  buildTimeVarianceRows,
  fmtDT,
  gpsDtList,
  normFuel,
  timeRisk,
  UNCLASSIFIED_LABEL,
  VARIANCE_TOLERANCE_L,
  vtype,
} from './gpsMatcherParsers';

const HDR_STRIP = [
  /Card Date & Time GPS Date & Time Vehicle Vehicle Type Shift Area Location\(s\) Pumped \(Card\) Pumped \(GPS\) Variance Vari % Status Note/g,
  /Vehicle Vehicle Type Card Date & Time GPS Date & Time Diff \(Mins\) GPS Location Risk/g,
  /Category Vehicle Observed Issue Fuel Impact \(L\) Action/g,
  /Page \d+ of \d+/g,
];

const numOf = (v: unknown): number => {
  const s = String(v == null ? '' : v)
    .replace(/,/g, '')
    .trim();
  if (s === '' || s === '-') return 0;
  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
};

const DT = '\\d{4}-\\d{2}-\\d{2}';

function splitSections(t: string) {
  const defs: [string, RegExp][] = [
    ['s1', /1\. Unusual filling variance \((\d+)\)/],
    ['s2', /2\. More than ±1\.5 L Variance[^()]{0,60}\((\d+)\)/],
    ['s3', /3\. GPS Unavailable & Offline Vehicle List \((\d+)\)/],
    ['s4', /4\. Fuel Pump Time Variance[^()]{0,80}\((\d+)\)/],
    ['s5', /5\. On GPS \([^)]*\), not on card \((\d+)\)/],
    ['s51', /5\.1 GPS fills not at [^,]*, not on card \((\d+)\)/],
    ['s6', /6\. Variance within [^()]{0,30}\((\d+)\)/],
    ['s7', /7\. Abnormal Highlights & Recommendations \((\d+)\)/],
    ['gt', /GRAND TOTAL /],
  ];
  const found: { k: string; start: number; bodyStart: number; count: number | null }[] = [];
  defs.forEach(([k, re]) => {
    const m = re.exec(t);
    if (m) found.push({ k, start: m.index, bodyStart: m.index + m[0].length, count: m[1] !== undefined ? +m[1] : null });
  });
  found.sort((a, b) => a.start - b.start);
  const out: Record<string, { count: number | null; body: string }> = {};
  found.forEach((f, i) => {
    out[f.k] = { count: f.count, body: t.slice(f.bodyStart, i + 1 < found.length ? found[i + 1].start : t.length) };
  });
  return out;
}

function parseMatched(seg: string): ConsolidatedMatchedRow[] {
  const re = new RegExp(
    '(' +
      DT +
      ') (\\d{2}:\\d{2}) ((?:' +
      DT +
      ' \\d{2}:\\d{2} )+)(\\S+) (.+?) ([\\d,]+\\.\\d{2}) ([\\d,]+\\.\\d{2}) ([+-][\\d,]+\\.\\d{2}) ([+-][\\d.,]+)% (UNUSUAL|Mismatch|OK) (.*?)(?= ' +
      DT +
      ' \\d{2}:\\d{2} ' +
      DT +
      '|Subtotal|$)',
    'g'
  );
  const rows: ConsolidatedMatchedRow[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(seg))) {
    const mid = m[5].match(/^(.*?) (-|[A-Z]) (\S+) ?(.*)$/);
    rows.push({
      date: m[1],
      cardTime: m[2],
      gpsTimes: m[3]
        .trim()
        .split(/ (?=\d{4}-)/)
        .map((x) => x.split(' ')[1])
        .join(', '),
      vehicle: m[4],
      type: mid ? mid[1] : m[5],
      area: mid ? mid[3] : '',
      cardL: numOf(m[6]),
      gpsL: numOf(m[7]),
      variance: numOf(m[8]),
      pct: numOf(m[9]),
      status: m[10],
      note: m[11].trim(),
    });
  }
  return rows;
}

function parseCardOnly(body: string) {
  const subRe = /3\.(\d) ([^()]+?) \((\d+)\)/g;
  const marks: { idx: number; end: number; label: string; count: number }[] = [];
  let m: RegExpExecArray | null;
  while ((m = subRe.exec(body))) {
    marks.push({ idx: m.index, end: subRe.lastIndex, label: m[2].split('—')[0].trim(), count: +m[3] });
  }
  const rows: ConsolidatedCardOnlyRow[] = [];
  const rowRe = new RegExp(
    '(' +
      DT +
      ')(?: (\\d{2}:\\d{2}))? - (.+?) ([\\d,]+\\.\\d{2}) - - - No GPS record (.*?)(?= ' +
      DT +
      '(?: \\d{2}:\\d{2})? - |Subtotal|$)',
    'g'
  );
  marks.forEach((mk, i) => {
    const part = body.slice(mk.end, i + 1 < marks.length ? marks[i + 1].idx : body.length);
    let r: RegExpExecArray | null;
    rowRe.lastIndex = 0;
    while ((r = rowRe.exec(part))) {
      const mid = r[3].replace(/ -$/, '').trim();
      let vehicle = '';
      let rest = '';
      const cm = mid.match(/^(Common\s*-\s*\w+)\s*(.*)$/);
      if (cm) {
        vehicle = cm[1].replace(/\s*-\s*/, ' - ');
        rest = cm[2];
      } else {
        const p = mid.split(' ');
        vehicle = p.shift() || '';
        rest = p.join(' ');
      }
      const tk = rest.split(' ');
      const area = tk.length ? tk.pop() || '' : '';
      const shift = tk.length ? tk.pop() || '' : '';
      rows.push({
        date: r[1],
        cardTime: r[2] || '',
        vehicle,
        type: tk.join(' '),
        shift,
        area,
        cardL: numOf(r[4]),
        note: r[5].trim(),
        category: mk.label,
      });
    }
  });
  return { rows, subCounts: marks.map((x) => ({ label: x.label, count: x.count })) };
}

function parseTime(body: string) {
  body = body.replace(/^.*?Critical above 240 min\./, '');
  const re = new RegExp(
    '(\\S+) (.+?) (' + DT + ') (\\d{2}:\\d{2}) (' + DT + ') (\\d{2}:\\d{2}) (\\d+) (.+?) (Low|High|Critical)(?= |$)',
    'g'
  );
  const rows = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(body))) {
    rows.push({
      vehicle: m[1],
      type: m[2],
      date: m[3],
      cardTime: m[4],
      gpsTime: m[6],
      mins: +m[7],
      location: m[8],
      risk: m[9],
    });
  }
  return rows;
}

const ABN_CATS = ['Multiple same-day fillings', 'Unusual filling variance'];
function parseAbnormal(body: string): ConsolidatedAbnormalRow[] {
  const pos: { i: number; c: string }[] = [];
  ABN_CATS.forEach((c) => {
    let i = -1;
    while ((i = body.indexOf(c + ' ', i + 1)) >= 0) pos.push({ i, c });
  });
  pos.sort((a, b) => a.i - b.i);
  const rows: ConsolidatedAbnormalRow[] = [];
  if (!pos.length) {
    const txt = body.replace(/Subtotal.*$/, '').trim();
    if (txt) rows.push({ cat: 'Other', vehicle: '-', issue: txt, litres: 0, action: '' });
    return rows;
  }
  pos.forEach((p, k) => {
    const chunk = body.slice(p.i + p.c.length + 1, k + 1 < pos.length ? pos[k + 1].i : body.length).trim();
    const m = chunk.match(/^(\S+) (.*) ([\d,]+\.\d{2}) ((?:(?![\d,]+\.\d{2}).)*)$/);
    if (m) {
      rows.push({
        cat: p.c,
        vehicle: m[1],
        issue: m[2].replace(/ on \d{4}-\d{2}-\d{2}:/, ':'),
        litres: numOf(m[3]),
        action: m[4].trim(),
      });
    } else {
      rows.push({ cat: p.c, vehicle: chunk.split(' ')[0] || '-', issue: chunk, litres: 0, action: '' });
    }
  });
  return rows;
}

function parseSummary(t: string) {
  const m = t.match(
    new RegExp(
      'Pumped Ltr Amount (' +
        DT +
        ') Diesel (\\S+) (\\S+) Petrol (\\S+) (\\S+) Super Diesel (\\S+) (\\S+)(?: Other (\\S+) (\\S+))? Total (\\S+) (\\S+)'
    )
  );
  if (!m) return null;
  return {
    date: m[1],
    diesel: numOf(m[2]),
    dieselAmt: numOf(m[3]),
    petrol: numOf(m[4]),
    petrolAmt: numOf(m[5]),
    superD: numOf(m[6]),
    superAmt: numOf(m[7]),
    other: numOf(m[8]),
    otherAmt: numOf(m[9]),
    total: numOf(m[10]),
    totalAmt: numOf(m[11]),
  };
}

export function parseDailyReport(rawText: string, fileName: string): ConsolidatedDay | { error: string; fileName: string } {
  let t = String(rawText || '').replace(/\s+/g, ' ');
  HDR_STRIP.forEach((re) => {
    t = t.replace(re, ' ');
  });
  t = t.replace(/\s+/g, ' ');
  const dm = t.match(new RegExp('DATE : (' + DT + ')'));
  const sum = parseSummary(t);
  const date = dm ? dm[1] : sum ? sum.date : null;
  if (!date) return { error: 'Not a daily reconciliation report (no "DATE :" line found)', fileName };
  const sec = splitSections(t);
  const day: ConsolidatedDay = {
    date,
    fileName,
    summary: sum,
    checks: [],
    warnings: [],
    unusual: [],
    flagged: [],
    ok: [],
    cardOnly: [],
    timeVar: [],
    abnormal: [],
    counts: { s5: sec.s5 ? sec.s5.count : null, s51: sec.s51 ? sec.s51.count : null, ok: sec.s6 ? sec.s6.count : null },
    grand: null,
    abnormalMissing: !sec.s7,
  };
  const chk = (label: string, got: number, want: number | null) =>
    day.checks.push({ label, got, want, ok: want == null ? true : got === want });
  const rowsOf = (k: string) => (sec[k] ? parseMatched(sec[k].body) : []);
  day.unusual = rowsOf('s1');
  day.flagged = rowsOf('s2');
  day.ok = rowsOf('s6');
  const co = sec.s3 ? parseCardOnly(sec.s3.body) : { rows: [], subCounts: [] };
  day.cardOnly = co.rows;
  day.timeVar = sec.s4 ? parseTime(sec.s4.body) : [];
  day.abnormal = sec.s7 ? parseAbnormal(sec.s7.body) : [];

  if (sec.s1) chk('Unusual filling variance', day.unusual.length, sec.s1.count);
  else day.warnings.push('Section 1 not found');
  if (sec.s2) chk('More than ±1.5 L variance', day.flagged.length, sec.s2.count);
  else day.warnings.push('Section 2 not found');
  if (sec.s3) chk('GPS unavailable & offline', day.cardOnly.length, sec.s3.count);
  else day.warnings.push('Section 3 not found');
  if (sec.s4) chk('Time variance > 5 min', day.timeVar.length, sec.s4.count);
  else day.warnings.push('Section 4 not found');
  if (sec.s6) chk('Variance OK', day.ok.length, sec.s6.count);
  if (sec.s7) chk('Abnormal highlights', day.abnormal.length, sec.s7.count);
  else {
    const m7 = t.match(/7\. Abnormal highlights (\d+)/);
    day.warnings.push(
      'Section 7 (Abnormal highlights) is missing from this PDF' + (m7 ? ` — its summary lists ${m7[1]}` : '')
    );
  }
  const gt = t.match(/GRAND TOTAL ([\d,.]+) ([\d,.]+) ([+-][\d,.]+) ([+-][\d.]+)%/);
  day.grand = gt ? { cardL: numOf(gt[1]), gpsL: numOf(gt[2]), variance: numOf(gt[3]), pct: numOf(gt[4]) } : null;
  if (!day.summary) day.warnings.push('Summary table (fuel type totals) not found');
  return day;
}

/* Direct conversion from Module 1 live workspace into ConsolidatedDay(s) grouped by date */
export function convertReportToConsolidatedDays(
  cardRows: CardRecord[],
  report: GpsReportData,
  abnormalItems: AbnormalItem[]
): ConsolidatedDay[] {
  const dates = [...new Set(cardRows.map((c) => c.date).filter(Boolean))].sort();
  if (!dates.length) return [];

  const toMatchedRow = (r: MatchedEntry, isUnusual: boolean): ConsolidatedMatchedRow => ({
    date: r.card.date,
    cardTime: r.card.time || '00:00',
    gpsTimes: gpsDtList(r.gpsList, ', '),
    vehicle: r.card.vehicle,
    type: vtype(r.card),
    area: r.card.area || '-',
    cardL: r.cardL,
    gpsL: r.gpsL,
    variance: r.variance,
    pct: r.pct ?? 0,
    status: isUnusual ? 'UNUSUAL' : r.status === 'ok' ? 'OK' : 'Mismatch',
    note: `${r.gpsList.length} fill(s)${r.unusual?.note ? ' — ' + r.unusual.note : ''}`,
  });

  const timeRowsAll = buildTimeVarianceRows(report);

  return dates.map((d) => {
    const dayCards = cardRows.filter((c) => c.date === d);
    let diesel = 0,
      dieselAmt = 0,
      petrol = 0,
      petrolAmt = 0,
      superD = 0,
      superAmt = 0,
      other = 0,
      otherAmt = 0;

    dayCards.forEach((c) => {
      const nf = normFuel(c.fuel);
      const amt = +(c.amount || 0);
      if (nf === 'Diesel') {
        diesel += c.liters;
        dieselAmt += amt;
      } else if (nf === 'Petrol') {
        petrol += c.liters;
        petrolAmt += amt;
      } else if (nf === 'Super Diesel') {
        superD += c.liters;
        superAmt += amt;
      } else {
        other += c.liters;
        otherAmt += amt;
      }
    });

    const total = diesel + petrol + superD + other;
    const totalAmt = dieselAmt + petrolAmt + superAmt + otherAmt;

    const unusual = report.matchedUnusual.filter((r) => r.card.date === d).map((r) => toMatchedRow(r, true));
    const flagged = report.matchedFlagged.filter((r) => r.card.date === d).map((r) => toMatchedRow(r, false));
    const ok = report.matchedOk.filter((r) => r.card.date === d).map((r) => toMatchedRow(r, false));

    const cardOnly: ConsolidatedCardOnlyRow[] = report.cardOnly
      .filter((r) => r.card.date === d)
      .map((r) => ({
        date: d,
        cardTime: r.card.time || '',
        vehicle: r.card.vehicle,
        type: vtype(r.card),
        shift: r.card.shift || '-',
        area: r.card.area || '-',
        cardL: r.card.liters,
        note: r.detail || r.note,
        category: r.category || UNCLASSIFIED_LABEL,
      }));

    const timeVar = timeRowsAll
      .filter((x) => x.r.card.date === d)
      .map((x) => ({
        vehicle: x.r.card.vehicle,
        type: vtype(x.r.card),
        date: x.first.date,
        cardTime: x.r.card.time,
        gpsTime: x.first.time,
        mins: x.mins,
        location: x.first.location || '-',
        risk: timeRisk(x.mins),
      }));

    const dayVehicles = new Set(dayCards.map((c) => c.vehicle.toUpperCase()));
    const abnormal: ConsolidatedAbnormalRow[] = abnormalItems
      .filter((a) => dayVehicles.has((a.veh || '').toUpperCase()) || dates.length === 1)
      .map((a) => ({
        cat: a.cat || 'Other',
        vehicle: a.veh,
        issue: a.issue,
        litres: +a.litres || 0,
        action: a.action,
      }));

    return {
      date: d,
      fileName: 'Direct sync from GPS Reconciliation',
      summary: { date: d, diesel, dieselAmt, petrol, petrolAmt, superD, superAmt, other, otherAmt, total, totalAmt },
      checks: [
        { label: 'Unusual filling variance', got: unusual.length, want: unusual.length, ok: true },
        { label: `More than ±${VARIANCE_TOLERANCE_L} L variance`, got: flagged.length, want: flagged.length, ok: true },
        { label: 'GPS unavailable & offline', got: cardOnly.length, want: cardOnly.length, ok: true },
        { label: 'Time variance > 5 min', got: timeVar.length, want: timeVar.length, ok: true },
        { label: 'Variance OK', got: ok.length, want: ok.length, ok: true },
        { label: 'Abnormal highlights', got: abnormal.length, want: abnormal.length, ok: true },
      ],
      warnings: [],
      unusual,
      flagged,
      ok,
      cardOnly,
      timeVar,
      abnormal,
      counts: { s5: null, s51: null, ok: ok.length },
      grand: null,
      abnormalMissing: false,
    };
  });
}

export function deriveAbnormal(day: ConsolidatedDay, opts: { multiples?: boolean } = {}): ConsolidatedAbnormalRow[] {
  const norm = (v: string) => String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const all = ([] as { v: string; t: string; l: number }[])
    .concat(day.unusual.map((r) => ({ v: r.vehicle, t: r.cardTime, l: r.cardL })))
    .concat(day.flagged.map((r) => ({ v: r.vehicle, t: r.cardTime, l: r.cardL })))
    .concat(day.ok.map((r) => ({ v: r.vehicle, t: r.cardTime, l: r.cardL })))
    .concat(day.cardOnly.map((r) => ({ v: r.vehicle, t: r.cardTime, l: r.cardL })));
  const g: Record<string, { v: string; t: string; l: number }[]> = {};
  all.forEach((e) => {
    (g[norm(e.v)] = g[norm(e.v)] || []).push(e);
  });
  const out: ConsolidatedAbnormalRow[] = [];
  if (opts.multiples) {
    Object.values(g)
      .filter((a) => a.length > 1)
      .forEach((a) => {
        a.sort((x, y) => (x.t || '').localeCompare(y.t || ''));
        out.push({
          cat: 'Multiple same-day fillings',
          vehicle: a[0].v,
          derived: true,
          issue: `Billed ${a.length} times: ` + a.map((e) => `${+e.l}L${e.t ? ' @ ' + e.t : ''}`).join(', '),
          litres: a.reduce((s, e) => s + e.l, 0),
          action: 'Verify pump receipts / filling station terminal logs',
        });
      });
  }
  day.unusual.forEach((r) => {
    const n = (r.note.match(/(\d+) fill/) || [0, 1])[1];
    out.push({
      cat: 'Unusual filling variance',
      vehicle: r.vehicle,
      derived: true,
      issue: `Card ${r.cardL.toFixed(2)} L vs GPS ${r.gpsL.toFixed(2)} L (${n} fill(s))`,
      litres: r.cardL,
      action: 'Review remark and supporting documents',
    });
  });
  return out;
}

const normV = (v: string) => String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
const sumOf = <T,>(a: T[], f: (r: T) => number) => a.reduce((s, r) => s + f(r), 0);
export const f2 = (n: number) => (+n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
export const f2s = (n: number) => (n > 0 ? '+' : '') + f2(n);

export interface ConsolidatedModel {
  dates: string[];
  s1: { date: string; diesel: number; petrol: number; superD: number; other: number; total: number; amt: number; missing?: boolean }[];
  s2: { date: string; time: string; vehicle: string; type: string; cardL: number; gpsL: number; unusual: boolean }[];
  s3: { date: string; vehicle: string; dt: string; litres: number; category: string }[];
  s4: { date: string; vehicle: string; cardDT: string; gpsDT: string; mins: number; risk: string }[];
  s5: { date: string; vehicle: string; dt: string; variance: number | null; issue: string; action: string }[];
  notes: string[];
}

export function buildModel(days: Record<string, ConsolidatedDay>, opts: { multiples?: boolean } = {}): ConsolidatedModel {
  const ds = Object.keys(days).sort();
  const m: ConsolidatedModel = { dates: ds, s1: [], s2: [], s3: [], s4: [], s5: [], notes: [] };
  ds.forEach((d) => {
    const x = days[d];
    const s = x.summary;
    if (s) m.s1.push({ date: d, diesel: s.diesel, petrol: s.petrol, superD: s.superD, other: s.other || 0, total: s.total, amt: s.totalAmt });
    else m.s1.push({ date: d, diesel: 0, petrol: 0, superD: 0, other: 0, total: 0, amt: 0, missing: true });

    const mk = (r: ConsolidatedMatchedRow, unusual: boolean) => ({
      date: d,
      time: r.cardTime,
      vehicle: r.vehicle,
      type: r.type,
      cardL: r.cardL,
      gpsL: r.gpsL,
      unusual,
    });
    const two = x.unusual.map((r) => mk(r, true)).concat(x.flagged.map((r) => mk(r, false)));
    two.sort((a, b) => a.time.localeCompare(b.time));
    m.s2.push(...two);

    x.timeVar.forEach((r) =>
      m.s4.push({ date: d, vehicle: r.vehicle, cardDT: d + ' ' + r.cardTime, gpsDT: r.date + ' ' + r.gpsTime, mins: r.mins, risk: r.risk })
    );
    x.cardOnly.forEach((r) =>
      m.s3.push({ date: d, vehicle: r.vehicle, dt: d + (r.cardTime ? ' ' + r.cardTime : ''), litres: r.cardL, category: r.category })
    );

    const ab = x.abnormalMissing ? deriveAbnormal(x, { multiples: !!opts.multiples }) : x.abnormal;
    if (x.abnormalMissing) m.notes.push(d + ': section 7 missing in PDF — rebuilt from unusual entries');
    const used: Record<number, number> = {};
    ab.forEach((r) => {
      let hit: ConsolidatedMatchedRow | undefined;
      if (r.cat === 'Unusual filling variance') {
        hit = x.unusual.find((u, i) => !used[i] && normV(u.vehicle) === normV(r.vehicle) && Math.abs(u.cardL - r.litres) < 0.005);
        if (hit) used[x.unusual.indexOf(hit)] = 1;
      }
      const timeMatches = r.issue.match(/@ \d{2}:\d{2}/g);
      m.s5.push({
        date: d,
        vehicle: r.vehicle,
        dt: hit ? d + ' ' + hit.cardTime : d + (timeMatches && timeMatches.length ? ' ' + timeMatches.map((z) => z.slice(2)).join(', ') : ''),
        variance: hit ? +(hit.gpsL - hit.cardL).toFixed(2) : null,
        issue: (r.cat === 'Unusual filling variance' ? '' : r.cat + ': ') + r.issue,
        action: r.action || '',
      });
    });
  });
  m.s3.sort((a, b) => a.dt.localeCompare(b.dt) || a.vehicle.localeCompare(b.vehicle));
  return m;
}

export interface MatrixDef<T = any> {
  title: string;
  sub: string;
  unit: string;
  signed: boolean;
  mat: {
    dates: string[];
    keys: string[];
    names: string[];
    cells: Record<string, Record<string, number[]>>;
  };
  _phantom?: T;
}

function buildMatrix<T extends { vehicle: string; date: string }>(
  dates: string[],
  rows: T[],
  valFn: (r: T) => number
) {
  const vm: Record<string, string> = {};
  const cells: Record<string, Record<string, number[]>> = {};
  rows.forEach((r) => {
    const k = normV(r.vehicle);
    vm[k] = vm[k] || r.vehicle;
    const d = r.date;
    ((cells[d] = cells[d] || {})[k] = cells[d][k] || []).push(valFn(r));
  });
  const keys = Object.keys(vm).sort((a, b) => vm[a].localeCompare(vm[b]));
  return { dates, keys, names: keys.map((k) => vm[k]), cells };
}

export function buildMatrices(m: ConsolidatedModel): MatrixDef[] {
  return [
    {
      title: '6. Filling variance — Pumped (GPS) less than card (vehicles across, days down)',
      sub: 'Only fills where GPS litres are lower than the card pump. Cell = variance in L (GPS − card).',
      unit: 'L',
      signed: true,
      mat: buildMatrix(
        m.dates,
        m.s2.filter((x) => x.gpsL < x.cardL),
        (r) => +(r.gpsL - r.cardL).toFixed(2)
      ),
    },
    {
      title: '7. Fuel Pump Time Variance GPS & Pump Report (vehicles across, days down)',
      sub: 'Cell = minutes between card time and first GPS fill (more than 5 min).',
      unit: 'min',
      signed: false,
      mat: buildMatrix(m.dates, m.s4, (r) => r.mins),
    },
    {
      title: '8. GPS Unavailable & Offline Vehicle List (vehicles across, days down)',
      sub: 'Cell = litres pumped on the card with no GPS fill matched.',
      unit: 'L',
      signed: false,
      mat: buildMatrix(m.dates, m.s3, (r) => r.litres),
    },
  ];
}

function addMatrixSheet(wb: ExcelJS.Workbook, m: ConsolidatedModel) {
  const ws = wb.addWorksheet('Vehicle x Day');
  const thin = { style: 'thin' as const, color: { argb: 'FFBDB8AA' } };
  const B = { top: thin, left: thin, bottom: thin, right: thin };
  const HF = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFEFECE2' } };
  const TF = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFE9E5D8' } };
  const SF = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFDFE9E2' } };
  ws.getColumn(1).width = 14;
  let r = 1;
  let maxC = 2;
  ws.getCell(r, 1).value = 'Vehicles across · days down (variance only)';
  ws.getCell(r, 1).font = { bold: true, size: 13 };
  r += 2;
  buildMatrices(m).forEach((x) => {
    const M = x.mat;
    const nC = M.keys.length + 2;
    maxC = Math.max(maxC, nC);
    const t = ws.getCell(r, 1);
    t.value = x.title;
    t.font = { bold: true };
    for (let c = 1; c <= Math.max(nC, 8); c++) ws.getCell(r, c).fill = TF;
    r++;
    const s = ws.getCell(r, 1);
    s.value = x.sub;
    s.font = { italic: true, size: 9, color: { argb: 'FF7A8390' } };
    r++;
    if (!M.keys.length) {
      ws.getCell(r, 1).value = 'None.';
      r += 2;
      return;
    }
    ws.getCell(r, 1).value = 'Date';
    M.names.forEach((n, i) => (ws.getCell(r, 2 + i).value = n));
    ws.getCell(r, nC).value = 'Day total (all)';
    for (let c = 1; c <= nC; c++) {
      const k = ws.getCell(r, c);
      k.font = { bold: true };
      k.fill = HF;
      k.border = B;
      k.alignment = { horizontal: 'center', wrapText: true };
    }
    r++;
    const fmt = x.unit === 'min' ? '0' : '#,##0.00;-#,##0.00';
    const first = r;
    M.dates.forEach((d) => {
      ws.getCell(r, 1).value = new Date(d + 'T00:00:00Z');
      ws.getCell(r, 1).numFmt = 'yyyy-mm-dd';
      ws.getCell(r, 1).border = B;
      M.keys.forEach((k, i) => {
        const a = (M.cells[d] || {})[k];
        const c = ws.getCell(r, 2 + i);
        c.border = B;
        c.numFmt = fmt;
        if (a) {
          c.value = +sumOf(a, (v) => v).toFixed(2);
          if (a.length > 1) c.note = 'Fills: ' + a.map((v) => (x.signed ? f2s(v) : v)).join(' / ');
        }
      });
      const L1 = ws.getColumn(2).letter;
      const L2 = ws.getColumn(nC - 1).letter;
      const dt = ws.getCell(r, nC);
      dt.value = {
        formula: `SUM(${L1}${r}:${L2}${r})`,
        result: sumOf(M.keys, (k) => sumOf((M.cells[d] || {})[k] || [], (v) => v)),
      };
      dt.numFmt = fmt;
      dt.font = { bold: true };
      dt.border = B;
      r++;
    });
    ws.getCell(r, 1).value = 'Total';
    M.keys.forEach((k, i) => {
      const L = ws.getColumn(2 + i).letter;
      const c = ws.getCell(r, 2 + i);
      c.value = {
        formula: `SUM(${L}${first}:${L}${r - 1})`,
        result: sumOf(M.dates, (d) => sumOf((M.cells[d] || {})[k] || [], (v) => v)),
      };
      c.numFmt = fmt;
    });
    for (let c = 1; c <= nC; c++) {
      const k = ws.getCell(r, c);
      k.fill = SF;
      k.font = { bold: true };
      k.border = B;
    }
    r += 3;
  });
  for (let c = 2; c <= maxC; c++) ws.getColumn(c).width = 12;
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 0 }];
  ws.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
}

export async function buildConsolidatedWorkbook(m: ConsolidatedModel): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  [18, 18, 18, 18, 16, 14, 10, 34].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  const thin = { style: 'thin' as const, color: { argb: 'FFBDB8AA' } };
  const B = { top: thin, left: thin, bottom: thin, right: thin };
  const HF = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFEFECE2' } };
  const TF = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFE9E5D8' } };
  const UF = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFFBEEE0' } };
  const SF = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFDFE9E2' } };
  let r = 1;
  const d0 = m.dates[0];
  const d1 = m.dates[m.dates.length - 1];
  ws.getCell(r, 1).value = `CLEANTECH (PVT) LTD — Fuel Card (Mobitel OneFuel) vs GPS Reconciliation · ${d0}${d0 !== d1 ? ' → ' + d1 : ''}`;
  ws.getCell(r, 1).font = { bold: true, size: 13 };
  r += 1;

  const title = (t: string) => {
    const c = ws.getCell(r, 1);
    c.value = t;
    c.font = { bold: true, size: 11 };
    for (let i = 1; i <= 8; i++) ws.getCell(r, i).fill = TF;
    r++;
  };
  const hdr = (arr: string[], row?: number) =>
    arr.forEach((t, i) => {
      const c = ws.getCell(row || r, i + 1);
      c.value = t;
      c.font = { bold: true };
      c.fill = HF;
      c.border = B;
      c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    });
  const box = (row: number, n: number, fill?: ExcelJS.Fill | null) => {
    for (let i = 1; i <= n; i++) {
      const c = ws.getCell(row, i);
      c.border = B;
      if (fill) c.fill = fill;
    }
  };
  const NUM = '#,##0.00';
  const PCT = '0.0%';

  // 1
  title('1. Mobitel Bill Summary');
  ws.mergeCells(r, 2, r, 4);
  ws.mergeCells(r, 5, r + 1, 5);
  ws.mergeCells(r, 6, r + 1, 6);
  hdr(['', 'Fuel Type', '', '', 'Total Ltr', 'Total Amount'], r);
  hdr(['Date', 'Diesel', 'Petrol', 'Super Diesel'], r + 1);
  ws.getCell(r, 1).border = B;
  ws.getCell(r, 1).fill = HF;
  r += 2;
  const s1a = r;
  m.s1.forEach((x) => {
    ws.getCell(r, 1).value = new Date(x.date + 'T00:00:00Z');
    ws.getCell(r, 1).numFmt = 'yyyy-mm-dd';
    [x.diesel, x.petrol, x.superD].forEach((v, i) => (ws.getCell(r, 2 + i).value = v));
    ws.getCell(r, 5).value = x.other ? x.total : { formula: `SUM(B${r}:D${r})`, result: x.total };
    ws.getCell(r, 6).value = x.amt;
    for (let c = 2; c <= 6; c++) ws.getCell(r, c).numFmt = NUM;
    box(r, 6);
    r++;
  });
  ws.getCell(r, 1).value = 'Total';
  const colL = ['', '', 'B', 'C', 'D', 'E', 'F'];
  const keys = [null, null, 'diesel', 'petrol', 'superD', 'total', 'amt'] as const;
  for (let c = 2; c <= 6; c++) {
    const k = keys[c]!;
    const cell = ws.getCell(r, c);
    cell.value = { formula: `SUM(${colL[c]}${s1a}:${colL[c]}${r - 1})`, result: sumOf(m.s1, (x) => x[k]) };
    cell.numFmt = NUM;
  }
  box(r, 6, SF);
  for (let c = 1; c <= 6; c++) ws.getCell(r, c).font = { bold: true };
  r += 2;

  // 2
  title('2. filling variance +1.5');
  hdr(['Date ', 'Vehicle ', 'Vehicle Type ', 'Pumped (Mobitel) ', 'Pumped (GPS) ', 'Variance ', 'Vari % ', 'Op. Comment']);
  r++;
  const s2a = r;
  m.s2.forEach((x) => {
    ws.getCell(r, 1).value = x.date + ' ' + x.time;
    ws.getCell(r, 2).value = x.vehicle;
    ws.getCell(r, 3).value = x.type;
    ws.getCell(r, 4).value = x.cardL;
    ws.getCell(r, 5).value = x.gpsL;
    ws.getCell(r, 6).value = { formula: `E${r}-D${r}`, result: +(x.gpsL - x.cardL).toFixed(2) };
    ws.getCell(r, 7).value = { formula: `IF(D${r}=0,0,F${r}/D${r})`, result: x.cardL ? (x.gpsL - x.cardL) / x.cardL : 0 };
    [4, 5, 6].forEach((c) => (ws.getCell(r, c).numFmt = NUM));
    ws.getCell(r, 7).numFmt = PCT;
    box(r, 8, x.unusual ? UF : null);
    r++;
  });
  if (m.s2.length) {
    ws.getCell(r, 1).value = 'Total';
    const c2 = sumOf(m.s2, (x) => x.cardL);
    const g2 = sumOf(m.s2, (x) => x.gpsL);
    ws.getCell(r, 4).value = { formula: `SUM(D${s2a}:D${r - 1})`, result: c2 };
    ws.getCell(r, 5).value = { formula: `SUM(E${s2a}:E${r - 1})`, result: g2 };
    ws.getCell(r, 6).value = { formula: `E${r}-D${r}`, result: g2 - c2 };
    ws.getCell(r, 7).value = { formula: `IF(D${r}=0,0,F${r}/D${r})`, result: c2 ? (g2 - c2) / c2 : 0 };
    [4, 5, 6].forEach((c) => (ws.getCell(r, c).numFmt = NUM));
    ws.getCell(r, 7).numFmt = PCT;
    box(r, 8, SF);
    for (let c = 1; c <= 8; c++) ws.getCell(r, c).font = { bold: true };
    r++;
  }
  ws.getCell(r, 1).value = 'Shaded rows = flagged UNUSUAL filling on the daily report.';
  ws.getCell(r, 1).font = { italic: true, size: 9, color: { argb: 'FF7A8390' } };
  r += 2;

  // 4
  title('4. Fuel Pump Time Variance GPS & Pump Report — more than 5 min');
  hdr(['Vehicle ', 'Card Time ', 'GPS Time ', 'Variance ', 'Risk']);
  r++;
  m.s4.forEach((x) => {
    ws.getCell(r, 1).value = x.vehicle;
    ws.getCell(r, 2).value = x.cardDT;
    ws.getCell(r, 3).value = x.gpsDT;
    ws.getCell(r, 4).value = x.mins;
    ws.getCell(r, 5).value = x.risk;
    ws.getCell(r, 4).numFmt = '0" min"';
    box(
      r,
      5,
      x.risk === 'Critical'
        ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBE9E6' } }
        : x.risk === 'High'
          ? UF
          : null
    );
    r++;
  });
  r += 1;

  // 3
  title('3. GPS Unavailable & Offline Vehicle List');
  hdr(['Vehicle ', 'Date & Time', 'Pumped (Mobitel)', 'Category']);
  r++;
  const s3a = r;
  m.s3.forEach((x) => {
    ws.getCell(r, 1).value = x.vehicle;
    ws.getCell(r, 2).value = x.dt;
    ws.getCell(r, 3).value = x.litres;
    ws.getCell(r, 3).numFmt = NUM;
    ws.getCell(r, 4).value = x.category;
    box(r, 4);
    r++;
  });
  if (m.s3.length) {
    ws.getCell(r, 1).value = 'Total';
    ws.getCell(r, 3).value = { formula: `SUM(C${s3a}:C${r - 1})`, result: sumOf(m.s3, (x) => x.litres) };
    ws.getCell(r, 3).numFmt = NUM;
    box(r, 4, SF);
    for (let c = 1; c <= 4; c++) ws.getCell(r, c).font = { bold: true };
    r++;
  }
  r += 1;

  // 5
  title('5. Abnormal Highlights & Recommendations');
  hdr(['Vehicle ', 'Date & Time', 'Variance ', 'Op. Comment', 'Observed issue / Action']);
  ws.mergeCells(r, 5, r, 8);
  r++;
  m.s5.forEach((x) => {
    ws.getCell(r, 1).value = x.vehicle;
    ws.getCell(r, 2).value = x.dt;
    if (x.variance != null) {
      ws.getCell(r, 3).value = x.variance;
      ws.getCell(r, 3).numFmt = NUM;
    }
    ws.mergeCells(r, 5, r, 8);
    ws.getCell(r, 5).value = x.issue + (x.action ? ' — ' + x.action : '');
    ws.getCell(r, 5).alignment = { wrapText: true, vertical: 'top' };
    box(r, 8);
    r++;
  });

  addMatrixSheet(wb, m);
  ws.views = [{ showGridLines: true }];
  ws.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  return wb;
}
