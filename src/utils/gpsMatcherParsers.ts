import * as XLSX from 'xlsx';
import type {
  AbnormalItem,
  CardOnlyEntry,
  CardRecord,
  GpsFillRecord,
  GpsReportData,
  MatchedEntry,
  OfflineReason,
  TimeVarianceRow,
  UnusualFlag,
} from '../types/fuel';
import type { PdfPosItem } from './pdfReader';

export const EXPECTED_LOCATION = 'Sedawatta Filling Station';
export const VARIANCE_TOLERANCE_L = 1.5;
export const UNUSUAL_VARIANCE_L = 5;
export const TIME_MIN_SHOW = 5;
export const OFFLINE_CATEGORIES = ['Offline', 'No GPS Tracker', 'Not filled', 'Staff fuel', 'Other'];
export const UNCLASSIFIED_LABEL = 'Not classified';

export function esc(s: unknown): string {
  return (s ?? '')
    .toString()
    .replace(/[&<>"']/g, (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m] || m));
}

export function normVeh(v: unknown): string {
  return String(v || '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

export function isOnSite(g: GpsFillRecord): boolean {
  return (g.location || '').toLowerCase().includes(EXPECTED_LOCATION.toLowerCase());
}

export function addDays(iso: string, n: number): string {
  if (!iso) return '';
  const d = new Date(iso + 'T00:00:00Z');
  if (isNaN(d.getTime())) return '';
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

export function dateDiffDays(a: string, b: string): number {
  const da = new Date(a + 'T00:00:00Z');
  const db = new Date(b + 'T00:00:00Z');
  return Math.round((da.getTime() - db.getTime()) / 86400000);
}

export function shiftPredictedDates(card: CardRecord): string[] {
  const s = String(card.shift || '')
    .toUpperCase()
    .trim();
  const d = card.date;
  if (!d) return [];
  if (s === 'N' || s === 'D') return [d, addDays(d, 1)];
  return [d, addDays(d, 1), addDays(d, -1)];
}

export function fmtDT(date: string, time: string): string {
  return (date || '-') + (time && time !== '00:00' ? ' ' + time : '');
}

export function vtype(c?: CardRecord | null): string {
  return c && c.vehicleType ? c.vehicleType : '-';
}

export function gpsDtList(list: GpsFillRecord[], sep = ', '): string {
  return list
    .map((g) => g.date + ' ' + g.time)
    .sort()
    .join(sep);
}

export function autoNoGpsTracker(c?: CardRecord | null): boolean {
  return !!c && /tractor|staff|common/i.test([c.vehicle, c.vehicleType, c.area, c.empName].join(' '));
}

export function getUnusual(
  c: CardRecord,
  variance: number | null,
  unusualMap: Record<string, UnusualFlag>
): UnusualFlag | null {
  const m = unusualMap[c.id];
  if (m) return m.off ? null : m;
  if (variance != null && Math.abs(variance) > UNUSUAL_VARIANCE_L) {
    return {
      note: `Auto-flagged: variance ${variance > 0 ? '+' : ''}${(+variance).toFixed(2)} L (more than ±${UNUSUAL_VARIANCE_L} L)`,
      auto: true,
    };
  }
  return null;
}

export function getReason(
  cardId: string,
  cardRows: CardRecord[],
  offlineReasons: Record<string, OfflineReason>
): { category: string; detail: string; auto: boolean } {
  const r = offlineReasons[cardId];
  if (r && r.category) return { category: r.category, detail: r.detail || '', auto: false };
  const c = cardRows.find((x) => x.id === cardId);
  if (autoNoGpsTracker(c)) return { category: 'No GPS Tracker', detail: (r && r.detail) || '', auto: true };
  return { category: '', detail: (r && r.detail) || '', auto: false };
}

function extractDate(v: unknown): string {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
  }
  if (typeof v === 'number') {
    try {
      const d = XLSX.SSF.parse_date_code(v);
      if (d && d.y) return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
    } catch {
      // ignore
    }
  }
  const s = String(v ?? '').trim();
  const m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[0] : s;
}

export function parseCardWorkbook(workbook: XLSX.WorkBook, sourceFile: string): Omit<CardRecord, 'id' | 'idx'>[] {
  const out: Omit<CardRecord, 'id' | 'idx'>[] = [];
  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1, raw: false, defval: '' });
    if (!raw.length) return;

    let headerIdx = -1;
    for (let i = 0; i < Math.min(raw.length, 8); i++) {
      const j = (raw[i] || []).map((c) => String(c || '').toUpperCase()).join(' ');
      if (j.includes('DEPOT') && (j.includes('VEHICLE') || j.includes('LITERS') || j.includes('LITRES'))) {
        headerIdx = i;
        break;
      }
    }
    if (headerIdx < 0) return;

    let lastDepot = '';
    let lastDate = '';
    for (let i = headerIdx + 1; i < raw.length; i++) {
      const row = raw[i] || [];
      const depotRaw = String(row[0] ?? '').trim();
      if (depotRaw.startsWith('=') || /SUBTOTAL/i.test(depotRaw)) continue;
      if (depotRaw) lastDepot = depotRaw;

      const vehicle = String(row[3] ?? '').trim();
      if (!vehicle || vehicle.toUpperCase() === 'VEHICLE NUMBER') continue;

      const litersRaw = row[6];
      const liters =
        typeof litersRaw === 'number'
          ? litersRaw
          : parseFloat(String(litersRaw).replace(/[^\d.-]/g, '')) || 0;
      if (!isFinite(liters) || liters === 0) continue;

      const dateStr = extractDate(row[5]);
      if (dateStr) lastDate = dateStr;

      out.push({
        depot: lastDepot,
        no: row[1] ?? '',
        vehicle,
        vehicleKey: normVeh(vehicle),
        fuel: String(row[4] ?? '').trim(),
        liters,
        shift: String(row[7] ?? '').trim(),
        area: String(row[14] ?? '').trim(),
        voucher: String(row[13] ?? '').trim(),
        date: dateStr || lastDate,
        time: '00:00',
        sourceFile,
        sheetName,
      });
    }
  });
  return out;
}

export function parseMobitelItems(
  pages: PdfPosItem[][],
  fullText: string,
  sourceFile: string
): { rows: Omit<CardRecord, 'id' | 'idx'>[]; all: Omit<CardRecord, 'id' | 'idx'>[]; warnings: string[]; footerOk: boolean } {
  const warnings: string[] = [];
  const none = { rows: [], all: [], warnings, footerOk: false };
  const isSeq = (s: string) => /^\d{9}$/.test(s);
  const seqAll: PdfPosItem[] = [];
  pages.forEach((pg) => pg.forEach((it) => { if (isSeq(it.s)) seqAll.push(it); }));
  if (!seqAll.length || !pages.length) return none;

  const spread = (a: number[]) => Math.max(...a) - Math.min(...a);
  const rotated = spread(seqAll.map((i) => i.y)) < spread(seqAll.map((i) => i.x));
  const U = (it: PdfPosItem) => (rotated ? it.x : -it.y);
  const C = (it: PdfPosItem) => (rotated ? it.y + it.w / 2 : it.x + it.w / 2);

  const HEAD =
    /^(Transaction|Status|DateTime|Mileage|\(km\)|Payable|Amt \(LKR\)|Balance|Amount|Litres|\(LKR\)|Ex Amt|Excess|Fill up Amt|Fill Up|Max|Max Limit|Limit\(LKR|\(Litres\)|PPL|Fuel|Type|Vehicle|No|Station|Cost Center Id 2|Cost Center Id|Office|Division|Card No|Emp Name|Emp Id|Id|Device|Seq\. No)$/;
  const seq0 = pages[0].filter((it) => isSeq(it.s));
  if (!seq0.length) return none;
  const firstU0 = Math.min(...seq0.map(U));
  const hdr = pages[0].filter((it) => HEAD.test(it.s) && U(it) < firstU0 - 1.5).sort((a, b) => C(a) - C(b));

  const groups: { c: number; items: PdfPosItem[] }[] = [];
  hdr.forEach((it) => {
    const g = groups[groups.length - 1];
    if (g && Math.abs(C(it) - g.c) <= 4) {
      g.items.push(it);
      g.c = g.items.reduce((a, i) => a + C(i), 0) / g.items.length;
    } else {
      groups.push({ c: C(it), items: [it] });
    }
  });

  const KEY: Record<string, string> = {
    'device seq. no': 'seq',
    'device id': 'devid',
    'emp id': 'emp',
    'emp name': 'emp',
    'card no': 'card',
    division: 'division',
    office: 'office',
    'cost center id': 'cc',
    'cost center id 2': 'cc2',
    station: 'station',
    'vehicle no': 'vehicle',
    'fuel type': 'fuel',
    ppl: 'ppl',
    'max limit (litres)': 'maxL',
    'max limit(lkr': 'maxA',
    'fill up litres': 'fillL',
    'fill up amt (lkr)': 'fillA',
    'excess litres': 'exL',
    'ex amt (lkr)': 'exA',
    'balance litres': 'balL',
    'balance amount': 'balA',
    'payable amt (lkr)': 'payable',
    'mileage (km)': 'mileage',
    datetime: 'dt',
    'transaction status': 'status',
  };

  const cols = groups
    .map((g) => ({
      c: g.c,
      key: KEY[g.items.sort((a, b) => U(a) - U(b)).map((i) => i.s).join(' ').toLowerCase().replace(/\s+/g, ' ')],
    }))
    .filter((c) => c.key);

  const need = ['seq', 'vehicle', 'fuel', 'fillL', 'dt', 'status'];
  if (need.some((k) => !cols.some((c) => c.key === k))) return none;
  const seqCol = cols.find((c) => c.key === 'seq')!;
  const colOf = (it: PdfPosItem) => {
    let b = cols[0];
    let bd = 1e9;
    cols.forEach((c) => {
      const d = Math.abs(C(it) - c.c);
      if (d < bd) {
        bd = d;
        b = c;
      }
    });
    return b.key;
  };

  const recs: { seq: string; main: PdfPosItem[]; cont: PdfPosItem[] }[] = [];
  let prev: { seq: string; main: PdfPosItem[]; cont: PdfPosItem[] } | null = null;
  pages.forEach((pg, pi) => {
    const anchors = pg.filter((it) => isSeq(it.s) && Math.abs(C(it) - seqCol.c) < 12).sort((a, b) => U(a) - U(b));
    const pageRecs = anchors.map((a) => ({ seq: a.s, main: [] as PdfPosItem[], cont: [] as PdfPosItem[] }));
    const firstU = anchors.length ? U(anchors[0]) : Infinity;
    pg.forEach((it) => {
      if (pi === 0 && U(it) < firstU - 1.5) return;
      let ri = -1;
      for (let k = anchors.length - 1; k >= 0; k--) {
        if (U(it) >= U(anchors[k]) - 1.5) {
          ri = k;
          break;
        }
      }
      if (ri < 0) {
        if (prev) prev.cont.push(it);
        return;
      }
      if (ri === anchors.length - 1 && U(it) > U(anchors[ri]) + 38) return;
      pageRecs[ri].main.push(it);
    });
    pageRecs.forEach((r) => recs.push(r));
    if (pageRecs.length) prev = pageRecs[pageRecs.length - 1];
  });

  const n = (v: unknown) => parseFloat(String(v).replace(/,/g, ''));
  const all = recs
    .map((r) => {
      const S = (it: PdfPosItem) => (rotated ? it.y : it.x);
      const cellMain: Record<string, PdfPosItem[]> = {};
      const cellCont: Record<string, PdfPosItem[]> = {};
      const put = (list: PdfPosItem[], store: Record<string, PdfPosItem[]>) =>
        list.forEach((it) => {
          const k = colOf(it);
          (store[k] = store[k] || []).push(it);
        });
      put(r.main, cellMain);
      put(r.cont, cellCont);

      const ordered = (k: string) => {
        const m = (cellMain[k] || []).slice();
        if (k === 'emp') {
          m.sort((a, b) => S(a) - S(b));
          let g = 0;
          const gi = new Map<PdfPosItem, number>();
          m.forEach((it, i) => {
            if (i && S(it) - S(m[i - 1]) > 15) g++;
            gi.set(it, g);
          });
          m.sort((a, b) => (gi.get(a) || 0) - (gi.get(b) || 0) || U(a) - U(b));
        } else {
          m.sort((a, b) => U(a) - U(b) || C(a) - C(b));
        }
        return m.concat((cellCont[k] || []).slice().sort((a, b) => U(a) - U(b)));
      };

      const glue = (list: PdfPosItem[]) =>
        list.reduce(
          (o, it) =>
            o && ((/^[a-z]/.test(it.s) && /[A-Za-z]$/.test(o)) || /[A-Za-z0-9]-$/.test(o))
              ? o + it.s
              : (o ? o + ' ' : '') + it.s,
          ''
        );

      const txt = (k: string) =>
        ['station', 'emp'].includes(k)
          ? glue(ordered(k)).replace(/\s+/g, ' ').trim()
          : ordered(k)
              .map((i) => i.s)
              .join(' ')
              .replace(/\s+/g, ' ')
              .trim();

      const one = (k: string) => {
        const l = ordered(k);
        return l.length ? l[0].s : '';
      };

      const veh = txt('vehicle').replace(/(?<=[A-Za-z0-9])-\s+(?=\d)/g, '-');
      const dtm = txt('dt')
        .replace(/(\d{2}\/\d{2}\/\d{3}) (\d)(?= \d{1,2}:\d{2})/, '$1$2')
        .match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{1,2}):(\d{2})\s*([AP]M)/);
      let hh = dtm ? parseInt(dtm[4], 10) : 0;
      if (dtm) {
        if (dtm[6] === 'AM' && hh === 12) hh = 0;
        else if (dtm[6] === 'PM' && hh < 12) hh += 12;
      }

      let emp = txt('emp');
      const vk = normVeh(veh);
      if (vk) {
        const pat = new RegExp(
          '^' +
            vk
              .split('')
              .map((ch) => ch.replace(/[^A-Z0-9]/g, ''))
              .filter(Boolean)
              .join('[\\s-]*') +
            '\\b[\\s-]*',
          'i'
        );
        const m2 = emp.match(pat);
        if (m2) emp = emp.slice(m2[0].length).trim();
      }
      if (emp === txt('emp')) emp = emp.replace(/^[A-Z0-9]{1,4}-\s?\d{2,6}\s+/, '').trim();
      const station = txt('station')
        .replace(/\b([A-Za-z]{4,}) ([a-z]{1,2})\b(?= [A-Z])/g, '$1$2')
        .trim();
      const office = txt('office');

      return {
        depot: office,
        no: r.seq,
        vehicle: veh,
        vehicleKey: normVeh(veh),
        fuel: txt('fuel'),
        liters: n(one('fillL')),
        shift: '',
        area: txt('cc2') || '-',
        voucher: r.seq,
        date: dtm ? `${dtm[3]}-${dtm[2]}-${dtm[1]}` : '',
        time: dtm ? `${String(hh).padStart(2, '0')}:${dtm[5]}` : '00:00',
        station,
        status: one('status'),
        mileage: one('mileage') !== '' && !isNaN(n(one('mileage'))) ? n(one('mileage')) : null,
        amount: n(one('fillA')) || 0,
        payable: n(one('payable')) || 0,
        ppl: n(one('ppl')) || 0,
        cardNo: txt('card'),
        empName: txt('emp'),
        vehicleType: emp,
        division: txt('division'),
        sourceFile,
        sheetName: 'Mobitel OneFuel',
      };
    })
    .filter((r) => r.vehicle && isFinite(r.liters));

  const rows = all.filter((r) => String(r.status).toUpperCase() === 'ACT');
  const skippedRows = all.filter((r) => String(r.status).toUpperCase() !== 'ACT');
  if (skippedRows.length) {
    warnings.push(
      `${skippedRows.length} non-ACT transaction(s) ignored (${skippedRows.map((r) => r.vehicle + ' ' + r.status).join(', ')})`
    );
  }
  const t = String(fullText || '').replace(/\s+/g, ' ');
  const fn = t.match(/Total Number of Transactions\s*:\s*(\d+)/);
  const fl = t.match(/Total Fill Up Litres\s*:\s*([\d,]+\.\d+)/);
  const sumL = all.reduce((a, r) => a + r.liters, 0);
  let ok = true;
  if (fn && +fn[1] !== all.length) {
    ok = false;
    warnings.push(`report footer says ${fn[1]} transactions but ${all.length} were read — check the PDF`);
  }
  if (fl && Math.abs(n(fl[1]) - sumL) > 0.05) {
    ok = false;
    warnings.push(`report footer says ${n(fl[1]).toFixed(3)} L but ${sumL.toFixed(3)} L were read — check the PDF`);
  }
  return { rows, all, warnings, footerOk: ok };
}

export function parseMobitelText(
  text: string,
  sourceFile: string
): { rows: Omit<CardRecord, 'id' | 'idx'>[]; warnings: string[] } {
  const warnings: string[] = [];
  let t = String(text || '').replace(/\s+/g, ' ').trim();
  t = t.replace(/(\d{2}\/\d{2}\/\d{3}) (\d)(?= \d{1,2}:\d{2})/g, '$1$2');
  t = t.replace(/\b([A-Z0-9]{1,4})-\s+(\d{2,6})\b/g, '$1-$2');
  t = t.replace(/(\d)(?=[A-Z]{1,3}-\d)/g, '$1 ').replace(/(\d)(?=Station\b)/g, '$1 ');
  const n = (v: unknown) => parseFloat(String(v).replace(/,/g, ''));
  const VEH = '[A-Z0-9]{1,4}-\\d{2,6}';
  const reHead = new RegExp(
    '(' +
      VEH +
      ")\\s+([A-Za-z][A-Za-z0-9 &'/()-]*?)\\s+(\\d{4} \\d{4} \\d{4} \\d{4})\\s+(" +
      VEH +
      ')\\s+([A-Za-z]+)\\s+' +
      '([\\d,]+\\.\\d+)\\s+([\\d,]+\\.\\d+)\\s+([\\d,]+\\.\\d+)\\s+([\\d,]+\\.\\d+)\\s+' +
      '(\\d{2})\\/(\\d{2})\\/(\\d{4})\\s+(\\d{1,2}):(\\d{2})\\s*([AP]M)\\s+(\\w+)',
    'g'
  );
  const heads: { m: RegExpExecArray; start: number; end: number }[] = [];
  let m: RegExpExecArray | null;
  while ((m = reHead.exec(t)) !== null) heads.push({ m, start: m.index, end: reHead.lastIndex });
  const eor = t.search(/\*\*\*\s*End of Report/);

  const all = heads.map((h, i) => {
    const hm = h.m;
    const tail = t.slice(h.end, i + 1 < heads.length ? heads[i + 1].start : eor >= 0 ? eor : t.length);
    const money = tail.match(/[\d,]+\.\d{2}(?!\d)/g) || [];
    const seqM = tail.match(/(?<![\d.,])(\d{8,})(?![\d.,])/);
    const afterSeq = seqM ? tail.slice(tail.indexOf(seqM[1]) + seqM[1].length) : '';
    const ints = afterSeq.split(' ').filter((x) => /^\d+$/.test(x));
    let station = (tail.match(/^\s*(.*?)\s*(?=[\d,]+\.\d{2}(?!\d)|(?<![\d.,])\d{8,}(?![\d.,])|$)/) || [, ''])[1]
      .replace(/\b([A-Za-z]{4,}) ([a-z]{1,2})\b(?= [A-Z])/g, '$1$2')
      .trim();
    if (/Filling$/.test(station)) station += ' Station';
    let hh = parseInt(hm[13], 10);
    const ap = hm[15].toUpperCase();
    if (ap === 'AM' && hh === 12) hh = 0;
    else if (ap === 'PM' && hh < 12) hh += 12;
    const mid = hm[2].trim().split(' ');
    const office = mid.length >= 3 ? mid[mid.length - 1] : '';
    const division = mid.length >= 3 ? mid[mid.length - 2] : '';
    const vehType = mid.length >= 3 ? mid.slice(0, mid.length - 2).join(' ') : hm[2].trim();
    const seq = seqM ? seqM[1] : hm[4] + '|' + hm[10] + hm[11] + hm[12] + hm[13] + hm[14];
    return {
      depot: office,
      no: seq,
      vehicle: hm[4].trim(),
      vehicleKey: normVeh(hm[4]),
      fuel: hm[5].trim(),
      liters: n(hm[7]),
      shift: '',
      area: '-',
      voucher: seq,
      date: `${hm[12]}-${hm[11]}-${hm[10]}`,
      time: `${String(hh).padStart(2, '0')}:${hm[14]}`,
      station,
      status: hm[16],
      mileage: ints.length > 1 ? n(ints[ints.length - 1]) : null,
      amount: money.length ? n(money[0]) : 0,
      payable: money.length ? n(money[money.length - 1]) : 0,
      ppl: n(hm[6]),
      cardNo: hm[3],
      empName: hm[1],
      vehicleType: vehType,
      division,
      sourceFile,
      sheetName: 'Mobitel OneFuel',
    };
  });

  const rows = all.filter((r) => String(r.status).toUpperCase() === 'ACT');
  const skippedRows = all.filter((r) => String(r.status).toUpperCase() !== 'ACT');
  if (skippedRows.length) {
    warnings.push(
      `${skippedRows.length} non-ACT transaction(s) ignored (${skippedRows.map((r) => r.vehicle + ' ' + r.status).join(', ')})`
    );
  }
  const fn = t.match(/Total Number of Transactions\s*:\s*(\d+)/);
  const fl = t.match(/Total Fill Up Litres\s*:\s*([\d,]+\.\d+)/);
  const sumL = all.reduce((a, r) => a + r.liters, 0);
  if (fn && +fn[1] !== all.length) {
    warnings.push(`report footer says ${fn[1]} transactions but ${all.length} were read — check the PDF`);
  }
  if (fl && Math.abs(n(fl[1]) - sumL) > 0.05) {
    warnings.push(`report footer says ${n(fl[1]).toFixed(3)} L but ${sumL.toFixed(3)} L were read — check the PDF`);
  }
  return { rows, warnings };
}

export function parseMobitelPdf(
  data: { pages: PdfPosItem[][]; text: string },
  sourceFile: string
): { rows: Omit<CardRecord, 'id' | 'idx'>[]; warnings: string[] } {
  const a = parseMobitelItems(data.pages || [], data.text, sourceFile);
  if (a.rows.length && a.footerOk) return { rows: a.rows, warnings: a.warnings };
  const b = parseMobitelText(data.text, sourceFile);
  const bOk = b.rows.length && !b.warnings.some((w) => /footer says/.test(w));
  if (bOk) return b;
  return a.rows.length ? { rows: a.rows, warnings: a.warnings } : b;
}

export function parseGpsText(text: string, sourceFile: string): Omit<GpsFillRecord, 'id' | 'idx'>[] {
  text = text
    .replace(/Page\s+\d+\s+of\s+\d+/g, ' ')
    .replace(
      /№\s+(\d+\.\d+)\s+Grouping\s+([\s\S]*?)\s+Time\s+(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})\s+Location\s+([\s\S]*?)\s+Filled\s+(\d+(?:\.\d+)?\s*l)/g,
      ' $1 $2 $3 $4 $5 $6 '
    );
  let t = text.replace(/([A-Za-z])(\d)/g, '$1 $2').replace(/(\d)([A-Za-z])/g, '$1 $2');
  t = t.replace(/\s+/g, ' ').trim();
  const re =
    /(\d+\.\d+)\s+([A-Z0-9]{1,4}-\d{2,6})\s*([\s\S]*?)\s*(\d{4}-\d{2}-\d{2})\s*(\d{2}:\d{2}:\d{2})\s*([\s\S]*?)\s+(\d+\.\d{1,2})\b/g;
  const rows: Omit<GpsFillRecord, 'id' | 'idx'>[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(t)) !== null) {
    const no = m[1];
    if (!no.includes('.')) continue;
    rows.push({
      no,
      vehicle: m[2].trim(),
      vehicleKey: normVeh(m[2]),
      description: m[3].trim().replace(/\s+/g, ' '),
      date: m[4],
      time: m[5].slice(0, 5),
      location: m[6].trim().replace(/\s+/g, ' ').replace(/[,\s]+$/, ''),
      liters: parseFloat(m[7]),
      sourceFile,
    });
  }
  return rows;
}

export function gpsFooterTotal(text: string): number | null {
  const m = String(text || '')
    .replace(/\s+/g, ' ')
    .match(/-{3,}\s*Total[\s\S]*?(\d[\d,]*\.\d{2})\s*l\b/);
  return m ? parseFloat(m[1].replace(/,/g, '')) : null;
}

export function vehicleHasOnSiteGps(vehicleKey: string, gpsRows: GpsFillRecord[]): boolean {
  return gpsRows.some((g) => g.vehicleKey === vehicleKey && isOnSite(g));
}

export function buildGpsReportData(
  cardRows: CardRecord[],
  gpsRows: GpsFillRecord[],
  allocations: Record<string, string[]>,
  offlineReasons: Record<string, OfflineReason>,
  skipped: Record<string, boolean>,
  unusualMap: Record<string, UnusualFlag>
): GpsReportData {
  const allocatedGpsIds = new Set<string>();
  Object.values(allocations).forEach((arr) => arr.forEach((id) => allocatedGpsIds.add(id)));

  const matched: MatchedEntry[] = [];
  const cardOnly: CardOnlyEntry[] = [];

  cardRows.forEach((c) => {
    const allocIds = allocations[c.id] || [];
    const allocs = allocIds.map((id) => gpsRows.find((g) => g.id === id)).filter((g): g is GpsFillRecord => Boolean(g));
    if (allocs.length) {
      const gpsTotal = allocs.reduce((s, g) => s + g.liters, 0);
      const variance = +(gpsTotal - c.liters).toFixed(2);
      const pct = c.liters ? +((variance / c.liters) * 100).toFixed(1) : null;
      const dt = allocs.map((g) => g.date + ' ' + g.time).sort()[0];
      const locations = [...new Set(allocs.map((g) => g.location))].join(' | ');
      const ok = Math.abs(variance) <= VARIANCE_TOLERANCE_L;
      matched.push({
        unusual: getUnusual(c, variance, unusualMap),
        card: c,
        gpsList: allocs,
        cardL: c.liters,
        gpsL: gpsTotal,
        variance,
        pct,
        dt,
        locations,
        status: ok ? 'ok' : 'mismatch',
      });
    } else {
      const hasGps = vehicleHasOnSiteGps(c.vehicleKey, gpsRows);
      const reason = getReason(c.id, cardRows, offlineReasons);
      cardOnly.push({
        card: c,
        hasGps,
        skipped: !!skipped[c.id],
        note: hasGps ? 'Not allocated' : 'Offline / No GPS tracker',
        category: reason.category,
        detail: reason.detail,
        auto: reason.auto,
      });
    }
  });

  const gpsOnlyOnSite: { gps: GpsFillRecord }[] = [];
  const gpsOnlyOffSite: { gps: GpsFillRecord }[] = [];
  gpsRows.forEach((g) => {
    if (allocatedGpsIds.has(g.id)) return;
    (isOnSite(g) ? gpsOnlyOnSite : gpsOnlyOffSite).push({ gps: g });
  });

  matched.sort((a, b) => a.dt.localeCompare(b.dt) || a.card.vehicle.localeCompare(b.card.vehicle));
  cardOnly.sort(
    (a, b) =>
      (a.card.date || '').localeCompare(b.card.date || '') ||
      a.card.vehicle.localeCompare(b.card.vehicle) ||
      a.card.idx - b.card.idx
  );
  gpsOnlyOnSite.sort((a, b) => (a.gps.date + ' ' + a.gps.time).localeCompare(b.gps.date + ' ' + b.gps.time));
  gpsOnlyOffSite.sort((a, b) => (a.gps.date + ' ' + a.gps.time).localeCompare(b.gps.date + ' ' + b.gps.time));

  const sumCard = (arr: MatchedEntry[]) => arr.reduce((s, r) => s + r.cardL, 0);
  const sumCardOnly = (arr: CardOnlyEntry[]) => arr.reduce((s, r) => s + r.card.liters, 0);
  const sumGps = (arr: MatchedEntry[]) => arr.reduce((s, r) => s + r.gpsL, 0);
  const sumGpsOnly = (arr: { gps: GpsFillRecord }[]) => arr.reduce((s, r) => s + r.gps.liters, 0);

  const grandCard = sumCard(matched) + sumCardOnly(cardOnly);
  const grandGps = sumGps(matched) + sumGpsOnly(gpsOnlyOnSite) + sumGpsOnly(gpsOnlyOffSite);
  const grandVar = +(grandGps - grandCard).toFixed(2);
  const grandPct = grandCard ? +((grandVar / grandCard) * 100).toFixed(1) : null;

  const matchedUnusual = matched.filter((r) => r.unusual);
  const matchedNormal = matched.filter((r) => !r.unusual);
  const matchedFlagged = matchedNormal
    .filter((r) => r.status === 'mismatch')
    .sort((a, b) => Math.abs(b.variance) - Math.abs(a.variance) || a.dt.localeCompare(b.dt));
  const matchedOk = matchedNormal.filter((r) => r.status === 'ok');

  return {
    matched,
    matchedNormal,
    matchedUnusual,
    matchedFlagged,
    matchedOk,
    cardOnly,
    gpsOnlyOnSite,
    gpsOnlyOffSite,
    grandCard,
    grandGps,
    grandVar,
    grandPct,
  };
}

export function groupCardOnlyByCategory(cardOnly: CardOnlyEntry[]): {
  order: string[];
  groups: Record<string, CardOnlyEntry[]>;
} {
  const order = OFFLINE_CATEGORIES.concat([UNCLASSIFIED_LABEL]);
  const groups: Record<string, CardOnlyEntry[]> = {};
  order.forEach((c) => (groups[c] = []));
  cardOnly.forEach((r) => {
    const cat = r.category && order.indexOf(r.category) >= 0 ? r.category : UNCLASSIFIED_LABEL;
    groups[cat].push(r);
  });
  return { order, groups };
}

export function timeRisk(mins: number): 'Low' | 'High' | 'Critical' {
  return mins <= 30 ? 'Low' : mins <= 240 ? 'High' : 'Critical';
}

function toMs(date: string, time: string): number | null {
  const m = String(date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const t = String(time || '').match(/^(\d{2}):(\d{2})/);
  if (!m || !t) return null;
  return Date.UTC(+m[1], +m[2] - 1, +m[3], +t[1], +t[2]);
}

export function buildTimeVarianceRows(d: GpsReportData): TimeVarianceRow[] {
  const rows: TimeVarianceRow[] = [];
  d.matched.forEach((r) => {
    const c = r.card;
    if (!c.time || c.time === '00:00') return;
    const first = r.gpsList.slice().sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time))[0];
    const a = toMs(c.date, c.time);
    const b = toMs(first.date, first.time);
    if (a === null || b === null) return;
    const mins = Math.round(Math.abs(a - b) / 60000);
    if (mins > TIME_MIN_SHOW) rows.push({ r, first, mins });
  });
  return rows.sort((x, y) => y.mins - x.mins);
}

export function buildAutoAbnormalRows(d: GpsReportData): Omit<AbnormalItem, 'id'>[] {
  const out: Omit<AbnormalItem, 'id'>[] = [];
  d.matchedUnusual.forEach((r) =>
    out.push({
      cat: 'Unusual filling variance',
      veh: r.card.vehicle,
      issue: `Card ${r.cardL.toFixed(2)} L vs GPS ${r.gpsL.toFixed(2)} L, variance ${r.variance > 0 ? '+' : ''}${r.variance.toFixed(2)} L (${r.gpsList.length} fill(s))${r.unusual && r.unusual.note && !r.unusual.auto ? ' — ' + r.unusual.note : ''}`,
      litres: r.cardL,
      action:
        r.unusual && r.unusual.auto
          ? `Variance above ±${UNUSUAL_VARIANCE_L} L — verify pump receipt, GPS fill and vehicle log`
          : 'Review remark and supporting documents',
    })
  );
  return out;
}

export function normFuel(f: string): 'Diesel' | 'Petrol' | 'Super Diesel' | 'Other' {
  const t = String(f || '').toLowerCase();
  if (/super/.test(t)) return 'Super Diesel';
  if (/diesel/.test(t)) return 'Diesel';
  if (/petrol|gasoline|octane|\b9[25]\b/.test(t)) return 'Petrol';
  return 'Other';
}

export function buildGpsPrintableHtml(
  cardRows: CardRecord[],
  d: GpsReportData,
  abnormalItems: AbnormalItem[]
): string {
  const dates = [...new Set(cardRows.map((r) => r.date).filter(Boolean))].sort();
  const dateLabel = dates.length === 1 ? dates[0] : dates.length ? `${dates[0]} → ${dates[dates.length - 1]}` : '-';
  const sumCard = (arr: MatchedEntry[]) => arr.reduce((s, r) => s + r.cardL, 0);
  const sumGps = (arr: MatchedEntry[]) => arr.reduce((s, r) => s + r.gpsL, 0);
  const sumGpsOnly = (arr: { gps: GpsFillRecord }[]) => arr.reduce((s, r) => s + r.gps.liters, 0);

  const tableHead =
    '<thead><tr><th>Card Date &amp; Time</th><th>GPS Date &amp; Time</th><th>Vehicle</th><th>Vehicle Type</th><th>Shift</th><th>Area</th><th>Location(s)</th><th class="num">Pumped (Card)</th><th class="num">Pumped (GPS)</th><th class="num">Variance</th><th class="num">Vari %</th><th>Status</th><th>Note</th></tr></thead><tbody>';

  const subtotalRow = (label: string, cardT: number | null, gpsT: number | null) => {
    const variance = cardT != null && gpsT != null ? +(gpsT - cardT).toFixed(2) : null;
    const pct = cardT && variance != null ? +((variance / cardT) * 100).toFixed(1) : null;
    return `<tr class="subtotal-row"><td colspan="7"><b>${esc(label)}</b></td><td class="num"><b>${cardT != null ? cardT.toFixed(2) : '-'}</b></td><td class="num"><b>${gpsT != null ? gpsT.toFixed(2) : '-'}</b></td><td class="num"><b>${variance != null ? (variance >= 0 ? '+' : '') + variance.toFixed(2) : '-'}</b></td><td class="num"><b>${pct != null ? (pct >= 0 ? '+' : '') + pct + '%' : '-'}</b></td><td colspan="2"></td></tr>`;
  };

  const matchedRowHtml = (r: MatchedEntry) => {
    const c = r.card;
    const statusCell = r.unusual ? 'UNUSUAL' : r.status === 'ok' ? 'OK' : 'Mismatch';
    const noteCell = r.unusual ? `${r.gpsList.length} fill(s)${r.unusual.note ? ' — ' + esc(r.unusual.note) : ''}` : r.gpsList.length;
    return `<tr>
      <td>${esc(fmtDT(c.date, c.time))}</td>
      <td>${gpsDtList(r.gpsList, '<br>')}</td>
      <td><b>${esc(c.vehicle)}</b></td>
      <td>${esc(vtype(c))}</td>
      <td>${esc(c.shift || '-')}</td>
      <td>${esc(c.area || '-')}</td>
      <td>${esc(r.locations || '-')}</td>
      <td class="num">${r.cardL.toFixed(2)}</td>
      <td class="num">${r.gpsL.toFixed(2)}</td>
      <td class="num">${(r.variance >= 0 ? '+' : '') + r.variance.toFixed(2)}</td>
      <td class="num">${r.pct != null ? (r.pct >= 0 ? '+' : '') + r.pct + '%' : '-'}</td>
      <td>${statusCell}</td>
      <td>${noteCell}</td>
    </tr>`;
  };

  const cardOnlyRowHtml = (r: CardOnlyEntry) => {
    const c = r.card;
    const note = r.detail
      ? r.detail
      : r.auto
        ? 'Auto: tractor / staff / common card — no GPS tracker'
        : r.hasGps
          ? r.skipped
            ? 'Skipped — no GPS fill to allocate'
            : 'GPS fills available — not allocated'
          : 'No GPS fills in PDF';
    return `<tr>
      <td>${esc(fmtDT(c.date, c.time))}</td>
      <td>-</td>
      <td><b>${esc(c.vehicle)}</b></td>
      <td>${esc(vtype(c))}</td>
      <td>${esc(c.shift || '-')}</td>
      <td>${esc(c.area || '-')}</td>
      <td>-</td>
      <td class="num">${c.liters.toFixed(2)}</td>
      <td class="num">-</td>
      <td class="num">-</td>
      <td class="num">-</td>
      <td>No GPS record</td>
      <td>${esc(note)}</td>
    </tr>`;
  };

  const gpsOnlyRowHtml = (r: { gps: GpsFillRecord }) => {
    const g = r.gps;
    const off = !isOnSite(g);
    const locCell = off ? `<span style="color:#a4342a;font-weight:700">⚠ ${esc(g.location)}</span>` : esc(g.location || '-');
    return `<tr>
      <td>-</td>
      <td>${esc(g.date + ' ' + g.time)}</td>
      <td><b>${esc(g.vehicle)}</b></td>
      <td>-</td>
      <td>-</td>
      <td>-</td>
      <td>${locCell}</td>
      <td class="num">-</td>
      <td class="num">${g.liters.toFixed(2)}</td>
      <td class="num">-</td>
      <td class="num">-</td>
      <td>No card record</td>
      <td style="font-size:11px;">${esc(g.description || '')}</td>
    </tr>`;
  };

  const buildSectionHtml = (title: string, subtitle: string, rowsHtml: string, subHtml: string) => {
    const body = rowsHtml.trim() ? rowsHtml : `<tr><td colspan="13" style="color:#94a3b8;font-size:12px;padding:12px;">None.</td></tr>`;
    return `<div class="report-section"><div class="sec-title">${esc(title)}</div>${subtitle ? `<div class="sec-sub">${esc(subtitle)}</div>` : ''}<table>${tableHead}${body}${subHtml}</tbody></table></div>`;
  };

  // Summary
  const fuels: ('Diesel' | 'Petrol' | 'Super Diesel' | 'Other')[] = ['Diesel', 'Petrol', 'Super Diesel', 'Other'];
  const agg: Record<string, { l: number; a: number }> = {};
  let hasAmt = false;
  cardRows.forEach((c) => {
    const k = (c.date || '-') + '|' + normFuel(c.fuel);
    const a = (agg[k] = agg[k] || { l: 0, a: 0 });
    a.l += c.liters;
    a.a += +(c.amount || 0);
    if (+(c.amount || 0)) hasAmt = true;
  });
  const hasOther = cardRows.some((c) => normFuel(c.fuel) === 'Other');
  const show = fuels.filter((f) => f !== 'Other' || hasOther);
  const L = (v: number) => (v ? v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '-');
  const A = (v: number) => (hasAmt && v ? v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '-');
  let sumRows = '';
  let tl = 0;
  let ta = 0;
  dates.forEach((dt) => {
    show.forEach((f, i) => {
      const a = agg[dt + '|' + f] || { l: 0, a: 0 };
      tl += a.l;
      ta += a.a;
      sumRows += `<tr>${i === 0 ? `<td rowspan="${show.length}" style="vertical-align:middle;text-align:center;font-weight:700">${esc(dt)}</td>` : ''}<td>${f}</td><td class="num">${L(a.l)}</td><td class="num">${A(a.a)}</td></tr>`;
    });
  });
  sumRows += `<tr class="subtotal-row"><td></td><td><b>Total</b></td><td class="num"><b>${L(tl)}</b></td><td class="num"><b>${A(ta)}</b></td></tr>`;

  const co = groupCardOnlyByCategory(d.cardOnly).groups;
  const timeRows = buildTimeVarianceRows(d);
  const abItems = abnormalItems.filter((x) => x.cat || x.veh || x.issue || x.action || x.litres);

  const cntRows = [
    ['1. Unusual filling variance', d.matchedUnusual.length],
    [`2. More than ±${VARIANCE_TOLERANCE_L} L variance`, d.matchedFlagged.length],
    ['3. GPS unavailable & offline', `${d.cardOnly.length} (Not classified ${(co[UNCLASSIFIED_LABEL] || []).length}, No GPS Tracker ${(co['No GPS Tracker'] || []).length})`],
    [`4. Time variance > ${TIME_MIN_SHOW} min`, timeRows.length],
    [`5. On GPS (${EXPECTED_LOCATION}), not on card`, d.gpsOnlyOnSite.length + (d.gpsOnlyOffSite.length ? ` (+ ${d.gpsOnlyOffSite.length} not at Sedawatta)` : '')],
    [`6. Variance within ±${VARIANCE_TOLERANCE_L} L — OK`, d.matchedOk.length],
    ['7. Abnormal highlights', abItems.length],
  ]
    .map((r) => `<tr><td>${esc(r[0])}</td><td class="num">${esc(String(r[1]))}</td></tr>`)
    .join('');

  let html = `<div class="report-heading">
    <div class="line1">CLEANTECH (PVT) LTD — FUEL RECONCILIATION REPORT</div>
    <div class="line2">Fuel card (Excel / Mobitel OneFuel) vs GPS fills</div>
    <div class="line3">DATE : ${esc(dateLabel)}</div>
  </div>`;

  html += `<div class="report-section"><div class="sec-title">Summary</div>
    <table class="simple" style="max-width:520px"><thead><tr><th colspan="4" style="text-align:left">(Mobitel Card)</th></tr><tr><th rowspan="2" style="text-align:center">Date</th><th rowspan="2" style="text-align:center">Fuel Type</th><th colspan="2" style="text-align:center">Pumped</th></tr><tr><th class="num">Pumped Ltr</th><th class="num">Amount</th></tr></thead><tbody>${sumRows}</tbody></table>
    <table class="simple" style="max-width:520px;margin-top:14px"><thead><tr><th>Report section</th><th class="num">Entries</th></tr></thead><tbody>${cntRows}</tbody></table>
  </div>`;

  html += buildSectionHtml(
    `1. Unusual filling variance (${d.matchedUnusual.length})`,
    `Entries flagged as unusual filling — automatically when |variance| is more than ${UNUSUAL_VARIANCE_L} L, or manually during allocation.`,
    d.matchedUnusual.map(matchedRowHtml).join(''),
    subtotalRow('Subtotal — Unusual filling variance', sumCard(d.matchedUnusual), sumGps(d.matchedUnusual))
  );

  html += buildSectionHtml(
    `2. More than ±${VARIANCE_TOLERANCE_L} L Variance — Fuel Pump vs GPS (${d.matchedFlagged.length})`,
    `On-site at ${EXPECTED_LOCATION}, sorted by largest variance. Variance above ±${UNUSUAL_VARIANCE_L} L is reported in section 1.`,
    d.matchedFlagged.map(matchedRowHtml).join(''),
    subtotalRow('Subtotal — Variance flagged', sumCard(d.matchedFlagged), sumGps(d.matchedFlagged))
  );

  const SEC3: [string, string][] = [
    [UNCLASSIFIED_LABEL, 'Not classified'],
    ['No GPS Tracker', 'No GPS Tracker — Tractors, Staff & Common'],
    ['Offline', 'Offline'],
    ['Not filled', 'Not filled'],
    ['Staff fuel', 'Staff fuel'],
    ['Other', 'Other'],
  ];
  let sec3Body = '';
  SEC3.forEach(([cat, label], i) => {
    const list = co[cat] || [];
    if (!list.length) return;
    sec3Body += `<div class="subsection"><div class="subsec-title">3.${i + 1} ${esc(label)} (${list.length})</div><table>${tableHead}${list.map(cardOnlyRowHtml).join('')}${subtotalRow(`Subtotal — ${label}`, list.reduce((s, r) => s + r.card.liters, 0), null)}</tbody></table></div>`;
  });
  if (!sec3Body) sec3Body = `<div class="subsection"><table>${tableHead}<tr><td colspan="13" style="color:#94a3b8;font-size:12px;padding:12px;">None.</td></tr></tbody></table></div>`;
  html += `<div class="report-section"><div class="sec-title">3. GPS Unavailable &amp; Offline Vehicle List (${d.cardOnly.length})</div><div class="sec-sub">Card entries not allocated to any Sedawatta GPS fill.</div>${sec3Body}</div>`;

  const timeBody =
    timeRows
      .map(
        (x) =>
          `<tr><td><b>${esc(x.r.card.vehicle)}</b></td><td>${esc(vtype(x.r.card))}</td><td>${esc(fmtDT(x.r.card.date, x.r.card.time))}</td><td>${esc(x.first.date + ' ' + x.first.time)}</td><td class="num">${x.mins}</td><td>${esc(x.first.location || '-')}</td><td>${timeRisk(x.mins)}</td></tr>`
      )
      .join('') || `<tr><td colspan="7" style="color:#94a3b8;font-size:12px;padding:12px;">None.</td></tr>`;
  html += `<div class="report-section"><div class="sec-title">4. Fuel Pump Time Variance GPS &amp; Pump Report — more than ${TIME_MIN_SHOW} min (${timeRows.length})</div><div class="sec-sub">Risk: Low ≤ 30 min, High ≤ 240 min, Critical above 240 min.</div><table class="simple"><thead><tr><th>Vehicle</th><th>Vehicle Type</th><th>Card Date &amp; Time</th><th>GPS Date &amp; Time</th><th class="num">Diff (Mins)</th><th>GPS Location</th><th>Risk</th></tr></thead><tbody>${timeBody}</tbody></table></div>`;

  html += `<div class="report-section"><div class="sec-title">5. On GPS (${esc(EXPECTED_LOCATION)}), not on card (${d.gpsOnlyOnSite.length})</div><div class="sec-sub">Sedawatta fills in the GPS PDF that were not allocated to any card entry.</div><div class="subsection"><table>${tableHead}${d.gpsOnlyOnSite.length ? d.gpsOnlyOnSite.map(gpsOnlyRowHtml).join('') : `<tr><td colspan="13" style="color:#94a3b8;font-size:12px;padding:12px;">None.</td></tr>`}${subtotalRow('Subtotal — On GPS Sedawatta, not on card', null, sumGpsOnly(d.gpsOnlyOnSite))}</tbody></table></div>${d.gpsOnlyOffSite.length ? `<div class="subsection"><div class="subsec-title">5.1 GPS fills not at ${esc(EXPECTED_LOCATION)}, not on card (${d.gpsOnlyOffSite.length})</div><table>${tableHead}${d.gpsOnlyOffSite.map(gpsOnlyRowHtml).join('')}${subtotalRow('Subtotal — Not at Sedawatta', null, sumGpsOnly(d.gpsOnlyOffSite))}</tbody></table></div>` : ''}</div>`;

  html += buildSectionHtml(
    `6. Variance within ±${VARIANCE_TOLERANCE_L} L — OK (${d.matchedOk.length})`,
    `Matched entries where |variance| ≤ ${VARIANCE_TOLERANCE_L} L.`,
    d.matchedOk.map(matchedRowHtml).join(''),
    subtotalRow('Subtotal — Variance OK', sumCard(d.matchedOk), sumGps(d.matchedOk))
  );

  const abBody =
    abItems
      .map(
        (x) =>
          `<tr><td>${esc(x.cat)}</td><td><b>${esc(x.veh)}</b></td><td>${esc(x.issue)}</td><td class="num">${(+x.litres || 0).toFixed(2)}</td><td>${esc(x.action)}</td></tr>`
      )
      .join('') || `<tr><td colspan="5" style="color:#94a3b8;font-size:12px;padding:12px;">None.</td></tr>`;
  html += `<div class="report-section"><div class="sec-title">7. Abnormal Highlights &amp; Recommendations (${abItems.length})</div><table class="simple"><thead><tr><th>Category</th><th>Vehicle</th><th>Observed Issue</th><th class="num">Fuel Impact (L)</th><th>Action</th></tr></thead><tbody>${abBody}</tbody></table></div>`;

  html += `<div class="report-section" style="page-break-after:auto"><table><tbody><tr class="grandtotal-row"><td colspan="7"><b>GRAND TOTAL</b></td><td class="num"><b>${d.grandCard.toFixed(2)}</b></td><td class="num"><b>${d.grandGps.toFixed(2)}</b></td><td class="num"><b>${(d.grandVar >= 0 ? '+' : '') + d.grandVar.toFixed(2)}</b></td><td class="num"><b>${d.grandPct != null ? (d.grandPct >= 0 ? '+' : '') + d.grandPct + '%' : '-'}</b></td><td colspan="2"></td></tr></tbody></table></div>`;

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Fuel Card vs GPS Reconciliation</title>
  <style>
    @page { size: A4 landscape; margin: 10mm 8mm; }
    body{font-family:Arial,Helvetica,sans-serif;color:#1b2430;padding:6px;margin:0;font-size:10pt;line-height:1.35}
    .report-heading{text-align:center;margin-bottom:14px}
    .report-heading .line1{font-size:16pt;font-weight:700;text-transform:uppercase;margin-bottom:3px;letter-spacing:.4px}
    .report-heading .line2{font-size:11pt;margin-bottom:3px;color:#333}
    .report-heading .line3{font-size:10pt;color:#333}
    .report-section{margin-bottom:12px;page-break-after:always;break-after:page}
    .report-section:last-child{page-break-after:auto;break-after:auto}
    .report-section .sec-title{font-size:12pt;font-weight:700;background:#e5e2d6;padding:7px 10px;border-left:5px solid #2c5f4f;margin:0 0 5px 0;border-radius:3px}
    .report-section .sec-sub{font-size:9.5pt;color:#555;margin:0 0 6px 4px;font-style:italic}
    .report-section .subsection{margin-bottom:10px}
    .report-section .subsec-title{font-size:10.5pt;font-weight:700;background:#efece2;padding:5px 9px;border-left:4px solid #a9631a;margin:0 0 4px 0;border-radius:2px;color:#7a3d0e}
    table{width:100%;border-collapse:collapse;font-size:8.5pt;margin-bottom:6px;table-layout:fixed}
    th,td{border:1px solid #999;padding:4px 5px;text-align:left;vertical-align:top;word-wrap:break-word;overflow-wrap:break-word;word-break:break-word;hyphens:auto}
    th{background:#eee;font-size:8pt;font-weight:700;overflow-wrap:normal;word-break:normal;hyphens:none}
    td.num{text-align:right;white-space:nowrap}
    th.num{text-align:right;white-space:normal;line-height:1.15}
    table th:nth-child(1), table td:nth-child(1){width:9%}
    table th:nth-child(2), table td:nth-child(2){width:9%}
    table th:nth-child(3), table td:nth-child(3){width:7%}
    table th:nth-child(4), table td:nth-child(4){width:8%}
    table th:nth-child(5), table td:nth-child(5){width:4%}
    table th:nth-child(6), table td:nth-child(6){width:7%}
    table th:nth-child(7), table td:nth-child(7){width:13%}
    table th:nth-child(8), table td:nth-child(8){width:6%}
    table th:nth-child(9), table td:nth-child(9){width:6%}
    table th:nth-child(10), table td:nth-child(10){width:6%}
    table th:nth-child(11), table td:nth-child(11){width:5%}
    table th:nth-child(12), table td:nth-child(12){width:7%}
    table th:nth-child(13), table td:nth-child(13){width:13%}
    table.simple th, table.simple td{width:auto !important}
    tr.subtotal-row td{background:#f0eee2;font-weight:700}
    tr.grandtotal-row td{background:#dce7e0;font-weight:700;font-size:9.5pt}
  </style></head><body>${html}
  <script>window.onload=function(){setTimeout(function(){window.focus();window.print();},300);};<\/script>
  </body></html>`;
}
