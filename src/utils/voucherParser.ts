import ExcelJS from 'exceljs';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  COMMON_USAGE_CATEGORIES_DIESEL,
  COMMON_USAGE_CATEGORIES_PETROL,
  COMMON_USAGE_VEHICLES_DIESEL,
  COMMON_USAGE_VEHICLES_PETROL,
} from '../data/defaultAreas';
import type { VoucherRow } from '../types/fuel';
import type { PdfViewportItem } from './pdfReader';

export const APPLIED_COLS: { key: keyof VoucherRow | '__serial'; label: string }[] = [
  { key: '__serial', label: 'Serial No' },
  { key: 'usage_id', label: 'Usage_id' },
  { key: 'vehicle', label: 'Vehicle Number' },
  { key: 'fuel_type', label: 'Fuel Type' },
  { key: 'usage_date', label: 'Usage Date' },
  { key: 'liters', label: 'Liters Used' },
  { key: 'shift', label: 'SHIFT' },
  { key: 'price', label: 'Price Per Liter' },
  { key: 'total_cost', label: 'Total Cost' },
  { key: 'meter_start', label: 'Meter Reading Start' },
  { key: 'meter_end', label: 'Meter Reading End' },
  { key: 'mileage', label: 'KMS' },
  { key: 'voucher', label: 'Voucher No' },
  { key: 'area', label: 'Area' },
  { key: 'created_date', label: 'created_Date' },
  { key: 'created_by', label: 'created_by' },
];

export const PENDING_COLS: { key: keyof VoucherRow; label: string }[] = [
  { key: 'vehicle', label: 'Vehicle Number' },
  { key: 'usage_date', label: 'Usage Date' },
  { key: 'liters', label: 'Liters Used' },
  { key: 'fuel_type', label: 'Fuel Type' },
  { key: 'price', label: 'Price Per Liter' },
  { key: 'meter_start', label: 'Meter Reading Start' },
  { key: 'meter_end', label: 'Meter Reading End' },
  { key: 'area', label: 'Area' },
];

export const CSV_COLS: { key: keyof VoucherRow; label: string }[] = [
  { key: 'usage_id', label: 'Usage ID' },
  { key: 'vehicle', label: 'Vehicle Number' },
  { key: 'fuel_type', label: 'Fuel Type' },
  { key: 'usage_date', label: 'Usage Date' },
  { key: 'liters', label: 'Liters Used' },
  { key: 'shift', label: 'SHIFT' },
  { key: 'price', label: 'Price Per Liter' },
  { key: 'total_cost', label: 'Total Cost' },
  { key: 'meter_start', label: 'Meter Reading Start' },
  { key: 'meter_end', label: 'Meter Reading End' },
  { key: 'mileage', label: 'KMS' },
  { key: 'voucher', label: 'Voucher No' },
  { key: 'area', label: 'Area' },
  { key: 'created_date', label: 'Created Date' },
  { key: 'created_by', label: 'Created By' },
];

export function isCommonVehicle(v: string): boolean {
  return /^common\s*-/i.test((v || '').trim());
}

export function commonFuelKind(v: string): 'petrol' | 'diesel' | null {
  const s = (v || '').trim();
  if (/petrol/i.test(s)) return 'petrol';
  if (/diesel/i.test(s)) return 'diesel';
  return null;
}

export function commonCategoriesFor(v: string): string[] {
  const kind = commonFuelKind(v);
  if (kind === 'petrol') return COMMON_USAGE_CATEGORIES_PETROL;
  if (kind === 'diesel') return COMMON_USAGE_CATEGORIES_DIESEL;
  return [...COMMON_USAGE_CATEGORIES_PETROL, ...COMMON_USAGE_CATEGORIES_DIESEL];
}

export function commonVehiclesFor(v: string): string[] {
  const kind = commonFuelKind(v);
  if (kind === 'petrol') return COMMON_USAGE_VEHICLES_PETROL;
  if (kind === 'diesel') return COMMON_USAGE_VEHICLES_DIESEL;
  return [...COMMON_USAGE_VEHICLES_PETROL, ...COMMON_USAGE_VEHICLES_DIESEL];
}

export function todayIso(): string {
  return new Date().toISOString().slice(0, 10);
}

export function fmt2(n: unknown): string {
  const v = parseFloat(String(n));
  if (Number.isNaN(v)) return n == null ? '' : String(n);
  return v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export function fmtPrice(v: unknown): string {
  if (v === '' || v == null) return '';
  const n = parseFloat(String(v));
  if (Number.isNaN(n)) return String(v);
  return 'Rs.' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export function fmtCost(v: unknown): string {
  if (v === '' || v == null) return '';
  const n = parseFloat(String(v));
  if (Number.isNaN(n)) return String(v);
  return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export function computeMileage(r: VoucherRow): void {
  const s = parseFloat(r.meter_start);
  const e = parseFloat(r.meter_end);
  if (!isNaN(s) && !isNaN(e)) r.mileage = (e - s).toFixed(2);
  else r.mileage = '';
}

export function calcTotal(r: VoucherRow): void {
  const l = parseFloat(r.liters);
  const p = parseFloat(r.price);
  if (!Number.isNaN(l) && !Number.isNaN(p)) r.total_cost = (l * p).toFixed(2);
  else r.total_cost = '';
}

export function determineShift(usageDate: string): 'DAY' | 'NIGHT' {
  const s = String(usageDate || '');
  const m = s.match(/(\d{1,2}):(\d{2})/);
  if (!m) return 'DAY';
  const h = parseInt(m[1], 10);
  return h >= 6 && h < 18 ? 'DAY' : 'NIGHT';
}

export function normalizeVehicleName(v: string): string {
  let s = String(v || '').replace(/\s+/g, ' ').trim();
  if (!s) return '';
  if (/^common/i.test(s)) s = s.replace(/\s*-\s*/g, ' - ').replace(/\s+/g, ' ').trim();
  return s;
}

export function normFuelName(s: string): string {
  const t = String(s || '').replace(/\s+/g, ' ').trim();
  if (/super/i.test(t)) return 'Super Diesel';
  if (/diesel/i.test(t)) return 'Diesel';
  if (/kerosen|kerosin/i.test(t)) return 'Kerosine Oil';
  if (/petrol/i.test(t)) {
    if (/95/.test(t)) return 'Petrol 95 Octan';
    return 'Petrol 92 Octan';
  }
  return t;
}

export function cleanMileage(s: string): string {
  const d = String(s || '').replace(/[^\d]/g, '');
  if (!d || /^0+$/.test(d) || d === '3') return '';
  return d.replace(/^0+(?=\d)/, '');
}

export interface ParsedOneFuelRow {
  seq?: string;
  vehicle: string;
  fuel?: string;
  ppl?: number | null;
  amount?: number | null;
  datetime: string;
  liters: string;
  meter_end?: string;
}

export function parseOneFuelLayout(pages: PdfViewportItem[][]): {
  rows: ParsedOneFuelRow[];
  bad: { seq: string; why: string }[];
  layoutOk: boolean;
  summary: { count: number | null; litres: number | null; amount: number | null };
} {
  const num = (s: string) => parseFloat(String(s || '').replace(/,/g, ''));
  let cols: { key: string; cx: number }[] | null = null;

  for (const items of pages) {
    const hdr = items.find((i) => i.str.trim() === 'Emp Id');
    if (!hdr) continue;
    const first = items.filter((i) => Math.abs(i.y - hdr.y) <= 3 && i.str.trim());
    const named = first.map((i) => ({ name: i.str.trim(), cx: i.x + i.w / 2, x: i.x })).sort((a, b) => a.x - b.x);
    let devSeen = 0;
    cols = named.map((c) => {
      let key = c.name;
      if (c.name === 'Device') {
        key = devSeen === 0 ? 'seq' : 'devid';
        devSeen++;
      } else if (c.name === 'Vehicle') key = 'vehicle';
      else if (c.name === 'Fuel') key = 'fuel';
      else if (c.name === 'PPL') key = 'ppl';
      else if (c.name === 'Fill Up') key = 'fill';
      else if (/^Fill up Amt/i.test(c.name)) key = 'amt';
      else if (/^Mileage/.test(c.name)) key = 'mileage';
      else if (c.name === 'DateTime') key = 'dt';
      return { key, cx: c.cx };
    });
    break;
  }

  const need = ['seq', 'vehicle', 'fuel', 'ppl', 'fill', 'amt', 'mileage', 'dt'];
  if (!cols || need.some((k) => !cols!.some((c) => c.key === k))) {
    return { rows: [], bad: [], summary: { count: null, litres: null, amount: null }, layoutOk: false };
  }
  const seqCx = cols.find((c) => c.key === 'seq')!.cx;
  const nearest = (cx: number) => {
    let b = cols![0];
    let d = 1e9;
    for (const c of cols!) {
      const dd = Math.abs(c.cx - cx);
      if (dd < d) {
        d = dd;
        b = c;
      }
    }
    return b;
  };

  const rows: ParsedOneFuelRow[] = [];
  const bad: { seq: string; why: string }[] = [];
  const seen = new Set<string>();
  const STOP = /^(\*\*\*|One\(TM\)|Total\b)/i;
  let allText = '';

  for (const items of pages) {
    allText += items.map((i) => i.str).join(' ') + '\n';
    const anchors = items
      .filter((i) => /^\d{7,10}$/.test(i.str.trim()) && Math.abs(i.x + i.w / 2 - seqCx) < 30)
      .sort((a, b) => a.y - b.y);
    anchors.forEach((a, k) => {
      const top = a.y - 4;
      const bottom = Math.min(k + 1 < anchors.length ? anchors[k + 1].y - 4 : 1e9, a.y + 48);
      const cell: Record<string, string[]> = {};
      items
        .filter((i) => i.y >= top && i.y < bottom && i.str.trim() && !STOP.test(i.str.trim()))
        .sort((p, q) => (Math.abs(p.y - q.y) < 2 ? p.x - q.x : p.y - q.y))
        .forEach((i) => {
          const c = nearest(i.x + i.w / 2);
          (cell[c.key] = cell[c.key] || []).push(i.str.trim());
        });
      const g = (kName: string) => (cell[kName] || []).join(' ').replace(/\s+/g, ' ').trim();
      const seq = a.str.trim();
      if (seen.has(seq)) return;
      const dm = g('dt').match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*(\d{0,2})\s+(\d{1,2}):(\d{2})\s*(AM|PM)/i);
      const liters = num(g('fill'));
      const ppl = num(g('ppl'));
      const amt = num(g('amt'));
      const veh = normalizeVehicleName(g('vehicle'));
      if (!dm || !isFinite(liters) || !veh) {
        bad.push({ seq, why: !veh ? 'vehicle' : !dm ? 'date/time' : 'litres' });
        return;
      }
      seen.add(seq);
      let year = dm[3] + (dm[3].length < 4 ? dm[4] : '');
      if (year.length === 2) year = '20' + year;
      let h = parseInt(dm[5], 10);
      const ap = dm[7].toUpperCase();
      if (ap === 'PM' && h < 12) h += 12;
      if (ap === 'AM' && h === 12) h = 0;
      rows.push({
        seq,
        vehicle: veh,
        fuel: normFuelName(g('fuel')),
        ppl: isFinite(ppl) ? ppl : null,
        amount: isFinite(amt) ? amt : null,
        datetime: `${year}-${dm[2].padStart(2, '0')}-${dm[1].padStart(2, '0')} ${String(h).padStart(2, '0')}:${dm[6]}`,
        liters: String(liters),
        meter_end: cleanMileage(g('mileage')),
      });
    });
  }

  const flat = allText.replace(/\s+/g, ' ');
  const m1 = flat.match(/Total Number of Transactions\s*:\s*(\d+)/i);
  const m2 = flat.match(/Total Fill Up Litres\s*:\s*([\d,.]+)/i);
  const m3 =
    flat.match(/Total Payable Amount\s*\(LKR\)\s*:?\s*([\d,]+\.\d+)/i) ||
    flat.match(/([\d,]+\.\d{2})\s+Total Amount\s*\(LKR\)/i);

  return {
    rows,
    bad,
    layoutOk: true,
    summary: { count: m1 ? +m1[1] : null, litres: m2 ? num(m2[1]) : null, amount: m3 ? num(m3[1]) : null },
  };
}

export function isOneFuelReport(text: string): boolean {
  return /OneFuel/i.test(text) || (/Transaction\s+Report/i.test(text) && /Fill\s*Up/i.test(text));
}

export function parseOneFuelText(text: string): ParsedOneFuelRow[] {
  const t = String(text || '').replace(/\s+/g, ' ');
  const FUEL = '(Diesel|Super\\s+Diesel|Petrol(?:\\s+(?:92|95)(?:\\s+Octane?)?)?|Kerosene|Auto\\s+Diesel)';
  const re = new RegExp(
    '([A-Z0-9]{1,3}\\s?-\\s?\\d{3,4})\\s+' +
      FUEL +
      '\\s+([\\d,.]+)\\s+' +
      '([\\d,.]+)\\s+([\\d,.]+)\\s+([\\d,.]+)\\s+' +
      '(\\d{1,2})/(\\d{1,2})/(\\d{2,4})\\s*(\\d{0,2})\\s+(\\d{1,2}):(\\d{2})\\s*(AM|PM)',
    'gi'
  );
  const rows: ParsedOneFuelRow[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(t))) {
    const vehicle = m[1].replace(/\s+/g, '');
    const fuelType = m[2].replace(/\s+/g, ' ').trim();
    const ppl2 = m[3];
    const fill = m[4];
    const day = m[7].padStart(2, '0');
    const mon = m[8].padStart(2, '0');
    let year = m[9] + (m[9].length < 4 ? m[10] : '');
    if (year.length === 2) year = '20' + year;
    let h = parseInt(m[11], 10);
    const ap2 = m[13].toUpperCase();
    if (ap2 === 'PM' && h < 12) h += 12;
    if (ap2 === 'AM' && h === 12) h = 0;
    const litersNum = parseFloat(String(fill).replace(/,/g, ''));
    if (!isFinite(litersNum)) continue;
    rows.push({
      vehicle: normalizeVehicleName(vehicle),
      fuel: fuelType,
      ppl: parseFloat(String(ppl2).replace(/,/g, '')),
      datetime: `${year}-${mon}-${day} ${String(h).padStart(2, '0')}:${m[12]}`,
      liters: String(litersNum),
    });
  }
  return rows;
}

function toIso(dateTok: string, timeTok: string, ampmTok: string): string {
  const dm = dateTok.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  const tm = timeTok.match(/^(\d{1,2}):(\d{2})/);
  if (!dm || !tm) return '';
  const mo = dm[1].padStart(2, '0');
  const da = dm[2].padStart(2, '0');
  let yr = dm[3];
  if (yr.length === 2) yr = '20' + yr;
  let hh = parseInt(tm[1], 10);
  const mins = tm[2];
  const ap = (ampmTok || '').toUpperCase();
  if (ap === 'PM' && hh < 12) hh += 12;
  if (ap === 'AM' && hh === 12) hh = 0;
  return `${yr}-${mo}-${da} ${String(hh).padStart(2, '0')}:${mins}`;
}

export function parsePumpedRecordText(text: string): ParsedOneFuelRow[] {
  const HEADER_WORDS = new Set([
    'vehicle', 'date', 'time', 'liters', 'litres', 'litre', 'liter', 'amount', 'no', 'number', 'id',
    'usage', 'used', 'use', 'fuel', 'type', 'voucher', 'area', 'created', 'by', 'meter', 'reading',
    'start', 'end', 'total', 'cost', 'price', 'per', 'milage', 'mileage', 'odometer', 'pumped',
    'record', 'qty', 'quantity', 'product', 'rate', 'order', 'bill', 'invoice', 'supplier', 'shift', 'day', 'night',
  ]);
  const allTokens = String(text || '').split(/\s+/).filter(Boolean);
  const tokens = allTokens.filter((t) => !HEADER_WORDS.has(t.toLowerCase()));
  const isDate = (t: string) => /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(t);
  const isTime = (t: string) => /^\d{1,2}:\d{2}(:\d{2})?$/.test(t);
  const isAmpm = (t: string) => /^(AM|PM)$/i.test(t);
  const isNum = (t: string) => /^\d+(\.\d+)?$/.test(t);
  const isVehLike = (t: string) => {
    if (!t) return false;
    if (isDate(t) || isTime(t) || isAmpm(t)) return false;
    if (/^[-/&]+$/.test(t)) return false;
    if (/^[\d.]+$/.test(t)) return false;
    if (HEADER_WORDS.has(t.toLowerCase())) return false;
    return /[A-Za-z]/.test(t);
  };
  const anchors: { i: number; litersIdx: number }[] = [];
  for (let i = 0; i + 3 < tokens.length; i++) {
    if (!isDate(tokens[i]) || !isTime(tokens[i + 1]) || !isAmpm(tokens[i + 2]) || !isNum(tokens[i + 3])) continue;
    anchors.push({ i, litersIdx: i + 3 });
  }
  if (!anchors.length) return [];

  let format: 'date-first' | 'vehicle-first' = 'date-first';
  const t0 = tokens[0];
  if (isVehLike(t0)) format = 'vehicle-first';
  else if (!isDate(t0)) {
    const a0 = anchors[0];
    if (a0.i > 0 && isVehLike(tokens[a0.i - 1])) format = 'vehicle-first';
  }

  const rows: ParsedOneFuelRow[] = [];
  if (format === 'date-first') {
    for (let k = 0; k < anchors.length; k++) {
      const a = anchors[k];
      const nextStart = k + 1 < anchors.length ? anchors[k + 1].i : tokens.length;
      const parts: string[] = [];
      for (let j = a.litersIdx + 1; j < nextStart; j++) {
        const t = tokens[j];
        if (isDate(t) || isTime(t) || isAmpm(t)) break;
        if (/^[-/&]+$/.test(t)) {
          if (parts.length) {
            parts.push(t);
            continue;
          }
          break;
        }
        if (isNum(t) && !/[A-Za-z]/.test(t)) break;
        parts.push(t);
      }
      const vehicle = normalizeVehicleName(parts.join(' '));
      if (!vehicle) continue;
      rows.push({ vehicle, datetime: toIso(tokens[a.i], tokens[a.i + 1], tokens[a.i + 2]), liters: tokens[a.litersIdx] });
    }
  } else {
    for (let k = 0; k < anchors.length; k++) {
      const a = anchors[k];
      const prevEnd = k > 0 ? anchors[k - 1].litersIdx + 1 : 0;
      const parts: string[] = [];
      for (let j = prevEnd; j < a.i; j++) {
        const t = tokens[j];
        if (isDate(t) || isTime(t) || isAmpm(t)) break;
        if (/^[-/&]+$/.test(t)) {
          if (parts.length) {
            parts.push(t);
            continue;
          }
        } else parts.push(t);
      }
      const vehicle = normalizeVehicleName(parts.join(' '));
      if (!vehicle) continue;
      rows.push({ vehicle, datetime: toIso(tokens[a.i], tokens[a.i + 1], tokens[a.i + 2]), liters: tokens[a.litersIdx] });
    }
  }
  return rows;
}

/* ---------- Date helpers for Applied Vouchers ---------- */
const MONTH_NAMES: Record<string, number> = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
};
const MONTH_LABEL = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

export function parseDateParts(v: string): { y: number | null; mo: number | null; d: number } | null {
  const s = String(v || '').trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})/);
  if (m) return { y: +m[1], mo: +m[2], d: +m[3] };
  m = s.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})/);
  if (m) {
    let y = +m[3];
    if (y < 100) y += 2000;
    return { y, mo: +m[2], d: +m[1] };
  }
  m = s.match(/^(\d{1,2})[\s-]([a-zA-Z]{3,})[\s-]?(\d{2,4})?/);
  if (m) {
    const mo = MONTH_NAMES[m[2].slice(0, 3).toLowerCase()];
    if (mo) {
      let y = m[3] ? +m[3] : null;
      if (y && y < 100) y += 2000;
      return { y, mo, d: +m[1] };
    }
  }
  m = s.match(/^(\d{1,2})$/);
  if (m) return { y: null, mo: null, d: +m[1] };
  return null;
}

export function dayKey(v: string): string {
  const p = parseDateParts(v);
  if (!p) return '';
  return p.y && p.mo
    ? `${p.y}-${String(p.mo).padStart(2, '0')}-${String(p.d).padStart(2, '0')}`
    : `d${String(p.d).padStart(2, '0')}`;
}

export function dayLabel(v: string): string {
  const p = parseDateParts(v);
  if (!p) return String(v || '—');
  return p.mo ? `${String(p.d).padStart(2, '0')} ${MONTH_LABEL[p.mo]}` : String(p.d).padStart(2, '0');
}

export function daySortValue(v: string): number {
  const p = parseDateParts(v);
  if (!p) return 0;
  return (p.y || 0) * 10000 + (p.mo || 0) * 100 + (p.d || 0);
}

/* ---------- ExcelJS & jsPDF Exports for Voucher Analysis ---------- */
const XL_NUM_KEYS = ['__serial', 'liters', 'price', 'total_cost', 'meter_start', 'meter_end', 'mileage'];
const XL_WIDTH: Record<string, number> = {
  __serial: 9, usage_id: 11, vehicle: 17, fuel_type: 17, usage_date: 18, liters: 12,
  shift: 9, price: 14, total_cost: 14, meter_start: 14, meter_end: 14, mileage: 10,
  voucher: 12, area: 24, created_date: 13, created_by: 20,
};
const PDF_WIDTH: Record<string, number> = {
  __serial: 10, usage_id: 13, vehicle: 20, fuel_type: 18, usage_date: 26, liters: 15,
  shift: 12, price: 16, total_cost: 19, meter_start: 17, meter_end: 17, mileage: 13,
  voucher: 14, area: 22, created_date: 19, created_by: 18,
};

function xlNumber(v: unknown): number | null {
  const n = parseFloat(String(v ?? '').replace(/,/g, ''));
  return Number.isFinite(n) ? n : null;
}

export function buildAppliedWorkbook(sheets: { name: string; rows: VoucherRow[]; totals?: boolean }[]): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Fuel Voucher Analysis Tool';
  wb.created = new Date();
  const GREY = 'FF999999';
  const border = {
    top: { style: 'thin' as const, color: { argb: GREY } },
    left: { style: 'thin' as const, color: { argb: GREY } },
    bottom: { style: 'thin' as const, color: { argb: GREY } },
    right: { style: 'thin' as const, color: { argb: GREY } },
  };
  sheets.forEach((sh) => {
    const ws = wb.addWorksheet(sh.name, { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = APPLIED_COLS.map((c) => ({ header: c.label, key: c.key, width: XL_WIDTH[c.key] || 14 }));
    const hr = ws.getRow(1);
    hr.height = 30;
    hr.eachCell((cell) => {
      cell.font = { bold: true, size: 10, name: 'Calibri' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEEEEE' } };
      cell.border = border;
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    });
    let tl = 0;
    let tc = 0;
    sh.rows.forEach((r, i) => {
      const row = ws.addRow({});
      APPLIED_COLS.forEach((c, ci) => {
        const cell = row.getCell(ci + 1);
        let v: string | number | null;
        if (c.key === '__serial') v = i + 1;
        else if (XL_NUM_KEYS.includes(c.key)) v = xlNumber(r[c.key as keyof VoucherRow]);
        else v = (r[c.key as keyof VoucherRow] as string) ?? '';
        cell.value = v === null || v === '' ? null : v;
        cell.border = border;
        cell.font = { size: 10, name: 'Calibri' };
        const isNum = XL_NUM_KEYS.includes(c.key);
        cell.alignment = { vertical: 'middle', horizontal: isNum ? 'right' : c.key === 'shift' ? 'center' : 'left' };
        if (c.key === 'liters') cell.numFmt = '#,##0.00#';
        if (c.key === 'price') cell.numFmt = '"Rs."#,##0.00';
        if (c.key === 'total_cost') cell.numFmt = '#,##0.00';
        if (c.key === 'mileage') cell.numFmt = '#,##0.00';
      });
      tl += parseFloat(r.liters) || 0;
      tc += parseFloat(r.total_cost) || 0;
    });
    if (sh.totals !== false) {
      const row = ws.addRow({});
      const n = row.number;
      ws.mergeCells(n, 1, n, 5);
      row.getCell(1).value = 'TOTALS';
      row.getCell(6).value = Math.round(tl * 1000) / 1000;
      row.getCell(6).numFmt = '#,##0.00';
      row.getCell(9).value = Math.round(tc * 100) / 100;
      row.getCell(9).numFmt = '#,##0.00';
      for (let ci = 1; ci <= APPLIED_COLS.length; ci++) {
        const cell = row.getCell(ci);
        cell.font = { bold: true, size: 10, name: 'Calibri' };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF4F4F4' } };
        cell.border = { ...border, top: { style: 'medium', color: { argb: 'FF333333' } } };
        cell.alignment = { vertical: 'middle', horizontal: ci === 1 || ci === 6 || ci === 9 ? 'right' : 'left' };
      }
    }
  });
  return wb;
}

export function buildAppliedPdf(opt: {
  title: string;
  meta: string;
  rows: VoucherRow[];
  pending?: VoucherRow[];
}): jsPDF {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const W = doc.internal.pageSize.getWidth();
  const fmt = (n: number) => (n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const labels = APPLIED_COLS.map((c) => c.label);
  const cellText = (r: VoucherRow, c: (typeof APPLIED_COLS)[number], i: number) => {
    if (c.key === '__serial') return String(i + 1);
    if (c.key === 'liters')
      return (parseFloat(r.liters) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 3 });
    if (c.key === 'price') return r.price === '' || r.price == null ? '' : fmtPrice(r.price);
    if (c.key === 'total_cost') return fmtCost(r.total_cost);
    return String(r[c.key as keyof VoucherRow] ?? '');
  };
  const numIdx: Record<number, boolean> = {};
  APPLIED_COLS.forEach((c, i) => {
    if (XL_NUM_KEYS.includes(c.key)) numIdx[i] = true;
  });
  const columnStyles: Record<number, any> = {};
  APPLIED_COLS.forEach((c, i) => {
    columnStyles[i] = {
      cellWidth: PDF_WIDTH[c.key] || 16,
      halign: numIdx[i] ? 'right' : c.key === 'shift' ? 'center' : 'left',
    };
  });
  const LINE: [number, number, number] = [153, 153, 153];
  let y = 12;
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(13);
  doc.text(opt.title, W / 2, y, { align: 'center' });
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8);
  doc.setTextColor(85);
  doc.text(opt.meta, W / 2, y + 5, { align: 'center' });
  doc.setTextColor(0);
  y += 9;

  const tables: { heading: string | null; rows: VoucherRow[]; totals: boolean }[] = [
    { heading: null, rows: opt.rows, totals: true },
  ];
  if (opt.pending && opt.pending.length) {
    tables.push({ heading: 'Pending (Not Yet Applied to a Voucher)', rows: opt.pending, totals: false });
  }
  tables.forEach((t) => {
    if (t.heading) {
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(10);
      doc.text(t.heading, 8, y + 4);
      y += 7;
    }
    const tl = t.rows.reduce((s, r) => s + (parseFloat(r.liters) || 0), 0);
    const tc = t.rows.reduce((s, r) => s + (parseFloat(r.total_cost) || 0), 0);
    const foot = t.totals
      ? [
          [
            { content: 'TOTALS', colSpan: 5, styles: { halign: 'right' as const } },
            { content: fmt(tl), styles: { halign: 'right' as const } },
            { content: '', colSpan: 2 },
            { content: fmt(tc), styles: { halign: 'right' as const } },
            { content: '', colSpan: 7 },
          ],
        ]
      : undefined;
    autoTable(doc, {
      startY: y,
      head: [labels],
      body: t.rows.map((r, i) => APPLIED_COLS.map((c) => cellText(r, c, i))),
      foot,
      theme: 'grid',
      showHead: 'everyPage',
      showFoot: 'lastPage',
      margin: { left: 8, right: 8, top: 12, bottom: 12 },
      styles: {
        font: 'helvetica',
        fontSize: 6.5,
        cellPadding: 1.3,
        lineColor: LINE,
        lineWidth: 0.15,
        textColor: [27, 36, 48],
        overflow: 'linebreak',
        valign: 'middle',
      },
      headStyles: {
        fillColor: [238, 238, 238],
        textColor: [20, 20, 20],
        fontStyle: 'bold',
        halign: 'center',
        lineColor: LINE,
        lineWidth: 0.15,
      },
      footStyles: {
        fillColor: [244, 244, 244],
        textColor: [20, 20, 20],
        fontStyle: 'bold',
        lineColor: [51, 51, 51],
        lineWidth: 0.2,
      },
      columnStyles,
      rowPageBreak: 'avoid',
    });
    y = (doc as any).lastAutoTable.finalY + 5;
    if (t.totals) {
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(9.5);
      doc.setTextColor(0);
      doc.text('USAGE ' + fmt(tl) + ' L     TOTAL Rs. ' + fmt(tc), W - 8, y, { align: 'right' });
      y += 8;
    }
  });
  const n = doc.getNumberOfPages();
  for (let i = 1; i <= n; i++) {
    doc.setPage(i);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(110);
    doc.text('Page ' + i + ' of ' + n, W - 8, doc.internal.pageSize.getHeight() - 5, { align: 'right' });
  }
  return doc;
}

export function parseCsvText(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cur = '';
  let q = false;
  text = String(text || '').replace(/^\ufeff/, '');
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          cur += '"';
          i++;
        } else q = false;
      } else cur += c;
    } else if (c === '"') q = true;
    else if (c === ',') {
      row.push(cur);
      cur = '';
    } else if (c === '\n') {
      row.push(cur);
      rows.push(row);
      row = [];
      cur = '';
    } else if (c !== '\r') cur += c;
  }
  if (cur !== '' || row.length) {
    row.push(cur);
    rows.push(row);
  }
  return rows;
}

export function xlCellText(v: any): string {
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((t: any) => t.text).join('');
    if (v.text != null) return String(v.text);
    if (v.result != null) return String(v.result);
    return '';
  }
  return String(v);
}

export function saveBlob(blob: Blob, name: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}
