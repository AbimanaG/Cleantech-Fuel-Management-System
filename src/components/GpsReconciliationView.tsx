import React, { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  Upload,
  RotateCcw,
  Download,
  Printer,
  Search,
  Check,
  X,
  Plus,
  RefreshCw,
  ArrowRight,
  AlertTriangle,
  Send,
} from 'lucide-react';
import type {
  AbnormalItem,
  CardRecord,
  CompletedFileModule,
  GpsFillRecord,
  ManagementComment,
  OfflineReason,
  UnusualFlag,
} from '../types/fuel';
import { extractPdfData, extractPdfText } from '../utils/pdfReader';
import {
  buildAutoAbnormalRows,
  buildGpsPrintableHtml,
  buildGpsReportData,
  buildTimeVarianceRows,
  dateDiffDays,
  EXPECTED_LOCATION,
  fmtDT,
  getReason,
  gpsDtList,
  gpsFooterTotal,
  groupCardOnlyByCategory,
  isOnSite,
  normFuel,
  OFFLINE_CATEGORIES,
  parseCardWorkbook,
  parseGpsText,
  parseMobitelPdf,
  shiftPredictedDates,
  TIME_MIN_SHOW,
  timeRisk,
  UNCLASSIFIED_LABEL,
  UNUSUAL_VARIANCE_L,
  VARIANCE_TOLERANCE_L,
  vehicleHasOnSiteGps,
  vtype,
} from '../utils/gpsMatcherParsers';
import { saveBlob } from '../utils/voucherParser';

interface GpsReconciliationViewProps {
  cardRows: CardRecord[];
  setCardRows: React.Dispatch<React.SetStateAction<CardRecord[]>>;
  gpsRows: GpsFillRecord[];
  setGpsRows: React.Dispatch<React.SetStateAction<GpsFillRecord[]>>;
  allocations: Record<string, string[]>;
  setAllocations: React.Dispatch<React.SetStateAction<Record<string, string[]>>>;
  offlineReasons: Record<string, OfflineReason>;
  setOfflineReasons: React.Dispatch<React.SetStateAction<Record<string, OfflineReason>>>;
  skipped: Record<string, boolean>;
  setSkipped: React.Dispatch<React.SetStateAction<Record<string, boolean>>>;
  unusualMap: Record<string, UnusualFlag>;
  setUnusualMap: React.Dispatch<React.SetStateAction<Record<string, UnusualFlag>>>;
  abnormalItems: AbnormalItem[] | null;
  setAbnormalItems: React.Dispatch<React.SetStateAction<AbnormalItem[] | null>>;
  abnormalEdited: boolean;
  setAbnormalEdited: React.Dispatch<React.SetStateAction<boolean>>;
  onPushToConsolidation: () => void;
  onPushToVouchers?: () => void;
  onPublishCompletedFile?: (payload: {
    title: string;
    fileName: string;
    module: CompletedFileModule;
    format: 'xlsx' | 'pdf' | 'csv' | 'html';
    mimeType: string;
    blob: Blob;
    summary: string;
    reportDate: string;
  }) => Promise<void>;
  managementComments?: ManagementComment[];
}

export const GpsReconciliationView: React.FC<GpsReconciliationViewProps> = ({
  cardRows,
  setCardRows,
  gpsRows,
  setGpsRows,
  allocations,
  setAllocations,
  offlineReasons,
  setOfflineReasons,
  skipped,
  setSkipped,
  unusualMap,
  setUnusualMap,
  abnormalItems,
  setAbnormalItems,
  abnormalEdited,
  setAbnormalEdited,
  onPushToConsolidation,
  onPushToVouchers,
  onPublishCompletedFile,
  managementComments = [],
}) => {
  const [publishing, setPublishing] = useState(false);
  const [cardFileLabels, setCardFileLabels] = useState<string[]>([]);
  const [gpsFileLabels, setGpsFileLabels] = useState<string[]>([]);
  const [alert, setAlert] = useState<{ msg: string; kind: 'info' | 'ok' | 'warn' } | null>(null);
  const [preferSameShift, setPreferSameShift] = useState(false);
  const [allocSearch, setAllocSearch] = useState('');
  const [onlyUnallocated, setOnlyUnallocated] = useState(false);
  const [showSkipped, setShowSkipped] = useState(false);

  // Modal state
  const [modalCardId, setModalCardId] = useState<string | null>(null);
  const [modalSelectedIds, setModalSelectedIds] = useState<string[]>([]);
  const [modalSearch, setModalSearch] = useState('');
  const [modalUnusualChecked, setModalUnusualChecked] = useState(false);
  const [modalUnusualNote, setModalUnusualNote] = useState('');
  const [modalUnusualTouched, setModalUnusualTouched] = useState(false);

  const handleCardUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (!files.length) return;
    setAlert({ msg: `Reading ${files.length} fuel-card file(s)…`, kind: 'info' });
    const nextRows: Omit<CardRecord, 'id' | 'idx'>[] = [];
    const labels: string[] = [];
    const warnings: string[] = [];
    const seenTxn = new Set<string>();

    for (const file of files) {
      try {
        const isPdf = /\.pdf$/i.test(file.name) || file.type === 'application/pdf';
        let rows: Omit<CardRecord, 'id' | 'idx'>[];
        if (isPdf) {
          const data = await extractPdfData(file);
          const res = parseMobitelPdf(data, file.name);
          rows = res.rows.filter((r) => {
            const key = String(r.no);
            if (seenTxn.has(key)) return false;
            seenTxn.add(key);
            return true;
          });
          res.warnings.forEach((w) => warnings.push(`${file.name}: ${w}`));
          if (!res.rows.length) {
            warnings.push(`${file.name}: no Mobitel OneFuel transactions recognised`);
          }
        } else {
          const buf = await file.arrayBuffer();
          const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
          rows = parseCardWorkbook(wb, file.name);
        }
        rows.forEach((r) => nextRows.push(r));
        labels.push(`${file.name} (${rows.length}${isPdf ? ' · Mobitel PDF' : ''})`);
      } catch {
        warnings.push(`${file.name}: could not be read`);
      }
    }

    const indexed: CardRecord[] = nextRows.map((r, i) => ({ ...r, id: 'C' + (i + 1), idx: i + 1 }));
    setCardRows(indexed);
    setAllocations({});
    setOfflineReasons({});
    setUnusualMap({});
    setSkipped({});
    setAbnormalItems(null);
    setAbnormalEdited(false);
    setCardFileLabels(labels);

    const distinctVeh = new Set(indexed.map((r) => r.vehicleKey)).size;
    setAlert({
      msg:
        `Card: ${indexed.length} entries, ${distinctVeh} distinct vehicles, ${files.length} file(s).` +
        (warnings.length ? '  ⚠ ' + warnings.join(' | ') : ''),
      kind: warnings.length ? 'warn' : 'ok',
    });
    e.target.value = '';
  };

  const handleGpsUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (!files.length) return;
    setAlert({ msg: `Reading ${files.length} GPS PDF file(s)…`, kind: 'info' });
    const rawRows: Omit<GpsFillRecord, 'id' | 'idx'>[] = [];
    const labels: string[] = [];
    const warnings: string[] = [];

    for (const file of files) {
      try {
        const text = await extractPdfText(file);
        const rows = parseGpsText(text, file.name);
        const foot = gpsFooterTotal(text);
        const got = rows.reduce((a, r) => a + r.liters, 0);
        if (foot != null && Math.abs(foot - got) > 0.1) {
          warnings.push(`${file.name}: report total ${foot.toFixed(2)} L but ${got.toFixed(2)} L read`);
        }
        rows.forEach((r) => rawRows.push(r));
        labels.push(`${file.name} (${rows.length})`);
      } catch {
        warnings.push(`${file.name}: could not be read`);
      }
    }

    const seen = new Set<string>();
    const unique: GpsFillRecord[] = [];
    for (const r of rawRows) {
      const k = `${r.vehicleKey}|${r.date}|${r.time}|${r.liters}|${r.location}`;
      if (seen.has(k)) continue;
      seen.add(k);
      unique.push({ ...r, id: 'G' + (unique.length + 1), idx: unique.length + 1 });
    }

    setGpsRows(unique);
    setAllocations({});
    setOfflineReasons({});
    setUnusualMap({});
    setSkipped({});
    setAbnormalItems(null);
    setAbnormalEdited(false);
    setGpsFileLabels(labels);

    const onSite = unique.filter(isOnSite).length;
    const offSite = unique.length - onSite;
    setAlert({
      msg:
        `GPS: ${onSite} Sedawatta fills selectable · ${offSite} off-site fills reported separately.` +
        (warnings.length ? '  ⚠ ' + warnings.join(' | ') : ''),
      kind: warnings.length ? 'warn' : 'ok',
    });
    e.target.value = '';
  };

  const resetAll = () => {
    setCardRows([]);
    setGpsRows([]);
    setAllocations({});
    setOfflineReasons({});
    setUnusualMap({});
    setSkipped({});
    setShowSkipped(false);
    setAbnormalItems(null);
    setAbnormalEdited(false);
    setCardFileLabels([]);
    setGpsFileLabels([]);
    setAllocSearch('');
    setOnlyUnallocated(false);
    setAlert(null);
  };

  const hasSelectableFills = (card: CardRecord) => {
    const taken = new Set<string>();
    Object.keys(allocations).forEach((k) => {
      if (k !== card.id) allocations[k].forEach((id) => taken.add(id));
    });
    const pred = shiftPredictedDates(card);
    return gpsRows.some(
      (g) =>
        g.vehicleKey === card.vehicleKey &&
        isOnSite(g) &&
        !taken.has(g.id) &&
        (!preferSameShift || pred.includes(g.date))
    );
  };

  const withGpsCards = useMemo(
    () => cardRows.filter((c) => vehicleHasOnSiteGps(c.vehicleKey, gpsRows)),
    [cardRows, gpsRows]
  );

  const visibleWithGps = useMemo(() => {
    const q = allocSearch.toLowerCase().trim();
    return withGpsCards.filter((c) => {
      const allocs = allocations[c.id] || [];
      if (!showSkipped && skipped[c.id] && !allocs.length) return false;
      if (onlyUnallocated && allocs.length > 0) return false;
      if (q) {
        const key = `${c.vehicle} ${c.area} ${c.shift} ${c.fuel} ${c.date} ${c.station || ''} ${c.vehicleType || ''}`.toLowerCase();
        if (!key.includes(q)) return false;
      }
      return true;
    });
  }, [withGpsCards, allocations, showSkipped, skipped, onlyUnallocated, allocSearch]);

  const skippedCount = useMemo(
    () => withGpsCards.filter((c) => skipped[c.id] && !(allocations[c.id] || []).length).length,
    [withGpsCards, skipped, allocations]
  );

  const reportData = useMemo(() => {
    if (!cardRows.length || !gpsRows.length) return null;
    return buildGpsReportData(cardRows, gpsRows, allocations, offlineReasons, skipped, unusualMap);
  }, [cardRows, gpsRows, allocations, offlineReasons, skipped, unusualMap]);

  const activeAbnormalItems = useMemo(() => {
    if (!reportData) return [];
    if (abnormalItems !== null && abnormalEdited) return abnormalItems;
    return buildAutoAbnormalRows(reportData).map((x, i) => ({ ...x, id: i + 1 }));
  }, [reportData, abnormalItems, abnormalEdited]);

  const updateAbnormalField = (id: number, field: keyof AbnormalItem, val: string) => {
    const base = activeAbnormalItems.map((x) => ({ ...x }));
    const target = base.find((x) => x.id === id);
    if (!target) return;
    if (field === 'litres') target.litres = parseFloat(val) || 0;
    else if (field !== 'id') (target as any)[field] = val;
    setAbnormalItems(base);
    setAbnormalEdited(true);
  };

  const removeAbnormalRow = (id: number) => {
    setAbnormalItems(activeAbnormalItems.filter((x) => x.id !== id));
    setAbnormalEdited(true);
  };

  const addAbnormalRow = () => {
    const maxId = activeAbnormalItems.reduce((m, x) => Math.max(m, x.id), 0);
    setAbnormalItems([
      ...activeAbnormalItems,
      { id: maxId + 1, cat: '', veh: '', issue: '', litres: 0, action: '' },
    ]);
    setAbnormalEdited(true);
  };

  const resetAbnormal = () => {
    setAbnormalItems(null);
    setAbnormalEdited(false);
  };

  // Modal helpers
  const modalCard = useMemo(
    () => cardRows.find((c) => c.id === modalCardId) || null,
    [cardRows, modalCardId]
  );

  const modalCandidates = useMemo(() => {
    if (!modalCard) return [];
    const predicted = shiftPredictedDates(modalCard);
    let list = gpsRows.filter((g) => g.vehicleKey === modalCard.vehicleKey && isOnSite(g));
    list.sort((a, b) => {
      const pa = predicted.indexOf(a.date);
      const pb = predicted.indexOf(b.date);
      const ra = pa >= 0 ? pa : 100 + Math.abs(dateDiffDays(a.date, modalCard.date));
      const rb = pb >= 0 ? pb : 100 + Math.abs(dateDiffDays(b.date, modalCard.date));
      if (ra !== rb) return ra - rb;
      return (a.date + ' ' + a.time).localeCompare(b.date + ' ' + b.time);
    });
    if (preferSameShift) list = list.filter((g) => predicted.includes(g.date));
    return list;
  }, [modalCard, gpsRows, preferSameShift]);

  const modalReservedBy = useMemo(() => {
    const map: Record<string, string> = {};
    if (!modalCardId) return map;
    Object.keys(allocations).forEach((k) => {
      if (k !== modalCardId) allocations[k].forEach((id) => (map[id] = k));
    });
    return map;
  }, [allocations, modalCardId]);

  const openGpsModal = (cardId: string) => {
    const card = cardRows.find((c) => c.id === cardId);
    if (!card) return;
    const current = allocations[cardId] || [];
    const mu = unusualMap[cardId];
    setModalCardId(cardId);
    setModalSelectedIds(current);
    setModalSearch('');
    setModalUnusualTouched(!!mu);
    setModalUnusualChecked(!!(mu && !mu.off));
    setModalUnusualNote(mu && !mu.off ? mu.note || '' : '');
  };

  const toggleModalGpsId = (gpsId: string) => {
    if (!modalCard) return;
    const next = modalSelectedIds.includes(gpsId)
      ? modalSelectedIds.filter((x) => x !== gpsId)
      : [...modalSelectedIds, gpsId];
    setModalSelectedIds(next);
    if (!modalUnusualTouched) {
      const litres = next.reduce((s, id) => {
        const g = gpsRows.find((x) => x.id === id);
        return s + (g ? g.liters : 0);
      }, 0);
      setModalUnusualChecked(next.length > 0 && Math.abs(litres - modalCard.liters) > UNUSUAL_VARIANCE_L);
    }
  };

  const saveModalAllocation = () => {
    if (!modalCard) return;
    const nextAlloc = { ...allocations };
    const nextSkip = { ...skipped };
    const nextUnusual = { ...unusualMap };

    if (modalSelectedIds.length) {
      nextAlloc[modalCard.id] = modalSelectedIds;
      delete nextSkip[modalCard.id];
    } else {
      delete nextAlloc[modalCard.id];
    }

    const litresNow = modalSelectedIds.reduce((s, id) => {
      const g = gpsRows.find((x) => x.id === id);
      return s + (g ? g.liters : 0);
    }, 0);
    const bigVar = Math.abs(litresNow - modalCard.liters) > UNUSUAL_VARIANCE_L;

    if (modalSelectedIds.length && modalUnusualChecked) {
      nextUnusual[modalCard.id] = { note: modalUnusualNote.trim() };
    } else if (modalSelectedIds.length && bigVar) {
      nextUnusual[modalCard.id] = { note: '', off: true };
    } else {
      delete nextUnusual[modalCard.id];
    }

    setAllocations(nextAlloc);
    setSkipped(nextSkip);
    setUnusualMap(nextUnusual);
    setModalCardId(null);
  };

  const skipFromModal = () => {
    if (modalCard && !(allocations[modalCard.id] || []).length) {
      setSkipped((prev) => ({ ...prev, [modalCard.id]: true }));
    }
    setModalCardId(null);
  };

  const clearCardAllocation = (cardId: string) => {
    const nextAlloc = { ...allocations };
    const nextUnusual = { ...unusualMap };
    delete nextAlloc[cardId];
    delete nextUnusual[cardId];
    setAllocations(nextAlloc);
    setUnusualMap(nextUnusual);
  };

  // Exports
  const buildCsvText = (): string => {
    if (!reportData) return '';
    const d = reportData;
    const header = [
      'Section',
      'Card Date & Time',
      'GPS Date & Time',
      'Vehicle',
      'Vehicle Type',
      'Shift',
      'Area',
      'Location(s)',
      'Pumped (Card)',
      'Pumped (GPS)',
      'Variance',
      'Vari %',
      'Status',
      'Note',
    ];
    const q = (v: unknown) => {
      const s = (v ?? '').toString();
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    const lines = [header.join(',')];
    const sumCard = (arr: typeof d.matched) => arr.reduce((s, r) => s + r.cardL, 0);
    const sumGps = (arr: typeof d.matched) => arr.reduce((s, r) => s + r.gpsL, 0);
    const sumGpsOnly = (arr: typeof d.gpsOnlyOnSite) => arr.reduce((s, r) => s + r.gps.liters, 0);

    d.matched.forEach((r) =>
      lines.push(
        [
          r.unusual
            ? 'Unusual filling variance'
            : r.status === 'ok'
              ? '6. Variance OK'
              : `2. Variance > ±${VARIANCE_TOLERANCE_L} L`,
          fmtDT(r.card.date, r.card.time),
          gpsDtList(r.gpsList, ' | '),
          r.card.vehicle,
          vtype(r.card),
          r.card.shift,
          r.card.area,
          r.locations,
          r.cardL.toFixed(2),
          r.gpsL.toFixed(2),
          r.variance.toFixed(2),
          r.pct ?? '',
          r.unusual ? 'unusual' : r.status,
          r.gpsList.length + ' fill(s)' + (r.unusual && r.unusual.note ? ' — ' + r.unusual.note : ''),
        ]
          .map(q)
          .join(',')
      )
    );
    lines.push(
      ['Subtotal — Matched', '', '', '', '', '', '', '', sumCard(d.matched).toFixed(2), sumGps(d.matched).toFixed(2), '', '', '', '']
        .map(q)
        .join(',')
    );

    const { order, groups } = groupCardOnlyByCategory(d.cardOnly);
    order.forEach((cat) => {
      const list = groups[cat] || [];
      if (!list.length) return;
      list.forEach((r) => {
        const note = r.detail
          ? r.detail
          : r.hasGps
            ? r.skipped
              ? 'Skipped — no GPS fill to allocate'
              : 'GPS fills available — not allocated'
            : 'No GPS fills in PDF';
        lines.push(
          [
            `On card, not on GPS — ${cat}`,
            fmtDT(r.card.date, r.card.time),
            '',
            r.card.vehicle,
            vtype(r.card),
            r.card.shift,
            r.card.area,
            '',
            r.card.liters.toFixed(2),
            '',
            '',
            '',
            'gps-missing',
            note,
          ]
            .map(q)
            .join(',')
        );
      });
      lines.push(
        [
          `Subtotal — ${cat}`,
          '',
          '',
          '',
          '',
          '',
          '',
          '',
          list.reduce((s, r) => s + r.card.liters, 0).toFixed(2),
          '',
          '',
          '',
          '',
          '',
        ]
          .map(q)
          .join(',')
      );
    });

    d.gpsOnlyOnSite.forEach((r) =>
      lines.push(
        [
          'On GPS Sedawatta, not on card',
          '',
          r.gps.date + ' ' + r.gps.time,
          r.gps.vehicle,
          '',
          '',
          '',
          r.gps.location,
          '',
          r.gps.liters.toFixed(2),
          '',
          '',
          'card-missing',
          r.gps.description || '',
        ]
          .map(q)
          .join(',')
      )
    );
    lines.push(
      ['Subtotal — On GPS Sedawatta, not on card', '', '', '', '', '', '', '', '', sumGpsOnly(d.gpsOnlyOnSite).toFixed(2), '', '', '', '']
        .map(q)
        .join(',')
    );

    d.gpsOnlyOffSite.forEach((r) =>
      lines.push(
        [
          'On GPS off-site, not on card',
          '',
          r.gps.date + ' ' + r.gps.time,
          r.gps.vehicle,
          '',
          '',
          '',
          r.gps.location,
          '',
          r.gps.liters.toFixed(2),
          '',
          '',
          'card-missing',
          r.gps.description || '',
        ]
          .map(q)
          .join(',')
      )
    );

    activeAbnormalItems.forEach((x) =>
      lines.push(
        [
          '7. Abnormal highlights',
          '',
          '',
          x.veh,
          '',
          '',
          '',
          '',
          (+x.litres || 0).toFixed(2),
          '',
          '',
          '',
          '',
          `${x.cat} — ${x.issue} — ${x.action}`,
        ]
          .map(q)
          .join(',')
      )
    );
    lines.push(
      [
        'GRAND TOTAL',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        d.grandCard.toFixed(2),
        d.grandGps.toFixed(2),
        d.grandVar.toFixed(2),
        d.grandPct ?? '',
        '',
        '',
      ]
        .map(q)
        .join(',')
    );

    return lines.join('\n');
  };

  const downloadCsv = () => {
    const csvText = buildCsvText();
    if (!csvText) return;
    saveBlob(new Blob([csvText], { type: 'text/csv;charset=utf-8;' }), 'fuel_card_vs_gps_manual_report.csv');
  };

  const downloadPdf = () => {
    if (!reportData) return;
    const w = window.open('', '_blank');
    if (!w) {
      setAlert({ msg: 'Please allow pop-ups to generate the PDF report.', kind: 'warn' });
      return;
    }
    w.document.write(buildGpsPrintableHtml(cardRows, reportData, activeAbnormalItems));
    w.document.close();
  };

  const publishCompleted = async () => {
    if (!reportData || !onPublishCompletedFile) return;
    setPublishing(true);
    try {
      const dateStr = cardDateLabel === '-' ? '' : cardDateLabel;
      const stamp = (cardDates[0] || new Date().toISOString().slice(0, 10)).replace(/[^0-9-]/g, '');
      const summary = `${reportData.matchedOk.length} OK · ${reportData.matchedFlagged.length} variance · ${reportData.matchedUnusual.length} unusual · ${reportData.cardOnly.length} card-only`;

      const csvText = buildCsvText();
      await onPublishCompletedFile({
        title: `GPS Reconciliation Data (${dateStr || stamp})`,
        fileName: `GPS_Reconciliation_${stamp}.csv`,
        module: 'gps',
        format: 'csv',
        mimeType: 'text/csv;charset=utf-8',
        blob: new Blob([csvText], { type: 'text/csv;charset=utf-8' }),
        summary,
        reportDate: dateStr,
      });

      const htmlDoc = buildGpsPrintableHtml(cardRows, reportData, activeAbnormalItems);
      await onPublishCompletedFile({
        title: `GPS Reconciliation Full Report (${dateStr || stamp})`,
        fileName: `GPS_Reconciliation_Report_${stamp}.html`,
        module: 'gps',
        format: 'html',
        mimeType: 'text/html;charset=utf-8',
        blob: new Blob([htmlDoc], { type: 'text/html;charset=utf-8' }),
        summary,
        reportDate: dateStr,
      });

      setAlert({ msg: 'Published GPS Reconciliation CSV & Printable Report for Fuel Officer and Management.', kind: 'ok' });
    } catch (err: any) {
      setAlert({ msg: err?.message || 'Could not publish completed GPS report.', kind: 'warn' });
    } finally {
      setPublishing(false);
    }
  };

  const cardDates = [...new Set(cardRows.map((r) => r.date).filter(Boolean))].sort();
  const cardDateLabel =
    cardDates.length === 1
      ? cardDates[0]
      : cardDates.length
        ? `${cardDates[0]} → ${cardDates[cardDates.length - 1]} (${cardDates.length}d)`
        : '-';
  const totalCardLiters = cardRows.reduce((s, r) => s + r.liters, 0);
  const distinctCardVeh = new Set(cardRows.map((r) => r.vehicleKey)).size;

  const cardOnlyEntries = useMemo(() => {
    return cardRows
      .filter((c) => !(allocations[c.id] || []).length)
      .map((c) => ({
        card: c,
        hasGps: vehicleHasOnSiteGps(c.vehicleKey, gpsRows),
      }));
  }, [cardRows, allocations, gpsRows]);

  const timeVarianceRows = useMemo(
    () => (reportData ? buildTimeVarianceRows(reportData) : []),
    [reportData]
  );

  const hasMobitelRows = cardRows.some((r) => r.sheetName === 'Mobitel OneFuel');

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-end justify-between gap-4 border-b border-slate-200 pb-5">
        <div>
          <h1 className="text-xl font-bold text-slate-900">Fuel Card vs GPS Reconciliation</h1>
          <p className="text-sm text-slate-600 mt-1">
            Match Excel or Mobitel OneFuel PDF records against GPS Fuel Fillings PDFs at {EXPECTED_LOCATION}. Tolerance is ±{VARIANCE_TOLERANCE_L} L; variance above ±{UNUSUAL_VARIANCE_L} L flags as unusual.
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2 shrink-0">
          {hasMobitelRows && onPushToVouchers && (
            <button
              type="button"
              onClick={onPushToVouchers}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-white hover:bg-slate-50 text-slate-800 border border-slate-300 text-xs font-medium rounded-md transition-colors whitespace-nowrap"
            >
              Send to Voucher Analysis <ArrowRight className="w-3.5 h-3.5" />
            </button>
          )}
          {reportData && (
            <button
              type="button"
              onClick={onPushToConsolidation}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-[#1F5143] hover:bg-[#173D32] text-white text-xs font-medium rounded-md transition-colors whitespace-nowrap"
            >
              <RefreshCw className="w-3.5 h-3.5" /> Push to Multi-Day Excel
            </button>
          )}
          <button
            type="button"
            onClick={resetAll}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-white hover:bg-red-50 text-red-700 border border-red-200 text-xs font-medium rounded-md transition-colors whitespace-nowrap"
          >
            <RotateCcw className="w-3.5 h-3.5" /> Reset All
          </button>
        </div>
      </div>

      {/* Step 1: Upload Files */}
      <section className="bg-white border border-slate-200 rounded-md p-5 space-y-4">
        <h2 className="text-sm font-semibold text-slate-900">Step 1 — Upload Fuel Card &amp; GPS Files</h2>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div
            className={`border border-dashed rounded-md p-4 space-y-2 ${
              cardRows.length ? 'border-emerald-400 bg-emerald-50/40' : 'border-slate-300 bg-slate-50/60'
            }`}
          >
            <div className="text-xs font-semibold text-slate-800">
              Fuel Card Files — Excel or Mobitel OneFuel PDF
            </div>
            <label className="inline-flex items-center gap-2 px-3.5 py-2 bg-white hover:bg-slate-50 border border-slate-300 rounded-md text-xs font-medium text-slate-800 cursor-pointer transition-colors">
              <Upload className="w-3.5 h-3.5" /> Choose .xlsx / .xls / .pdf (multi-select)
              <input
                type="file"
                accept=".xlsx,.xls,.pdf,application/pdf"
                multiple
                className="hidden"
                onChange={handleCardUpload}
              />
            </label>
            <div className="text-xs text-slate-600">
              {cardRows.length
                ? `${cardRows.length} entries loaded · ${cardFileLabels.join(' · ')}`
                : 'No fuel card file selected'}
            </div>
          </div>

          <div
            className={`border border-dashed rounded-md p-4 space-y-2 ${
              gpsRows.length ? 'border-emerald-400 bg-emerald-50/40' : 'border-slate-300 bg-slate-50/60'
            }`}
          >
            <div className="text-xs font-semibold text-slate-800">
              GPS Fuel Fillings PDF Files
            </div>
            <label className="inline-flex items-center gap-2 px-3.5 py-2 bg-white hover:bg-slate-50 border border-slate-300 rounded-md text-xs font-medium text-slate-800 cursor-pointer transition-colors">
              <Upload className="w-3.5 h-3.5" /> Choose GPS PDF (multi-select)
              <input
                type="file"
                accept="application/pdf"
                multiple
                className="hidden"
                onChange={handleGpsUpload}
              />
            </label>
            <div className="text-xs text-slate-600">
              {gpsRows.length
                ? `${gpsRows.length} fills (${gpsRows.filter(isOnSite).length} on-site, ${
                    gpsRows.length - gpsRows.filter(isOnSite).length
                  } off-site) · ${gpsFileLabels.join(' · ')}`
                : 'No GPS PDF selected'}
            </div>
          </div>
        </div>

        <div className="flex flex-wrap items-center justify-between gap-4 pt-2">
          <label className="inline-flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer">
            <input
              type="checkbox"
              checked={preferSameShift}
              onChange={(e) => setPreferSameShift(e.target.checked)}
              className="rounded border-slate-300 text-[#1F5143] focus:ring-[#1F5143]"
            />
            Filter allocation modal to shift-predicted dates only
          </label>
        </div>

        {alert && (
          <div
            className={`px-3.5 py-2.5 rounded-md text-xs font-medium border ${
              alert.kind === 'warn'
                ? 'bg-amber-50 text-amber-800 border-amber-200'
                : alert.kind === 'ok'
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-200'
                  : 'bg-slate-50 text-slate-700 border-slate-200'
            }`}
          >
            {alert.msg}
          </div>
        )}
      </section>

      {/* Loaded Card Stats */}
      {cardRows.length > 0 && (
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
          <div className="bg-white border border-slate-200 rounded-md p-4">
            <div className="text-xs text-slate-500">Card Entries</div>
            <div className="font-data text-xl font-semibold text-slate-900 mt-1 tabular-nums">{cardRows.length}</div>
          </div>
          <div className="bg-white border border-slate-200 rounded-md p-4">
            <div className="text-xs text-slate-500">Distinct Vehicles</div>
            <div className="font-data text-xl font-semibold text-emerald-700 mt-1 tabular-nums">{distinctCardVeh}</div>
          </div>
          <div className="bg-white border border-slate-200 rounded-md p-4">
            <div className="text-xs text-slate-500">Total Litres (Card)</div>
            <div className="font-data text-xl font-semibold text-slate-900 mt-1 tabular-nums">
              {totalCardLiters.toFixed(2)} L
            </div>
          </div>
          <div className="bg-white border border-slate-200 rounded-md p-4">
            <div className="text-xs text-slate-500">Fuel Date(s)</div>
            <div className="font-data text-sm font-semibold text-amber-700 mt-2 tabular-nums">{cardDateLabel}</div>
          </div>
        </div>
      )}

      {/* Step 2: Allocate GPS fills */}
      <section className="bg-white border border-slate-200 rounded-md overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h2 className="text-sm font-semibold text-slate-900">
              Step 2 — Allocate GPS Fills ({EXPECTED_LOCATION} only)
            </h2>
            {skippedCount > 0 && (
              <div className="text-xs text-slate-500 mt-0.5">
                {skippedCount} skipped entr{skippedCount === 1 ? 'y' : 'ies'} {showSkipped ? 'shown' : 'hidden'} ·{' '}
                <button
                  type="button"
                  onClick={() => setShowSkipped((v) => !v)}
                  className="text-[#1F5143] font-medium hover:underline"
                >
                  {showSkipped ? 'Hide skipped' : 'Show / restore skipped'}
                </button>
              </div>
            )}
          </div>
          <div className="flex flex-wrap items-center gap-3">
            <div className="relative">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" />
              <input
                type="text"
                value={allocSearch}
                onChange={(e) => setAllocSearch(e.target.value)}
                placeholder="Search vehicle, area, fuel…"
                className="pl-8 pr-3 py-1.5 border border-slate-300 rounded-md text-xs w-56 focus:outline-none focus:border-[#1F5143]"
              />
            </div>
            <label className="inline-flex items-center gap-1.5 text-xs font-medium text-slate-700 cursor-pointer">
              <input
                type="checkbox"
                checked={onlyUnallocated}
                onChange={(e) => setOnlyUnallocated(e.target.checked)}
                className="rounded border-slate-300 text-[#1F5143] focus:ring-[#1F5143]"
              />
              Only unallocated
            </label>
          </div>
        </div>

        <div className="overflow-x-auto max-h-[65vh]">
          {!cardRows.length ? (
            <div className="p-10 text-center text-xs text-slate-400">Upload fuel-card file(s) to begin.</div>
          ) : !gpsRows.length ? (
            <div className="p-10 text-center text-xs text-slate-400">Upload GPS PDF file(s) to begin allocating.</div>
          ) : !withGpsCards.length ? (
            <div className="p-10 text-center text-xs text-slate-400">
              None of the card vehicles have {EXPECTED_LOCATION} GPS fills in the uploaded PDFs.
            </div>
          ) : !visibleWithGps.length ? (
            <div className="p-10 text-center text-xs text-slate-400">
              All matching entries are allocated or skipped.
            </div>
          ) : (
            <table className="w-full text-xs text-left border-collapse">
              <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 sticky top-0 z-10">
                <tr>
                  <th className="px-3 py-2.5 font-semibold">#</th>
                  <th className="px-3 py-2.5 font-semibold">Card Date &amp; Time</th>
                  <th className="px-3 py-2.5 font-semibold">Vehicle</th>
                  <th className="px-3 py-2.5 font-semibold">Vehicle Type</th>
                  <th className="px-3 py-2.5 font-semibold">Shift</th>
                  <th className="px-3 py-2.5 font-semibold">Area</th>
                  <th className="px-3 py-2.5 font-semibold">Fuel</th>
                  <th className="px-3 py-2.5 font-semibold text-right">Card L</th>
                  <th className="px-3 py-2.5 font-semibold">Allocated GPS Fills</th>
                  <th className="px-3 py-2.5 font-semibold text-right">GPS L</th>
                  <th className="px-3 py-2.5 font-semibold text-right">Variance</th>
                  <th className="px-3 py-2.5 font-semibold">Status</th>
                  <th className="px-3 py-2.5 font-semibold text-right">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {visibleWithGps.map((c) => {
                  const allocIds = allocations[c.id] || [];
                  const allocs = allocIds
                    .map((id) => gpsRows.find((g) => g.id === id))
                    .filter((g): g is GpsFillRecord => Boolean(g));
                  const gpsTotal = allocs.reduce((s, g) => s + g.liters, 0);
                  const variance = allocs.length ? +(gpsTotal - c.liters).toFixed(2) : null;
                  const within = variance != null && Math.abs(variance) <= VARIANCE_TOLERANCE_L;
                  const un = allocs.length ? (unusualMap[c.id]?.off ? null : unusualMap[c.id] || (variance != null && Math.abs(variance) > UNUSUAL_VARIANCE_L ? { note: 'auto', auto: true } : null)) : null;
                  const canPick = allocs.length > 0 || hasSelectableFills(c);

                  let statusText = 'Not allocated';
                  let statusClass = 'text-amber-700';
                  if (skipped[c.id] && !allocs.length) {
                    statusText = 'Skipped';
                    statusClass = 'text-slate-500';
                  } else if (!allocs.length && !canPick) {
                    statusText = 'No free fills';
                    statusClass = 'text-slate-400';
                  } else if (allocs.length > 0) {
                    statusText = within ? 'OK' : 'Variance';
                    statusClass = within ? 'text-emerald-700' : 'text-red-700';
                  }

                  return (
                    <tr key={c.id} className="hover:bg-slate-50/80">
                      <td className="px-3 py-2 font-data text-slate-400 tabular-nums">{c.idx}</td>
                      <td className="px-3 py-2 font-data whitespace-nowrap">{fmtDT(c.date, c.time)}</td>
                      <td className="px-3 py-2 font-data font-semibold text-slate-900">{c.vehicle}</td>
                      <td className="px-3 py-2 text-slate-600">{vtype(c)}</td>
                      <td className="px-3 py-2 font-data">{c.shift || '-'}</td>
                      <td className="px-3 py-2 text-slate-600">{c.area}</td>
                      <td className="px-3 py-2 text-slate-600">{c.fuel}</td>
                      <td className="px-3 py-2 font-data text-right tabular-nums">{c.liters.toFixed(2)}</td>
                      <td className="px-3 py-2">
                        {allocs.length ? (
                          <div className="space-y-0.5">
                            {allocs.map((g) => (
                              <div key={g.id} className="font-data text-[11px] text-slate-700">
                                {g.date} <span className="font-semibold text-slate-900">{g.time}</span> · {g.liters.toFixed(2)} L
                              </div>
                            ))}
                          </div>
                        ) : (
                          <span className="text-slate-400">— none —</span>
                        )}
                      </td>
                      <td className="px-3 py-2 font-data text-right tabular-nums">
                        {allocs.length ? gpsTotal.toFixed(2) : '-'}
                      </td>
                      <td className="px-3 py-2 font-data text-right tabular-nums font-medium">
                        {variance != null ? `${variance >= 0 ? '+' : ''}${variance.toFixed(2)} L` : '-'}
                      </td>
                      <td className={`px-3 py-2 font-medium whitespace-nowrap ${statusClass}`}>
                        {statusText}
                        {un && (
                          <span className="ml-1.5 text-red-700 font-semibold">
                            · Unusual{un.auto ? ' (auto)' : ''}
                          </span>
                        )}
                      </td>
                      <td className="px-3 py-2 text-right whitespace-nowrap space-x-1.5">
                        {skipped[c.id] && !allocs.length ? (
                          <button
                            type="button"
                            onClick={() => {
                              const next = { ...skipped };
                              delete next[c.id];
                              setSkipped(next);
                            }}
                            className="px-2.5 py-1 border border-slate-300 rounded text-slate-700 hover:bg-slate-100 font-medium"
                          >
                            Restore
                          </button>
                        ) : canPick ? (
                          <>
                            <button
                              type="button"
                              onClick={() => openGpsModal(c.id)}
                              className="px-2.5 py-1 bg-[#1F5143] hover:bg-[#173D32] text-white rounded font-medium"
                            >
                              {allocs.length ? 'Edit' : 'Allocate'}
                            </button>
                            {!allocs.length && (
                              <button
                                type="button"
                                onClick={() => setSkipped((p) => ({ ...p, [c.id]: true }))}
                                className="px-2 py-1 border border-slate-300 rounded text-slate-600 hover:bg-slate-100 font-medium"
                              >
                                Skip
                              </button>
                            )}
                          </>
                        ) : (
                          <span className="text-slate-400">No free fills</span>
                        )}
                        {allocs.length > 0 && (
                          <button
                            type="button"
                            onClick={() => clearCardAllocation(c.id)}
                            className="px-2 py-1 border border-red-200 text-red-700 hover:bg-red-50 rounded font-medium"
                          >
                            Clear
                          </button>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
        </div>
      </section>

      {/* Classify: On Card, Not on GPS */}
      {cardRows.length > 0 && gpsRows.length > 0 && cardOnlyEntries.length > 0 && (
        <section className="bg-white border border-slate-200 rounded-md overflow-hidden">
          <div className="px-5 py-4 border-b border-slate-200">
            <h2 className="text-sm font-semibold text-slate-900">
              On Card, Not on GPS — Classify Each Entry ({cardOnlyEntries.length})
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Categories: Offline, No GPS Tracker (auto-selected for tractors/staff/common), Not filled, Staff fuel, and Other.
            </p>
          </div>
          <div className="overflow-x-auto max-h-96">
            <table className="w-full text-xs text-left border-collapse">
              <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 sticky top-0">
                <tr>
                  <th className="px-3 py-2.5 font-semibold">#</th>
                  <th className="px-3 py-2.5 font-semibold">Card Date &amp; Time</th>
                  <th className="px-3 py-2.5 font-semibold">Vehicle</th>
                  <th className="px-3 py-2.5 font-semibold">Type</th>
                  <th className="px-3 py-2.5 font-semibold">Area</th>
                  <th className="px-3 py-2.5 font-semibold text-right">Card L</th>
                  <th className="px-3 py-2.5 font-semibold">GPS Availability</th>
                  <th className="px-3 py-2.5 font-semibold">Reason Category &amp; Detail</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {cardOnlyEntries.map(({ card: c, hasGps }) => {
                  const reason = getReason(c.id, cardRows, offlineReasons);
                  return (
                    <tr key={c.id} className="hover:bg-slate-50/80">
                      <td className="px-3 py-2 font-data text-slate-400">{c.idx}</td>
                      <td className="px-3 py-2 font-data whitespace-nowrap">{fmtDT(c.date, c.time)}</td>
                      <td className="px-3 py-2 font-data font-semibold text-slate-900">{c.vehicle}</td>
                      <td className="px-3 py-2 text-slate-600">{vtype(c)}</td>
                      <td className="px-3 py-2 text-slate-600">{c.area}</td>
                      <td className="px-3 py-2 font-data text-right tabular-nums">{c.liters.toFixed(2)}</td>
                      <td className="px-3 py-2">
                        <span className={hasGps ? 'text-blue-700 font-medium' : 'text-slate-500'}>
                          {hasGps ? 'GPS fills available — not allocated' : 'No GPS fills in PDF'}
                        </span>
                      </td>
                      <td className="px-3 py-2 space-y-1.5">
                        <select
                          value={reason.category}
                          onChange={(e) => {
                            const cat = e.target.value;
                            setOfflineReasons((prev) => ({
                              ...prev,
                              [c.id]: {
                                category: cat,
                                detail: cat === 'Other' ? prev[c.id]?.detail || '' : '',
                              },
                            }));
                          }}
                          className="px-2.5 py-1.5 border border-slate-300 rounded text-xs bg-white focus:outline-none focus:border-[#1F5143]"
                        >
                          <option value="">— Select —</option>
                          {OFFLINE_CATEGORIES.map((cat) => (
                            <option key={cat} value={cat}>
                              {cat}
                            </option>
                          ))}
                        </select>
                        {reason.category === 'Other' && (
                          <input
                            type="text"
                            value={reason.detail}
                            onChange={(e) => {
                              const val = e.target.value;
                              setOfflineReasons((prev) => ({
                                ...prev,
                                [c.id]: { category: 'Other', detail: val },
                              }));
                            }}
                            placeholder="Specify reason…"
                            className="block w-full min-w-[200px] px-2.5 py-1 border border-slate-300 rounded text-xs bg-white focus:outline-none focus:border-[#1F5143]"
                          />
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </section>
      )}

      {/* Step 3: Full 7-Section Variance Report */}
      {reportData && (
        <section className="space-y-6 pt-2">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-t border-slate-200 pt-6">
            <div>
              <h2 className="text-lg font-bold text-slate-900">Step 3 — Variance Report</h2>
              <p className="text-xs text-slate-500 mt-0.5">
                {reportData.matchedOk.length} OK · {reportData.matchedFlagged.length} Variance &gt; ±{VARIANCE_TOLERANCE_L} L ·{' '}
                {reportData.matchedUnusual.length} Unusual · {reportData.cardOnly.length} On card only ·{' '}
                {reportData.gpsOnlyOnSite.length} On GPS Sedawatta only · {reportData.gpsOnlyOffSite.length} Off-site GPS
              </p>
            </div>
            <div className="flex flex-wrap items-center gap-2">
              {managementComments.length > 0 && (
                <span className="text-xs text-slate-600 mr-2">
                  Management notes: {managementComments.length}
                </span>
              )}
              <button
                type="button"
                onClick={downloadCsv}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-white hover:bg-slate-50 text-slate-800 border border-slate-300 text-xs font-medium rounded-md transition-colors whitespace-nowrap"
              >
                <Download className="w-3.5 h-3.5" /> Download CSV
              </button>
              <button
                type="button"
                onClick={downloadPdf}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-[#1F5143] hover:bg-[#173D32] text-white text-xs font-medium rounded-md transition-colors whitespace-nowrap"
              >
                <Printer className="w-3.5 h-3.5" /> Download PDF Report
              </button>
              {onPublishCompletedFile && (
                <button
                  type="button"
                  disabled={publishing}
                  onClick={() => void publishCompleted()}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-slate-900 hover:bg-slate-800 text-white text-xs font-medium rounded-md transition-colors whitespace-nowrap disabled:opacity-60"
                >
                  <Send className="w-3.5 h-3.5" />{' '}
                  {publishing ? 'Publishing...' : 'Publish for Fuel Officer'}
                </button>
              )}
            </div>
          </div>

          {/* Summary Table */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <div className="bg-white border border-slate-200 rounded-md overflow-hidden">
              <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-800">
                Mobitel Card Fuel Summary
              </div>
              <table className="w-full text-xs text-left border-collapse">
                <thead className="bg-slate-50/50 border-b border-slate-200 text-slate-600">
                  <tr>
                    <th className="px-3 py-2">Date</th>
                    <th className="px-3 py-2">Fuel Type</th>
                    <th className="px-3 py-2 text-right">Pumped Ltr</th>
                    <th className="px-3 py-2 text-right">Amount</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {cardDates.map((dt) => {
                    const fuels = ['Diesel', 'Petrol', 'Super Diesel'] as const;
                    return fuels.map((f, idx) => {
                      const list = cardRows.filter((c) => c.date === dt && normFuel(c.fuel) === f);
                      const l = list.reduce((s, c) => s + c.liters, 0);
                      const a = list.reduce((s, c) => s + (c.amount || 0), 0);
                      return (
                        <tr key={`${dt}-${f}`}>
                          {idx === 0 && (
                            <td rowSpan={3} className="px-3 py-2 font-data font-semibold align-middle border-r border-slate-100">
                              {dt}
                            </td>
                          )}
                          <td className="px-3 py-1.5">{f}</td>
                          <td className="px-3 py-1.5 font-data text-right tabular-nums">{l ? l.toFixed(2) : '-'}</td>
                          <td className="px-3 py-1.5 font-data text-right tabular-nums">
                            {a ? a.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '-'}
                          </td>
                        </tr>
                      );
                    });
                  })}
                </tbody>
              </table>
            </div>

            <div className="bg-white border border-slate-200 rounded-md overflow-hidden">
              <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-800">
                Report Section Breakdown
              </div>
              <table className="w-full text-xs text-left border-collapse">
                <thead className="bg-slate-50/50 border-b border-slate-200 text-slate-600">
                  <tr>
                    <th className="px-3 py-2">Section</th>
                    <th className="px-3 py-2 text-right">Entries</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  <tr>
                    <td className="px-3 py-2">1. Unusual filling variance</td>
                    <td className="px-3 py-2 font-data text-right tabular-nums font-medium">{reportData.matchedUnusual.length}</td>
                  </tr>
                  <tr>
                    <td className="px-3 py-2">2. More than ±{VARIANCE_TOLERANCE_L} L variance</td>
                    <td className="px-3 py-2 font-data text-right tabular-nums font-medium">{reportData.matchedFlagged.length}</td>
                  </tr>
                  <tr>
                    <td className="px-3 py-2">3. GPS unavailable &amp; offline</td>
                    <td className="px-3 py-2 font-data text-right tabular-nums font-medium">{reportData.cardOnly.length}</td>
                  </tr>
                  <tr>
                    <td className="px-3 py-2">4. Time variance &gt; {TIME_MIN_SHOW} min</td>
                    <td className="px-3 py-2 font-data text-right tabular-nums font-medium">{timeVarianceRows.length}</td>
                  </tr>
                  <tr>
                    <td className="px-3 py-2">5. On GPS ({EXPECTED_LOCATION}), not on card</td>
                    <td className="px-3 py-2 font-data text-right tabular-nums font-medium">
                      {reportData.gpsOnlyOnSite.length}
                      {reportData.gpsOnlyOffSite.length ? ` (+${reportData.gpsOnlyOffSite.length} off-site)` : ''}
                    </td>
                  </tr>
                  <tr>
                    <td className="px-3 py-2">6. Variance within ±{VARIANCE_TOLERANCE_L} L — OK</td>
                    <td className="px-3 py-2 font-data text-right tabular-nums font-medium">{reportData.matchedOk.length}</td>
                  </tr>
                  <tr>
                    <td className="px-3 py-2">7. Abnormal highlights</td>
                    <td className="px-3 py-2 font-data text-right tabular-nums font-medium">{activeAbnormalItems.length}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          {/* Section 1 & Section 2 Matched Tables */}
          {[
            {
              title: `1. Unusual Filling Variance (${reportData.matchedUnusual.length})`,
              rows: reportData.matchedUnusual,
            },
            {
              title: `2. More than ±${VARIANCE_TOLERANCE_L} L Variance — Fuel Pump vs GPS (${reportData.matchedFlagged.length})`,
              rows: reportData.matchedFlagged,
            },
            {
              title: `6. Variance within ±${VARIANCE_TOLERANCE_L} L — OK (${reportData.matchedOk.length})`,
              rows: reportData.matchedOk,
            },
          ].map((sec) => (
            <div key={sec.title} className="bg-white border border-slate-200 rounded-md overflow-hidden">
              <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-900">
                {sec.title}
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs text-left border-collapse">
                  <thead className="bg-slate-50/50 border-b border-slate-200 text-slate-600">
                    <tr>
                      <th className="px-3 py-2">Card Date &amp; Time</th>
                      <th className="px-3 py-2">GPS Date &amp; Time</th>
                      <th className="px-3 py-2">Vehicle</th>
                      <th className="px-3 py-2">Type</th>
                      <th className="px-3 py-2">Area</th>
                      <th className="px-3 py-2 text-right">Card L</th>
                      <th className="px-3 py-2 text-right">GPS L</th>
                      <th className="px-3 py-2 text-right">Variance</th>
                      <th className="px-3 py-2 text-right">Vari %</th>
                      <th className="px-3 py-2">Status</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {sec.rows.length === 0 ? (
                      <tr>
                        <td colSpan={10} className="px-3 py-6 text-center text-slate-400">
                          None.
                        </td>
                      </tr>
                    ) : (
                      sec.rows.map((r) => (
                        <tr key={r.card.id} className="hover:bg-slate-50">
                          <td className="px-3 py-2 font-data whitespace-nowrap">{fmtDT(r.card.date, r.card.time)}</td>
                          <td className="px-3 py-2 font-data whitespace-nowrap">{gpsDtList(r.gpsList, ', ')}</td>
                          <td className="px-3 py-2 font-data font-semibold text-slate-900">{r.card.vehicle}</td>
                          <td className="px-3 py-2 text-slate-600">{vtype(r.card)}</td>
                          <td className="px-3 py-2 text-slate-600">{r.card.area || '-'}</td>
                          <td className="px-3 py-2 font-data text-right tabular-nums">{r.cardL.toFixed(2)}</td>
                          <td className="px-3 py-2 font-data text-right tabular-nums">{r.gpsL.toFixed(2)}</td>
                          <td className="px-3 py-2 font-data text-right tabular-nums font-medium">
                            {(r.variance >= 0 ? '+' : '') + r.variance.toFixed(2)}
                          </td>
                          <td className="px-3 py-2 font-data text-right tabular-nums">
                            {r.pct != null ? `${r.pct >= 0 ? '+' : ''}${r.pct}%` : '-'}
                          </td>
                          <td
                            className={`px-3 py-2 font-medium ${
                              r.unusual || r.status === 'mismatch' ? 'text-red-700' : 'text-emerald-700'
                            }`}
                          >
                            {r.unusual ? 'Unusual' : r.status === 'ok' ? 'OK' : 'Mismatch'}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          ))}

          {/* Section 3: Offline grouped summary + Section 4: Time Variance */}
          <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
            <div className="bg-white border border-slate-200 rounded-md overflow-hidden">
              <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-900">
                3. GPS Unavailable &amp; Offline Breakdown ({reportData.cardOnly.length})
              </div>
              <div className="overflow-x-auto max-h-80">
                <table className="w-full text-xs text-left border-collapse">
                  <thead className="bg-slate-50/50 border-b border-slate-200 text-slate-600 sticky top-0">
                    <tr>
                      <th className="px-3 py-2">Category</th>
                      <th className="px-3 py-2">Vehicle</th>
                      <th className="px-3 py-2">Date</th>
                      <th className="px-3 py-2 text-right">Card L</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {reportData.cardOnly.length === 0 ? (
                      <tr>
                        <td colSpan={4} className="px-3 py-6 text-center text-slate-400">
                          None.
                        </td>
                      </tr>
                    ) : (
                      reportData.cardOnly.map((r) => (
                        <tr key={r.card.id} className="hover:bg-slate-50">
                          <td className="px-3 py-2 font-medium text-amber-800">{r.category || UNCLASSIFIED_LABEL}</td>
                          <td className="px-3 py-2 font-data font-semibold text-slate-900">{r.card.vehicle}</td>
                          <td className="px-3 py-2 font-data whitespace-nowrap">{fmtDT(r.card.date, r.card.time)}</td>
                          <td className="px-3 py-2 font-data text-right tabular-nums">{r.card.liters.toFixed(2)}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-md overflow-hidden">
              <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-900">
                4. Fuel Pump Time Variance &gt; {TIME_MIN_SHOW} min ({timeVarianceRows.length})
              </div>
              <div className="overflow-x-auto max-h-80">
                <table className="w-full text-xs text-left border-collapse">
                  <thead className="bg-slate-50/50 border-b border-slate-200 text-slate-600 sticky top-0">
                    <tr>
                      <th className="px-3 py-2">Vehicle</th>
                      <th className="px-3 py-2">Card Time</th>
                      <th className="px-3 py-2">GPS Time</th>
                      <th className="px-3 py-2 text-right">Diff (Min)</th>
                      <th className="px-3 py-2">Risk</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {timeVarianceRows.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="px-3 py-6 text-center text-slate-400">
                          None.
                        </td>
                      </tr>
                    ) : (
                      timeVarianceRows.map((x, i) => {
                        const risk = timeRisk(x.mins);
                        return (
                          <tr key={i} className="hover:bg-slate-50">
                            <td className="px-3 py-2 font-data font-semibold text-slate-900">{x.r.card.vehicle}</td>
                            <td className="px-3 py-2 font-data whitespace-nowrap">{fmtDT(x.r.card.date, x.r.card.time)}</td>
                            <td className="px-3 py-2 font-data whitespace-nowrap">{x.first.date} {x.first.time}</td>
                            <td className="px-3 py-2 font-data text-right tabular-nums">{x.mins}</td>
                            <td
                              className={`px-3 py-2 font-medium ${
                                risk === 'Critical'
                                  ? 'text-red-700'
                                  : risk === 'High'
                                    ? 'text-amber-700'
                                    : 'text-emerald-700'
                              }`}
                            >
                              {risk}
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>

          {/* Section 7: Editable Abnormal Highlights */}
          <div className="bg-white border border-slate-200 rounded-md overflow-hidden">
            <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-2">
              <span className="text-xs font-semibold text-slate-900">
                7. Abnormal Highlights &amp; Recommendations ({activeAbnormalItems.length})
              </span>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={addAbnormalRow}
                  className="inline-flex items-center gap-1 px-2.5 py-1 bg-[#1F5143] hover:bg-[#173D32] text-white text-xs font-medium rounded"
                >
                  <Plus className="w-3 h-3" /> Add Row
                </button>
                <button
                  type="button"
                  onClick={resetAbnormal}
                  className="px-2.5 py-1 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 text-xs font-medium rounded"
                >
                  Reset to Auto-Detected
                </button>
              </div>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-xs text-left border-collapse">
                <thead className="bg-slate-50/50 border-b border-slate-200 text-slate-600">
                  <tr>
                    <th className="px-3 py-2 w-44">Category</th>
                    <th className="px-3 py-2 w-32">Vehicle</th>
                    <th className="px-3 py-2">Observed Issue</th>
                    <th className="px-3 py-2 w-24 text-right">Impact (L)</th>
                    <th className="px-3 py-2">Action</th>
                    <th className="px-3 py-2 w-12" />
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {activeAbnormalItems.length === 0 ? (
                    <tr>
                      <td colSpan={6} className="px-3 py-6 text-center text-slate-400">
                        None.
                      </td>
                    </tr>
                  ) : (
                    activeAbnormalItems.map((x) => (
                      <tr key={x.id}>
                        <td className="px-2 py-1.5">
                          <input
                            type="text"
                            value={x.cat}
                            onChange={(e) => updateAbnormalField(x.id, 'cat', e.target.value)}
                            className="w-full px-2 py-1 border border-slate-200 rounded text-xs"
                          />
                        </td>
                        <td className="px-2 py-1.5">
                          <input
                            type="text"
                            value={x.veh}
                            onChange={(e) => updateAbnormalField(x.id, 'veh', e.target.value)}
                            className="w-full px-2 py-1 border border-slate-200 rounded text-xs font-data"
                          />
                        </td>
                        <td className="px-2 py-1.5">
                          <input
                            type="text"
                            value={x.issue}
                            onChange={(e) => updateAbnormalField(x.id, 'issue', e.target.value)}
                            className="w-full px-2 py-1 border border-slate-200 rounded text-xs"
                          />
                        </td>
                        <td className="px-2 py-1.5">
                          <input
                            type="number"
                            step="0.01"
                            value={x.litres}
                            onChange={(e) => updateAbnormalField(x.id, 'litres', e.target.value)}
                            className="w-full px-2 py-1 border border-slate-200 rounded text-xs text-right font-data"
                          />
                        </td>
                        <td className="px-2 py-1.5">
                          <input
                            type="text"
                            value={x.action}
                            onChange={(e) => updateAbnormalField(x.id, 'action', e.target.value)}
                            className="w-full px-2 py-1 border border-slate-200 rounded text-xs"
                          />
                        </td>
                        <td className="px-2 py-1.5 text-center">
                          <button
                            type="button"
                            onClick={() => removeAbnormalRow(x.id)}
                            className="p-1 text-slate-400 hover:text-red-700"
                            aria-label="Remove row"
                          >
                            <X className="w-3.5 h-3.5" />
                          </button>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* Grand Total */}
          <div className="bg-emerald-50/60 border border-emerald-200 rounded-md px-5 py-3.5 flex flex-wrap items-center justify-between gap-4 text-sm font-semibold text-slate-900">
            <span>GRAND TOTAL RECONCILIATION</span>
            <div className="flex flex-wrap items-center gap-6 font-data tabular-nums">
              <span>Card: {reportData.grandCard.toFixed(2)} L</span>
              <span>GPS: {reportData.grandGps.toFixed(2)} L</span>
              <span>
                Variance: {(reportData.grandVar >= 0 ? '+' : '') + reportData.grandVar.toFixed(2)} L
                {reportData.grandPct != null
                  ? ` (${reportData.grandPct >= 0 ? '+' : ''}${reportData.grandPct}%)`
                  : ''}
              </span>
            </div>
          </div>
        </section>
      )}

      {/* GPS Allocation Modal */}
      {modalCard && (
        <div className="fixed inset-0 bg-slate-900/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-md border border-slate-200 w-full max-w-5xl max-h-[88vh] flex flex-col overflow-hidden shadow-xl">
            <div className="px-5 py-3.5 bg-[#1F5143] text-white flex items-center justify-between">
              <div>
                <h3 className="text-sm font-semibold">
                  Allocate GPS Fills — {modalCard.vehicle}
                  {modalCard.vehicleType ? ` (${modalCard.vehicleType})` : ''}
                </h3>
                <div className="text-xs text-emerald-100 mt-0.5 font-data">
                  Card #{modalCard.idx} · {fmtDT(modalCard.date, modalCard.time)} · Shift {modalCard.shift || '-'} ·{' '}
                  {modalCard.liters.toFixed(2)} L ({modalCard.fuel}) · Predicted:{' '}
                  {shiftPredictedDates(modalCard).join(' / ')} · Selected: {modalSelectedIds.length} fill(s),{' '}
                  {modalSelectedIds
                    .reduce((s, id) => s + (gpsRows.find((g) => g.id === id)?.liters || 0), 0)
                    .toFixed(2)}{' '}
                  L
                </div>
              </div>
              <button
                type="button"
                onClick={() => setModalCardId(null)}
                className="p-1 text-emerald-100 hover:text-white"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-5 overflow-y-auto space-y-3 flex-1">
              <div className="px-3.5 py-2 bg-slate-50 border border-slate-200 rounded text-xs text-slate-700">
                Only fills at <b>{EXPECTED_LOCATION}</b> are shown. Fills already allocated to another card entry are greyed out.
              </div>

              {modalCandidates.some((g) => !modalReservedBy[g.id]) && (
                <div className="px-3.5 py-2.5 bg-red-50/60 border border-red-200 rounded space-y-2 text-xs">
                  <label className="flex items-center gap-2 font-semibold text-red-800 cursor-pointer">
                    <input
                      type="checkbox"
                      disabled={modalSelectedIds.length === 0}
                      checked={modalUnusualChecked}
                      onChange={(e) => {
                        setModalUnusualTouched(true);
                        setModalUnusualChecked(e.target.checked);
                      }}
                    />
                    <AlertTriangle className="w-3.5 h-3.5" />
                    Mark as unusual filling variance (auto-ticked when variance &gt; ±{UNUSUAL_VARIANCE_L} L)
                  </label>
                  {modalUnusualChecked && (
                    <input
                      type="text"
                      value={modalUnusualNote}
                      onChange={(e) => setModalUnusualNote(e.target.value)}
                      placeholder="Reason / remark (optional) — printed in Unusual Variance section"
                      className="w-full px-3 py-1.5 border border-red-200 rounded bg-white text-xs"
                    />
                  )}
                </div>
              )}

              <input
                type="text"
                value={modalSearch}
                onChange={(e) => setModalSearch(e.target.value)}
                placeholder="Search date, time…"
                className="w-full px-3 py-2 border border-slate-300 rounded text-xs focus:outline-none focus:border-[#1F5143]"
              />

              <div className="border border-slate-200 rounded overflow-hidden">
                <table className="w-full text-xs text-left border-collapse">
                  <thead className="bg-slate-50 border-b border-slate-200 text-slate-600">
                    <tr>
                      <th className="px-3 py-2 w-10" />
                      <th className="px-3 py-2">No</th>
                      <th className="px-3 py-2">GPS Date</th>
                      <th className="px-3 py-2">GPS Time</th>
                      <th className="px-3 py-2">Vehicle</th>
                      <th className="px-3 py-2 text-right">Litres</th>
                      <th className="px-3 py-2">Location</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {modalCandidates.length === 0 ? (
                      <tr>
                        <td colSpan={7} className="px-4 py-8 text-center text-slate-400">
                          No {EXPECTED_LOCATION} GPS fills available for {modalCard.vehicle}.
                        </td>
                      </tr>
                    ) : (
                      modalCandidates
                        .filter((g) => {
                          if (!modalSearch.trim()) return true;
                          return `${g.date} ${g.time} ${g.vehicle} ${g.location}`
                            .toLowerCase()
                            .includes(modalSearch.toLowerCase().trim());
                        })
                        .map((g) => {
                          const isReserved = !!modalReservedBy[g.id];
                          const reservedCard = isReserved
                            ? cardRows.find((x) => x.id === modalReservedBy[g.id])
                            : null;
                          const isPredicted = shiftPredictedDates(modalCard).includes(g.date);
                          const checked = modalSelectedIds.includes(g.id);
                          return (
                            <tr
                              key={g.id}
                              onClick={() => !isReserved && toggleModalGpsId(g.id)}
                              className={
                                isReserved
                                  ? 'bg-slate-50 text-slate-400 cursor-not-allowed'
                                  : checked
                                    ? 'bg-emerald-50/60 cursor-pointer'
                                    : 'hover:bg-slate-50 cursor-pointer'
                              }
                            >
                              <td className="px-3 py-2">
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  disabled={isReserved}
                                  onChange={() => {}}
                                />
                              </td>
                              <td className="px-3 py-2 font-data">{g.no}</td>
                              <td className="px-3 py-2 font-data">
                                {g.date}
                                {!isPredicted && (
                                  <span className="ml-1.5 text-[11px] text-slate-400">(other date)</span>
                                )}
                                {isReserved && (
                                  <span className="ml-1.5 text-[11px] text-blue-700 font-medium">
                                    Allocated → #{reservedCard?.idx} {reservedCard?.vehicle}
                                  </span>
                                )}
                              </td>
                              <td className="px-3 py-2 font-data font-semibold">{g.time}</td>
                              <td className="px-3 py-2">
                                <div className="font-data font-semibold">{g.vehicle}</div>
                                {g.description && (
                                  <div className="text-[11px] text-slate-400">{g.description}</div>
                                )}
                              </td>
                              <td className="px-3 py-2 font-data text-right tabular-nums">{g.liters.toFixed(2)}</td>
                              <td className="px-3 py-2">{g.location}</td>
                            </tr>
                          );
                        })
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={skipFromModal}
                className="px-4 py-2 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 text-xs font-medium rounded-md"
              >
                Skip Entry
              </button>
              <button
                type="button"
                onClick={saveModalAllocation}
                className="inline-flex items-center gap-1.5 px-4 py-2 bg-[#1F5143] hover:bg-[#173D32] text-white text-xs font-medium rounded-md"
              >
                <Check className="w-3.5 h-3.5" /> Save Allocation
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
