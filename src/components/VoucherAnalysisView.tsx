import React, { useEffect, useMemo, useState } from 'react';
import {
  Upload,
  Check,
  Search,
  Download,
  Printer,
  Edit3,
  Trash2,
  Plus,
  X,
  RefreshCw,
  Send,
} from 'lucide-react';
import type {
  CompletedFileModule,
  ManagementComment,
  SplitLineDraft,
  UploadSummaryInfo,
  VoucherRow,
} from '../types/fuel';
import { AREA_LIST } from '../data/defaultAreas';
import { extractPdfLayout } from '../utils/pdfReader';
import {
  APPLIED_COLS,
  buildAppliedPdf,
  buildAppliedWorkbook,
  calcTotal,
  commonCategoriesFor,
  commonFuelKind,
  commonVehiclesFor,
  computeMileage,
  CSV_COLS,
  dayKey,
  dayLabel,
  daySortValue,
  determineShift,
  fmt2,
  fmtCost,
  fmtPrice,
  isCommonVehicle,
  isOneFuelReport,
  normFuelName,
  parseOneFuelLayout,
  parseOneFuelText,
  parsePumpedRecordText,
  saveBlob,
  todayIso,
  type ParsedOneFuelRow,
} from '../utils/voucherParser';
import { esc } from '../utils/gpsMatcherParsers';

interface VoucherAnalysisViewProps {
  pendingRows: VoucherRow[];
  setPendingRows: React.Dispatch<React.SetStateAction<VoucherRow[]>>;
  appliedRows: VoucherRow[];
  setAppliedRows: React.Dispatch<React.SetStateAction<VoucherRow[]>>;
  loadedSeqs: string[];
  setLoadedSeqs: React.Dispatch<React.SetStateAction<string[]>>;
  rowSeq: number;
  setRowSeq: React.Dispatch<React.SetStateAction<number>>;
  currentUser: string;
  setCurrentUser: (v: string) => void;
  userId: string;
  setUserId: (v: string) => void;
  selectedAppliedDays: string[];
  setSelectedAppliedDays: React.Dispatch<React.SetStateAction<string[]>>;
  selectedAppliedVouchers: string[];
  setSelectedAppliedVouchers: React.Dispatch<React.SetStateAction<string[]>>;
  vehicleAreas: Record<string, string[]>;
  canImportFromGps: boolean;
  onImportFromGps: () => void;
  onClearWorkspace: () => void;
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

export const VoucherAnalysisView: React.FC<VoucherAnalysisViewProps> = ({
  pendingRows,
  setPendingRows,
  appliedRows,
  setAppliedRows,
  loadedSeqs,
  setLoadedSeqs,
  rowSeq,
  setRowSeq,
  currentUser,
  setCurrentUser,
  userId,
  setUserId,
  selectedAppliedDays,
  setSelectedAppliedDays,
  selectedAppliedVouchers,
  setSelectedAppliedVouchers,
  vehicleAreas,
  canImportFromGps,
  onImportFromGps,
  onClearWorkspace,
  onPublishCompletedFile,
  managementComments = [],
}) => {
  const [publishing, setPublishing] = useState(false);
  const [subTab, setSubTab] = useState<'upload' | 'pending' | 'applied'>(
    pendingRows.length ? 'pending' : appliedRows.length ? 'applied' : 'upload'
  );
  const [pdfStatus, setPdfStatus] = useState('');
  const [uploadSummary, setUploadSummary] = useState<UploadSummaryInfo | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [pendingSearch, setPendingSearch] = useState('');
  const [appliedSearch, setAppliedSearch] = useState('');
  const [openSplitId, setOpenSplitId] = useState<string | null>(null);
  const [activeSuggestRowId, setActiveSuggestRowId] = useState<string | null>(null);
  const [bulkVoucher, setBulkVoucher] = useState('');

  // Edit modal
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editForm, setEditForm] = useState({
    voucher: '',
    fuel_type: '',
    shift: 'DAY' as 'DAY' | 'NIGHT',
    price: '',
    meter_start: '',
    meter_end: '',
  });

  // Expose live-edit callbacks for Print Preview popup
  useEffect(() => {
    (window as any).handlePrintMeterInput = (id: string, field: 'meter_start' | 'meter_end', value: string) => {
      setAppliedRows((prev) =>
        prev.map((r) => {
          if (r.__id !== id) return r;
          const next = { ...r, [field]: value };
          computeMileage(next);
          return next;
        })
      );
    };
    (window as any).handlePrintShiftInput = (id: string, value: 'DAY' | 'NIGHT') => {
      setAppliedRows((prev) => prev.map((r) => (r.__id === id ? { ...r, shift: value } : r)));
    };
    return () => {
      delete (window as any).handlePrintMeterInput;
      delete (window as any).handlePrintShiftInput;
    };
  }, [setAppliedRows]);

  const allKnownAreas = useMemo(() => {
    const set = new Set(AREA_LIST);
    Object.values(vehicleAreas)
      .flat()
      .forEach((a) => {
        if (a) set.add(a);
      });
    return [...set];
  }, [vehicleAreas]);

  const suggestedArea = (vehicle: string) => (vehicleAreas[vehicle] || [])[0] || '';

  const areaSuggestions = (vehicle: string, query: string) => {
    const q = (query || '').trim().toLowerCase();
    const vehAreas = vehicleAreas[vehicle] || [];
    const preferred = [...vehAreas, ...allKnownAreas.filter((a) => !vehAreas.includes(a))];
    return preferred.filter((a) => !q || a.toLowerCase().includes(q)).slice(0, 10);
  };

  const rowStatus = (r: VoucherRow): 'purple' | 'red' | 'amber' => {
    if (isCommonVehicle(r.vehicle)) return 'purple';
    if (!r.__areaConfirmed || !r.area) return 'red';
    return 'amber';
  };

  const handlePdfUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploadSummary(null);
    setPdfStatus('Reading PDF…');
    try {
      const { pages, text } = await extractPdfLayout(file);
      let parsed: ParsedOneFuelRow[] = [];
      let bad: { seq: string; why: string }[] = [];
      let meta: { count: number | null; litres: number | null; amount: number | null } = {
        count: null,
        litres: null,
        amount: null,
      };
      const lay = parseOneFuelLayout(pages);
      if (lay.layoutOk && lay.rows.length) {
        parsed = lay.rows;
        bad = lay.bad;
        meta = lay.summary;
      }
      if (!parsed.length && isOneFuelReport(text)) parsed = parseOneFuelText(text);
      if (!parsed.length) parsed = parsePumpedRecordText(text);
      if (!parsed.length) {
        setPdfStatus('Could not find transaction rows in that PDF.');
        e.target.value = '';
        return;
      }

      const seqSet = new Set(loadedSeqs);
      const fresh: ParsedOneFuelRow[] = [];
      let dups = 0;
      parsed.forEach((p) => {
        if (p.seq && seqSet.has(p.seq)) {
          dups++;
          return;
        }
        fresh.push(p);
      });

      let seqCounter = rowSeq;
      const newRows: VoucherRow[] = fresh.map((p) => {
        const row: VoucherRow = {
          __id: 'r' + seqCounter++,
          __seq: p.seq || '',
          __selected: false,
          __areaConfirmed: false,
          __isSplitChild: false,
          usage_id: '',
          vehicle: p.vehicle,
          fuel_type: p.fuel ? normFuelName(p.fuel) : '',
          usage_date: p.datetime,
          liters: p.liters,
          shift: determineShift(p.datetime),
          price: p.ppl != null && !isNaN(p.ppl) ? String(p.ppl) : '',
          total_cost: '',
          mileage: '',
          voucher: '',
          area: '',
          created_date: '',
          created_by: '',
          meter_start: '',
          meter_end: p.meter_end || '',
        };
        calcTotal(row);
        if (p.seq) seqSet.add(p.seq);
        return row;
      });

      setRowSeq(seqCounter);
      setLoadedSeqs([...seqSet]);
      setPendingRows((prev) => [...prev, ...newRows]);

      const nums = parsed
        .map((p) => parseInt(p.seq || '', 10))
        .filter((n) => !isNaN(n))
        .sort((a, b) => a - b);
      const have = new Set(nums);
      const missing: string[] = [];
      if (nums.length && meta.count != null && nums.length < meta.count) {
        for (let n = nums[0]; n <= nums[nums.length - 1] && missing.length < 200; n++) {
          if (!have.has(n)) missing.push(String(n));
        }
      }
      const fuelCounts: Record<string, number> = {};
      parsed.forEach((p) => {
        const k = (p.fuel || 'Unknown') + (p.ppl != null ? ' @ Rs.' + p.ppl.toFixed(2) : '');
        fuelCounts[k] = (fuelCounts[k] || 0) + 1;
      });

      setUploadSummary({
        taken: newRows.length,
        dups,
        bad,
        missing,
        reportCount: meta.count,
        reportLitres: meta.litres,
        reportAmount: meta.amount,
        sumLitres: parsed.reduce((s, p) => s + (parseFloat(p.liters) || 0), 0),
        sumAmount: parsed.reduce(
          (s, p) => s + (p.amount != null ? p.amount : (parseFloat(p.liters) || 0) * (p.ppl || 0)),
          0
        ),
        fuelCounts,
        commonCount: parsed.filter((p) => isCommonVehicle(p.vehicle)).length,
        noMileage: parsed.filter((p) => !p.meter_end).length,
      });
      setPdfStatus('Upload finished — continue in Pending Transactions.');
    } catch (err: any) {
      setPdfStatus('Could not read that PDF: ' + err.message);
    }
    e.target.value = '';
  };

  // Pending actions
  const tickAllSuggestedAreas = () => {
    if (!pendingRows.length) {
      setNotice('No pending transactions loaded yet.');
      return;
    }
    let filled = 0;
    let alreadyConfirmed = 0;
    let noSuggestion = 0;
    let commonSkipped = 0;

    setPendingRows((prev) =>
      prev.map((r) => {
        if (isCommonVehicle(r.vehicle)) {
          commonSkipped++;
          return r;
        }
        if (r.__areaConfirmed && r.area) {
          alreadyConfirmed++;
          return r;
        }
        const sug = suggestedArea(r.vehicle);
        if (sug) {
          filled++;
          return { ...r, area: sug, __areaConfirmed: true };
        }
        noSuggestion++;
        return r;
      })
    );

    const bits = [`${filled} auto-filled & confirmed`];
    if (alreadyConfirmed) bits.push(`${alreadyConfirmed} already confirmed`);
    if (commonSkipped) bits.push(`${commonSkipped} Common rows skipped (need Split usage)`);
    if (noSuggestion) bits.push(`${noSuggestion} without suggestion`);
    setNotice(bits.join(' · '));
  };

  const selectAllReady = () => {
    setPendingRows((prev) => prev.map((r) => (rowStatus(r) === 'amber' ? { ...r, __selected: true } : r)));
  };

  const clearSelection = () => {
    setPendingRows((prev) => prev.map((r) => ({ ...r, __selected: false })));
  };

  const updatePendingField = (id: string, patch: Partial<VoucherRow>) => {
    setPendingRows((prev) =>
      prev.map((r) => {
        if (r.__id !== id) return r;
        const next = { ...r, ...patch };
        computeMileage(next);
        return next;
      })
    );
  };

  const confirmArea = (id: string, explicitVal?: string) => {
    setPendingRows((prev) =>
      prev.map((r) => {
        if (r.__id !== id) return r;
        const val = (explicitVal !== undefined ? explicitVal : r.area || suggestedArea(r.vehicle)).trim();
        return {
          ...r,
          area: val,
          __areaConfirmed: !!val,
          __selected: val ? r.__selected : false,
        };
      })
    );
    setActiveSuggestRowId(null);
  };

  // Common Split actions
  const toggleSplitDialog = (id: string) => {
    if (openSplitId === id) {
      setOpenSplitId(null);
      return;
    }
    setPendingRows((prev) =>
      prev.map((r) =>
        r.__id === id && !r.__splitDraft
          ? { ...r, __splitDraft: [{ label: '', liters: '', meter_start: '', meter_end: '' }] }
          : r
      )
    );
    setOpenSplitId(id);
  };

  const updateSplitDraft = (id: string, updater: (draft: SplitLineDraft[]) => SplitLineDraft[]) => {
    setPendingRows((prev) =>
      prev.map((r) => (r.__id === id ? { ...r, __splitDraft: updater(r.__splitDraft || []) } : r))
    );
  };

  const confirmSplit = (r: VoucherRow) => {
    const target = parseFloat(r.liters) || 0;
    const draft = r.__splitDraft || [];
    const total = draft.reduce((s, l) => s + (parseFloat(l.liters) || 0), 0);
    if (Math.abs(total - target) >= 0.01) {
      setNotice(`Split total must match ${target.toFixed(2)} L exactly.`);
      return;
    }
    if (!draft.length || draft.some((l) => !l.label || !l.liters)) {
      setNotice('Each split line needs a "Used by" label and liters.');
      return;
    }
    const kindDefault =
      commonFuelKind(r.vehicle) === 'petrol'
        ? 'Petrol 92 Octan'
        : commonFuelKind(r.vehicle) === 'diesel'
          ? 'Diesel'
          : '';
    let seqCounter = rowSeq;
    const children: VoucherRow[] = draft.map((l) => {
      const c: VoucherRow = {
        __id: 'r' + seqCounter++,
        __isSplitChild: true,
        __splitParentVehicle: r.vehicle,
        __selected: false,
        __areaConfirmed: false,
        __seq: r.__seq || '',
        usage_id: r.usage_id || '',
        vehicle: l.label,
        fuel_type: r.fuel_type || kindDefault,
        usage_date: r.usage_date || '',
        liters: l.liters,
        shift: r.shift || determineShift(r.usage_date),
        price: r.price || '',
        total_cost: '',
        mileage: '',
        voucher: '',
        area: '',
        created_date: r.created_date || '',
        created_by: r.created_by || '',
        meter_start: l.meter_start || '',
        meter_end: l.meter_end || '',
      };
      computeMileage(c);
      calcTotal(c);
      return c;
    });
    setRowSeq(seqCounter);
    setPendingRows((prev) => {
      const idx = prev.findIndex((x) => x.__id === r.__id);
      if (idx < 0) return prev;
      const copy = [...prev];
      copy.splice(idx, 1, ...children);
      return copy;
    });
    setOpenSplitId(null);
  };

  // Bulk apply
  const selectedPending = useMemo(() => pendingRows.filter((r) => r.__selected), [pendingRows]);
  const selectedLitersSum = useMemo(
    () => selectedPending.reduce((s, r) => s + (parseFloat(r.liters) || 0), 0),
    [selectedPending]
  );

  const existingVoucherMatches = useMemo(() => {
    const v = bulkVoucher.trim();
    if (!v) return [];
    return appliedRows.filter((r) => r.voucher === v);
  }, [bulkVoucher, appliedRows]);

  const applyVoucher = () => {
    const voucher = bulkVoucher.trim();
    if (!selectedPending.length) {
      setNotice('Select at least one ready transaction.');
      return;
    }
    if (!voucher) {
      setNotice('Voucher No is required.');
      return;
    }
    const createdBy = currentUser.trim() || '—';
    const uid = userId.trim() || '—';
    const today = todayIso();

    const moved = selectedPending.map((r) => {
      const next: VoucherRow = {
        ...r,
        voucher,
        usage_id: uid,
        shift: r.shift || determineShift(r.usage_date),
        created_by: createdBy,
        created_date: today,
        __selected: false,
      };
      calcTotal(next);
      computeMileage(next);
      return next;
    });

    const selectedSet = new Set(selectedPending.map((r) => r.__id));
    setPendingRows((prev) => prev.filter((r) => !selectedSet.has(r.__id)));
    setAppliedRows((prev) => [...prev, ...moved].map((r, i) => ({ ...r, __serial: i + 1 })));
    setBulkVoucher('');
    setNotice(`Applied ${moved.length} record(s) to Voucher ${voucher}.`);
    setSubTab('applied');
  };

  // Applied filters
  const appliedDayKeysList = useMemo(() => {
    const seen = new Map<string, string>();
    appliedRows.forEach((r) => {
      const k = dayKey(r.usage_date);
      if (k && !seen.has(k)) seen.set(k, r.usage_date);
    });
    return [...seen.entries()]
      .sort((a, b) => daySortValue(a[1]) - daySortValue(b[1]))
      .map((e) => e[0]);
  }, [appliedRows]);

  const allVouchersList = useMemo(() => {
    return [...new Set(appliedRows.map((r) => r.voucher || '—'))].sort((a, b) =>
      a.localeCompare(b, undefined, { numeric: true })
    );
  }, [appliedRows]);

  const vouchersMatchingDays = useMemo(() => {
    if (!selectedAppliedDays.length) return new Set(allVouchersList);
    const daySet = new Set(selectedAppliedDays);
    return new Set(appliedRows.filter((r) => daySet.has(dayKey(r.usage_date))).map((r) => r.voucher || '—'));
  }, [selectedAppliedDays, appliedRows, allVouchersList]);

  const filteredAppliedRows = useMemo(() => {
    const daySet = new Set(selectedAppliedDays);
    const vSet = new Set(selectedAppliedVouchers);
    return appliedRows.filter(
      (r) =>
        (!daySet.size || daySet.has(dayKey(r.usage_date))) &&
        (!vSet.size || vSet.has(r.voucher || '—'))
    );
  }, [appliedRows, selectedAppliedDays, selectedAppliedVouchers]);

  const visibleAppliedRows = useMemo(() => {
    const q = appliedSearch.trim().toLowerCase();
    if (!q) return filteredAppliedRows;
    return filteredAppliedRows.filter((r) =>
      APPLIED_COLS.map((c) => (c.key === '__serial' ? r.__serial : r[c.key as keyof VoucherRow])).some((v) =>
        String(v ?? '')
          .toLowerCase()
          .includes(q)
      )
    );
  }, [filteredAppliedRows, appliedSearch]);

  const openEditModal = (r: VoucherRow) => {
    setEditingId(r.__id);
    setEditForm({
      voucher: r.voucher || '',
      fuel_type: r.fuel_type || '',
      shift: r.shift || 'DAY',
      price: r.price || '',
      meter_start: r.meter_start || '',
      meter_end: r.meter_end || '',
    });
  };

  const saveEditModal = () => {
    if (!editingId) return;
    setAppliedRows((prev) =>
      prev.map((r) => {
        if (r.__id !== editingId) return r;
        const next: VoucherRow = {
          ...r,
          voucher: editForm.voucher.trim(),
          fuel_type: editForm.fuel_type,
          shift: editForm.shift,
          price: editForm.price.trim(),
          meter_start: editForm.meter_start.trim(),
          meter_end: editForm.meter_end.trim(),
        };
        computeMileage(next);
        calcTotal(next);
        return next;
      })
    );
    setEditingId(null);
  };

  const deleteApplied = (id: string) => {
    const target = appliedRows.find((r) => r.__id === id);
    if (!target) return;
    const nextApplied: VoucherRow[] = appliedRows
      .filter((r) => r.__id !== id)
      .map((r, i) => ({ ...r, __serial: i + 1 }));
    setAppliedRows(nextApplied);
    if (target.__seq && ![...nextApplied, ...pendingRows].some((x) => x.__seq === target.__seq)) {
      setLoadedSeqs((prev) => prev.filter((s) => s !== target.__seq));
    }
  };

  // Exports
  const selectionLabels = () => {
    const dLabel = selectedAppliedDays.length
      ? selectedAppliedDays
          .map((k) => {
            const row = appliedRows.find((r) => dayKey(r.usage_date) === k);
            return row ? dayLabel(row.usage_date) : k;
          })
          .join(', ')
      : 'All days';
    const vLabel = selectedAppliedVouchers.length ? selectedAppliedVouchers.join(', ') : 'All vouchers';
    return { dLabel, vLabel };
  };

  const exportContext = (scope: 'selection' | 'full') => {
    const { dLabel, vLabel } = selectionLabels();
    const stamp = todayIso();
    const safe = (s: string) =>
      String(s)
        .replace(/[^A-Za-z0-9._-]+/g, '_')
        .replace(/^_+|_+$/g, '')
        .slice(0, 60) || 'All';
    if (scope === 'selection') {
      return {
        rows: filteredAppliedRows,
        pending: [] as VoucherRow[],
        title: 'Fuel Voucher Record',
        meta: `Days: ${dLabel}   |   Vouchers: ${vLabel}   |   Generated ${new Date().toLocaleString()}`,
        base: `Fuel_Voucher_${safe(vLabel === 'All vouchers' ? 'All' : vLabel)}_${stamp}`,
      };
    }
    return {
      rows: appliedRows,
      pending: pendingRows,
      title: 'Fuel Voucher — Complete Record',
      meta: `Generated ${new Date().toLocaleString()}   |   Applied: ${appliedRows.length}   |   Pending: ${pendingRows.length}`,
      base: `Fuel_Voucher_Complete_${stamp}`,
    };
  };

  const downloadAppliedExcel = async (scope: 'selection' | 'full') => {
    const ctx = exportContext(scope);
    if (!ctx.rows.length) {
      setNotice('No applied records to export for this selection.');
      return;
    }
    const sheets: { name: string; rows: VoucherRow[]; totals?: boolean }[] = [
      { name: scope === 'selection' ? 'Fuel Voucher' : 'Applied', rows: ctx.rows },
    ];
    if (ctx.pending.length) sheets.push({ name: 'Pending', rows: ctx.pending, totals: false });
    const wb = buildAppliedWorkbook(sheets);
    const buf = await wb.xlsx.writeBuffer();
    saveBlob(
      new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      ctx.base + '.xlsx'
    );
  };

  const downloadAppliedPdf = (scope: 'selection' | 'full') => {
    const ctx = exportContext(scope);
    if (!ctx.rows.length) {
      setNotice('No applied records to export for this selection.');
      return;
    }
    const doc = buildAppliedPdf(ctx);
    doc.save(ctx.base + '.pdf');
  };

  const publishVoucherFiles = async (scope: 'selection' | 'full') => {
    if (!onPublishCompletedFile) return;
    const ctx = exportContext(scope);
    if (!ctx.rows.length) {
      setNotice('No applied records to publish for this selection.');
      return;
    }
    setPublishing(true);
    try {
      const sheets: { name: string; rows: VoucherRow[]; totals?: boolean }[] = [
        { name: scope === 'selection' ? 'Fuel Voucher' : 'Applied', rows: ctx.rows },
      ];
      if (ctx.pending.length) sheets.push({ name: 'Pending', rows: ctx.pending, totals: false });
      const wb = buildAppliedWorkbook(sheets);
      const buf = await wb.xlsx.writeBuffer();
      const xlsxMime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      const litersTotal = ctx.rows.reduce((s, r) => s + (parseFloat(r.liters) || 0), 0);
      const costTotal = ctx.rows.reduce((s, r) => s + (parseFloat(r.total_cost) || 0), 0);
      const summary = `${ctx.rows.length} voucher rows · ${litersTotal.toFixed(2)} L · Rs. ${costTotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
      const labels = selectionLabels();

      await onPublishCompletedFile({
        title: `${ctx.title} (Excel)`,
        fileName: `${ctx.base}.xlsx`,
        module: 'voucher',
        format: 'xlsx',
        mimeType: xlsxMime,
        blob: new Blob([buf], { type: xlsxMime }),
        summary,
        reportDate: labels.dLabel,
      });

      const doc = buildAppliedPdf(ctx);
      const pdfBlob = doc.output('blob');
      await onPublishCompletedFile({
        title: `${ctx.title} (PDF)`,
        fileName: `${ctx.base}.pdf`,
        module: 'voucher',
        format: 'pdf',
        mimeType: 'application/pdf',
        blob: pdfBlob,
        summary,
        reportDate: labels.dLabel,
      });

      setNotice(`Published ${ctx.base}.xlsx and ${ctx.base}.pdf for Fuel Officer and Management.`);
    } catch (err: any) {
      setNotice(err?.message || 'Could not publish voucher files.');
    } finally {
      setPublishing(false);
    }
  };

  const downloadCsv = () => {
    const all = [...pendingRows, ...appliedRows];
    const header = CSV_COLS.map((c) => c.label).join(',');
    const escCsv = (v: unknown) => {
      const s = (v ?? '').toString();
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    const body = all.map((r) => CSV_COLS.map((c) => escCsv(r[c.key])).join(',')).join('\n');
    saveBlob(new Blob([header + '\n' + body], { type: 'text/csv;charset=utf-8;' }), 'fuel_transactions_analysis.csv');
  };

  const openPrintPreview = (scope: 'selection' | 'full') => {
    const ctx = exportContext(scope);
    if (!ctx.rows.length && !ctx.pending.length) {
      setNotice('No data loaded yet.');
      return;
    }
    const w = window.open('', '_blank');
    if (!w) {
      setNotice('Please allow pop-ups to open print preview.');
      return;
    }
    const renderTable = (rows: VoucherRow[], editable: boolean) => {
      const tl = rows.reduce((s, r) => s + (parseFloat(r.liters) || 0), 0);
      const tc = rows.reduce((s, r) => s + (parseFloat(r.total_cost) || 0), 0);
      let h = '<table><thead><tr>';
      APPLIED_COLS.forEach((c) => {
        const isNum = ['liters', 'price', 'total_cost', 'mileage'].includes(c.key);
        h += `<th class="${isNum ? 'num' : ''}">${esc(c.label)}</th>`;
      });
      h += '</tr></thead><tbody>';
      rows.forEach((r, idx) => {
        h += '<tr>';
        APPLIED_COLS.forEach((c) => {
          const isNum = ['liters', 'price', 'total_cost', 'mileage'].includes(c.key);
          if (editable && (c.key === 'meter_start' || c.key === 'meter_end')) {
            h += `<td class="num"><input class="meter-input" type="text" value="${esc(
              r[c.key] || ''
            )}" data-meter-row="${r.__id}" data-meter-field="${c.key}" oninput="window.sendMeterUpdate('${
              r.__id
            }','${c.key}',this.value,this)"></td>`;
          } else if (editable && c.key === 'shift') {
            const cur = r.shift || 'DAY';
            h += `<td><select class="shift-input" onchange="window.sendShiftUpdate('${r.__id}',this.value)"><option ${
              cur === 'DAY' ? 'selected' : ''
            }>DAY</option><option ${cur === 'NIGHT' ? 'selected' : ''}>NIGHT</option></select></td>`;
          } else if (c.key === 'mileage') {
            h += `<td class="num" data-mileage-for="${r.__id}">${esc(r.mileage ?? '')}</td>`;
          } else if (c.key === 'price') {
            h += `<td class="num">${esc(fmtPrice(r.price))}</td>`;
          } else if (c.key === 'total_cost') {
            h += `<td class="num">${esc(fmtCost(r.total_cost))}</td>`;
          } else if (c.key === '__serial') {
            h += `<td>${idx + 1}</td>`;
          } else {
            h += `<td class="${isNum ? 'num' : ''}">${esc(r[c.key as keyof VoucherRow] ?? '')}</td>`;
          }
        });
        h += '</tr>';
      });
      h += `</tbody><tfoot><tr><td colspan="5" style="text-align:right">TOTALS</td><td class="num">${fmt2(
        tl
      )}</td><td colspan="2"></td><td class="num">${fmt2(tc)}</td><td colspan="7"></td></tr></tfoot></table>`;
      return h;
    };

    const tl = ctx.rows.reduce((s, r) => s + (parseFloat(r.liters) || 0), 0);
    const tc = ctx.rows.reduce((s, r) => s + (parseFloat(r.total_cost) || 0), 0);
    let body = `<div class="banner"><b>Editable:</b> Meter Reading Start / End and SHIFT can be typed directly below and sync back automatically.</div>`;
    body += `<h1>${esc(ctx.title)}</h1><div class="meta">${esc(ctx.meta)}</div>`;
    if (ctx.rows.length) {
      body += renderTable(ctx.rows, true);
      body += `<div class="total-line">USAGE ${fmt2(tl)} L &nbsp;&nbsp; TOTAL Rs. ${fmt2(tc)}</div>`;
    }
    if (ctx.pending.length) {
      body += `<h2>Pending (Not Yet Applied to a Voucher)</h2>` + renderTable(ctx.pending, false);
    }
    body += `<div class="actions"><button onclick="window.print()">Print / Save PDF</button><button class="ghost" onclick="window.close()">Close</button></div>`;

    w.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>${esc(ctx.title)}</title>
    <style>
      body{font-family:Arial,Helvetica,sans-serif;color:#1b2430;padding:20px}
      h1{font-size:15px;text-align:center;margin:0 0 3px;text-transform:uppercase}
      .meta{text-align:center;font-size:11px;color:#555;margin-bottom:14px}
      h2{font-size:12px;margin:16px 0 6px;border-bottom:1px solid #999;padding-bottom:3px}
      table{width:100%;border-collapse:collapse;font-size:10px;margin-bottom:6px}
      th,td{border:1px solid #999;padding:4px 6px;text-align:left}
      th{background:#eee} td.num,th.num{text-align:right}
      tfoot td{font-weight:700;background:#f4f4f4;border-top:2px solid #333}
      .total-line{text-align:right;font-weight:700;font-size:12px;margin:6px 0 14px}
      .banner{background:#eef5f1;border:1px solid #9ab8a5;padding:8px 12px;border-radius:4px;font-size:11px;margin-bottom:12px}
      input.meter-input{font-size:10.5px;padding:2px 4px;border:1px solid #ccc;border-radius:3px;width:90px;text-align:right}
      select.shift-input{font-size:10.5px;padding:2px 4px;border:1px solid #ccc;border-radius:3px}
      .actions{text-align:center;margin:14px 0;display:flex;gap:8px;justify-content:center}
      .actions button{padding:7px 14px;border-radius:4px;border:1px solid #1F5143;background:#1F5143;color:#fff;cursor:pointer;font-weight:600}
      .actions button.ghost{background:#fff;color:#1b2430;border-color:#bbb}
      @media print{.actions,.banner{display:none}input.meter-input{border:none;padding:0}select.shift-input{border:none;appearance:none;-webkit-appearance:none}@page{margin:10mm;size:landscape}}
    </style></head><body>${body}
    <script>
      window.recalcMileage=function(inputEl,rowId){
        var doc=inputEl.ownerDocument;
        var sInput=doc.querySelector('[data-meter-row="'+rowId+'"][data-meter-field="meter_start"]');
        var eInput=doc.querySelector('[data-meter-row="'+rowId+'"][data-meter-field="meter_end"]');
        var out=doc.querySelector('[data-mileage-for="'+rowId+'"]');
        if(!sInput||!eInput||!out) return;
        var s=parseFloat(sInput.value), e=parseFloat(eInput.value);
        out.textContent=(!isNaN(s)&&!isNaN(e))?(e-s).toFixed(2):'';
      };
      window.sendMeterUpdate=function(rowId,field,value,inputEl){
        if(window.opener&&window.opener.handlePrintMeterInput){try{window.opener.handlePrintMeterInput(rowId,field,value);}catch(e){}}
        window.recalcMileage(inputEl,rowId);
      };
      window.sendShiftUpdate=function(rowId,value){
        if(window.opener&&window.opener.handlePrintShiftInput){try{window.opener.handlePrintShiftInput(rowId,value);}catch(e){}}
      };
    <\/script></body></html>`);
    w.document.close();
  };

  // Stats
  const missingAreaCount = pendingRows.filter(
    (r) => !isCommonVehicle(r.vehicle) && (!r.__areaConfirmed || !r.area)
  ).length;
  const needsSplitCount = pendingRows.filter((r) => isCommonVehicle(r.vehicle)).length;
  const readyCount = pendingRows.filter((r) => rowStatus(r) === 'amber').length;

  const visiblePending = useMemo(() => {
    const q = pendingSearch.trim().toLowerCase();
    if (!q) return pendingRows;
    return pendingRows.filter((r) =>
      [r.vehicle, r.usage_id, r.liters, r.total_cost, r.area, r.voucher, r.fuel_type, r.usage_date]
        .map((v) => String(v ?? '').toLowerCase())
        .some((v) => v.includes(q))
    );
  }, [pendingRows, pendingSearch]);

  const appliedTotals = useMemo(() => {
    const l = filteredAppliedRows.reduce((s, r) => s + (parseFloat(r.liters) || 0), 0);
    const c = filteredAppliedRows.reduce((s, r) => s + (parseFloat(r.total_cost) || 0), 0);
    return { l, c };
  }, [filteredAppliedRows]);

  const editMileagePreview = useMemo(() => {
    const s = parseFloat(editForm.meter_start);
    const e = parseFloat(editForm.meter_end);
    return !isNaN(s) && !isNaN(e) ? (e - s).toFixed(2) : '—';
  }, [editForm.meter_start, editForm.meter_end]);

  return (
    <div className="space-y-6 pb-24">
      <div className="flex flex-col md:flex-row md:items-end justify-between gap-4 border-b border-slate-200 pb-5">
        <div>
          <h1 className="text-xl font-bold text-slate-900">Fuel Voucher Analysis</h1>
          <p className="text-sm text-slate-600 mt-1">
            Confirm areas, split Common fuel cards, assign Voucher Numbers, and export audited voucher records.
          </p>
        </div>

        {/* Sub-navigation segmented control */}
        <div className="inline-flex items-center gap-1 p-1 bg-slate-200/70 rounded-md self-start">
          <button
            type="button"
            onClick={() => setSubTab('upload')}
            className={`px-3.5 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap ${
              subTab === 'upload' ? 'bg-white text-slate-900 shadow-xs' : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            1. Upload Report
          </button>
          <button
            type="button"
            onClick={() => setSubTab('pending')}
            className={`px-3.5 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap ${
              subTab === 'pending' ? 'bg-white text-slate-900 shadow-xs' : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            2. Pending ({pendingRows.length})
          </button>
          <button
            type="button"
            onClick={() => setSubTab('applied')}
            className={`px-3.5 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap ${
              subTab === 'applied' ? 'bg-white text-slate-900 shadow-xs' : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            3. Applied Vouchers ({appliedRows.length})
          </button>
        </div>
      </div>

      {notice && (
        <div className="px-4 py-2.5 bg-slate-100 border border-slate-300 rounded-md text-xs font-medium text-slate-800 flex items-center justify-between">
          <span>{notice}</span>
          <button type="button" onClick={() => setNotice(null)} className="text-slate-500 hover:text-slate-900">
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* Stats Strip when viewing Pending or Applied */}
      {(pendingRows.length > 0 || appliedRows.length > 0) && subTab !== 'upload' && (
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">
          <div className="bg-white border border-slate-200 rounded-md p-3.5">
            <div className="text-xs text-slate-500">Pending</div>
            <div className="font-data text-lg font-semibold text-slate-900 mt-0.5 tabular-nums">{pendingRows.length}</div>
          </div>
          <div className="bg-white border border-slate-200 rounded-md p-3.5">
            <div className="text-xs text-slate-500">Area Not Confirmed</div>
            <div className="font-data text-lg font-semibold text-red-700 mt-0.5 tabular-nums">{missingAreaCount}</div>
          </div>
          {needsSplitCount > 0 && (
            <div className="bg-white border border-slate-200 rounded-md p-3.5">
              <div className="text-xs text-slate-500">Needs Fuel Split</div>
              <div className="font-data text-lg font-semibold text-purple-700 mt-0.5 tabular-nums">{needsSplitCount}</div>
            </div>
          )}
          <div className="bg-white border border-slate-200 rounded-md p-3.5">
            <div className="text-xs text-slate-500">Ready to Select</div>
            <div className="font-data text-lg font-semibold text-amber-700 mt-0.5 tabular-nums">{readyCount}</div>
          </div>
          <div className="bg-white border border-slate-200 rounded-md p-3.5">
            <div className="text-xs text-slate-500">Applied to Voucher</div>
            <div className="font-data text-lg font-semibold text-emerald-700 mt-0.5 tabular-nums">{appliedRows.length}</div>
          </div>
        </div>
      )}

      {/* SUBTAB 1: UPLOAD & SETUP */}
      {subTab === 'upload' && (
        <div className="bg-white border border-slate-200 rounded-md p-6 space-y-5">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-slate-50 border border-slate-200 rounded-md p-4">
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">Your Name (created_by)</label>
              <input
                type="text"
                value={currentUser}
                onChange={(e) => setCurrentUser(e.target.value)}
                placeholder="e.g. G P S Abimana"
                className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md text-sm focus:outline-none focus:border-[#1F5143]"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">
                User ID (Usage_id — applied to every transaction)
              </label>
              <input
                type="text"
                value={userId}
                onChange={(e) => setUserId(e.target.value)}
                className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md text-sm font-data focus:outline-none focus:border-[#1F5143]"
              />
            </div>
          </div>

          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-3">
              <label className="inline-flex items-center gap-2 px-4 py-2 bg-[#1F5143] hover:bg-[#173D32] text-white text-xs font-medium rounded-md cursor-pointer transition-colors whitespace-nowrap">
                <Upload className="w-4 h-4" /> Upload Mobitel OneFuel Transaction Report PDF
                <input type="file" accept="application/pdf" className="hidden" onChange={handlePdfUpload} />
              </label>
              {canImportFromGps && (
                <button
                  type="button"
                  onClick={onImportFromGps}
                  className="inline-flex items-center gap-1.5 px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-medium rounded-md transition-colors whitespace-nowrap"
                >
                  <RefreshCw className="w-3.5 h-3.5" /> Import Mobitel Rows from GPS Reconciliation
                </button>
              )}
            </div>
            {(pendingRows.length > 0 || appliedRows.length > 0) && (
              <button
                type="button"
                onClick={onClearWorkspace}
                className="px-3.5 py-2 bg-white hover:bg-red-50 text-red-700 border border-red-200 text-xs font-medium rounded-md"
              >
                Clear Workspace
              </button>
            )}
          </div>

          {pdfStatus && <div className="text-xs text-slate-600">{pdfStatus}</div>}

          {uploadSummary && (
            <div
              className={`border rounded-md p-4 space-y-2 text-xs ${
                uploadSummary.bad.length === 0 ? 'border-emerald-300 bg-emerald-50/50' : 'border-red-300 bg-red-50/50'
              }`}
            >
              <div className="text-sm font-semibold text-slate-900">
                {uploadSummary.reportCount != null
                  ? `${uploadSummary.taken + uploadSummary.dups} of ${uploadSummary.reportCount} transactions read`
                  : `${uploadSummary.taken} transactions read`}
              </div>
              <div>
                <b>{uploadSummary.taken}</b> new transactions added
                {uploadSummary.dups ? ` · ${uploadSummary.dups} duplicates skipped` : ''}
              </div>
              {uploadSummary.reportLitres != null && (
                <div className="font-data">
                  Litres: {fmt2(uploadSummary.sumLitres)} L read vs {fmt2(uploadSummary.reportLitres)} L in report
                </div>
              )}
              {uploadSummary.reportAmount != null && (
                <div className="font-data">
                  Amount: Rs. {fmt2(uploadSummary.sumAmount)} read vs Rs. {fmt2(uploadSummary.reportAmount)} in report
                </div>
              )}
              <div className="text-slate-600">
                {Object.entries(uploadSummary.fuelCounts)
                  .map(([k, v]) => `${k} × ${v}`)
                  .join(' · ')}
              </div>
              <div className="pt-2">
                <button
                  type="button"
                  onClick={() => setSubTab('pending')}
                  className="px-4 py-2 bg-[#1F5143] hover:bg-[#173D32] text-white text-xs font-medium rounded-md"
                >
                  Continue → Pending Transactions
                </button>
              </div>
            </div>
          )}

          <p className="text-xs text-slate-500 leading-relaxed">
            Usage Date, Vehicle Number, Liters, Fuel Type, and Price are locked once read from the report. A mileage value of 3 in the report means the vehicle has no odometer and leaves Meter Reading End blank.
          </p>
        </div>
      )}

      {/* SUBTAB 2: PENDING TRANSACTIONS */}
      {subTab === 'pending' && (
        <div className="bg-white border border-slate-200 rounded-md overflow-hidden">
          <div className="p-4 border-b border-slate-200 space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={tickAllSuggestedAreas}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-300 text-xs font-semibold rounded-md whitespace-nowrap"
                >
                  <Check className="w-3.5 h-3.5" /> Tick All Suggested Areas
                </button>
                <button
                  type="button"
                  onClick={selectAllReady}
                  className="px-3 py-2 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 text-xs font-medium rounded-md whitespace-nowrap"
                >
                  Select All Ready
                </button>
                <button
                  type="button"
                  onClick={clearSelection}
                  className="px-3 py-2 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 text-xs font-medium rounded-md whitespace-nowrap"
                >
                  Clear Selection
                </button>
                <button
                  type="button"
                  onClick={downloadCsv}
                  className="px-3 py-2 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 text-xs font-medium rounded-md whitespace-nowrap"
                >
                  Download CSV
                </button>
              </div>

              <div className="flex items-center gap-2">
                <div className="relative">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" />
                  <input
                    type="text"
                    value={pendingSearch}
                    onChange={(e) => setPendingSearch(e.target.value)}
                    placeholder="Vehicle No or Liters…"
                    className="pl-8 pr-3 py-1.5 border border-slate-300 rounded-md text-xs w-56 focus:outline-none focus:border-[#1F5143]"
                  />
                </div>
                {pendingSearch && (
                  <button
                    type="button"
                    onClick={() => setPendingSearch('')}
                    className="px-2.5 py-1.5 text-xs text-slate-600 hover:text-slate-900"
                  >
                    Clear
                  </button>
                )}
              </div>
            </div>
          </div>

          <div className="overflow-x-auto max-h-[64vh]">
            {visiblePending.length === 0 ? (
              <div className="p-12 text-center text-xs text-slate-400">
                {pendingRows.length
                  ? 'No pending transactions match your search.'
                  : 'No pending transactions yet — upload a report on the Upload tab.'}
              </div>
            ) : (
              <table className="w-full text-xs text-left border-collapse">
                <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 sticky top-0 z-10">
                  <tr>
                    <th className="px-3 py-2.5 font-semibold w-12">Select</th>
                    <th className="px-3 py-2.5 font-semibold w-28">Status</th>
                    <th className="px-3 py-2.5 font-semibold">Vehicle Number</th>
                    <th className="px-3 py-2.5 font-semibold">Usage Date</th>
                    <th className="px-3 py-2.5 font-semibold text-right">Liters Used</th>
                    <th className="px-3 py-2.5 font-semibold">Fuel Type</th>
                    <th className="px-3 py-2.5 font-semibold text-right">Price/L</th>
                    <th className="px-3 py-2.5 font-semibold">Meter Start</th>
                    <th className="px-3 py-2.5 font-semibold">Meter End</th>
                    <th className="px-3 py-2.5 font-semibold min-w-[260px]">Area</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {visiblePending.map((r) => {
                    const st = rowStatus(r);
                    const canSelect = st === 'amber';
                    const sug = suggestedArea(r.vehicle);
                    const areaDisplay = r.area || sug || '';
                    const suggestions = areaSuggestions(r.vehicle, areaDisplay);
                    const splitOpen = openSplitId === r.__id;
                    const draft = r.__splitDraft || [];
                    const splitTarget = parseFloat(r.liters) || 0;
                    const splitSum = draft.reduce((s, l) => s + (parseFloat(l.liters) || 0), 0);
                    const splitBalanced = Math.abs(splitSum - splitTarget) < 0.01 && draft.length > 0;

                    return (
                      <React.Fragment key={r.__id}>
                        <tr className={r.__selected ? 'bg-emerald-50/60' : 'hover:bg-slate-50'}>
                          <td className="px-3 py-2">
                            <input
                              type="checkbox"
                              checked={r.__selected}
                              disabled={!canSelect}
                              onChange={(e) => updatePendingField(r.__id, { __selected: e.target.checked })}
                              title={
                                canSelect
                                  ? 'Select transaction'
                                  : st === 'purple'
                                    ? 'Split this Common row first'
                                    : 'Confirm Area first'
                              }
                            />
                          </td>
                          <td className="px-3 py-2 font-medium whitespace-nowrap">
                            {st === 'red' ? (
                              <span className="text-red-700">Area needed</span>
                            ) : st === 'purple' ? (
                              <span className="text-purple-700">Split usage</span>
                            ) : (
                              <span className="text-amber-700">Ready</span>
                            )}
                          </td>
                          <td className="px-3 py-2 font-data font-semibold text-slate-900 whitespace-nowrap">
                            <div className="flex items-center gap-2">
                              <span>{r.vehicle}</span>
                              {isCommonVehicle(r.vehicle) && (
                                <button
                                  type="button"
                                  onClick={() => toggleSplitDialog(r.__id)}
                                  className="px-2 py-0.5 border border-purple-300 bg-purple-50 text-purple-800 rounded text-[11px] font-medium"
                                >
                                  {splitOpen ? 'Hide Split' : 'Split Usage'}
                                </button>
                              )}
                              {r.__isSplitChild && (
                                <span className="text-[11px] font-sans font-normal text-purple-700">split line</span>
                              )}
                            </div>
                          </td>
                          <td className="px-3 py-2 font-data whitespace-nowrap">{r.usage_date}</td>
                          <td className="px-3 py-2 font-data text-right tabular-nums">{r.liters}</td>
                          <td className="px-3 py-2 whitespace-nowrap">{r.fuel_type || '—'}</td>
                          <td className="px-3 py-2 font-data text-right tabular-nums">{fmtPrice(r.price) || '—'}</td>
                          <td className="px-2 py-1.5">
                            <input
                              type="text"
                              value={r.meter_start}
                              onChange={(e) => updatePendingField(r.__id, { meter_start: e.target.value })}
                              placeholder="Start"
                              className="w-24 px-2 py-1 border border-slate-200 rounded text-xs font-data"
                            />
                          </td>
                          <td className="px-2 py-1.5">
                            <input
                              type="text"
                              value={r.meter_end}
                              onChange={(e) => updatePendingField(r.__id, { meter_end: e.target.value })}
                              placeholder="End"
                              className="w-24 px-2 py-1 border border-slate-200 rounded text-xs font-data"
                            />
                          </td>
                          <td className="px-2 py-1.5">
                            <div className="relative flex items-center gap-1.5">
                              <div className="relative flex-1">
                                <input
                                  type="text"
                                  value={areaDisplay}
                                  onFocus={() => setActiveSuggestRowId(r.__id)}
                                  onChange={(e) =>
                                    updatePendingField(r.__id, {
                                      area: e.target.value,
                                      __areaConfirmed: false,
                                      __selected: false,
                                    })
                                  }
                                  onKeyDown={(e) => {
                                    if (e.key === 'Enter' && areaDisplay.trim()) {
                                      confirmArea(r.__id, areaDisplay);
                                    } else if (e.key === 'Escape') {
                                      setActiveSuggestRowId(null);
                                    }
                                  }}
                                  placeholder="Type Area…"
                                  className={`w-full px-2.5 py-1 border rounded text-xs ${
                                    r.__areaConfirmed
                                      ? 'border-emerald-400 bg-emerald-50/40'
                                      : 'border-slate-300 bg-white'
                                  }`}
                                />
                                {activeSuggestRowId === r.__id && suggestions.length > 0 && (
                                  <div className="absolute left-0 right-0 top-full mt-0.5 bg-white border border-slate-200 rounded shadow-lg max-h-44 overflow-y-auto z-30">
                                    {suggestions.map((a) => (
                                      <button
                                        key={a}
                                        type="button"
                                        onMouseDown={(ev) => {
                                          ev.preventDefault();
                                          confirmArea(r.__id, a);
                                        }}
                                        className="block w-full text-left px-2.5 py-1.5 text-xs hover:bg-slate-100"
                                      >
                                        {a}
                                      </button>
                                    ))}
                                  </div>
                                )}
                              </div>
                              <button
                                type="button"
                                onClick={() => confirmArea(r.__id, areaDisplay)}
                                className={`px-2 py-1 rounded text-xs font-semibold border whitespace-nowrap ${
                                  r.__areaConfirmed
                                    ? 'bg-emerald-50 border-emerald-300 text-emerald-800'
                                    : 'bg-slate-50 border-slate-300 text-slate-600 hover:bg-slate-100'
                                }`}
                              >
                                ✓ OK
                              </button>
                            </div>
                          </td>
                        </tr>

                        {/* Inline Common Split Panel */}
                        {isCommonVehicle(r.vehicle) && splitOpen && (
                          <tr className="bg-slate-50/80">
                            <td colSpan={10} className="p-4">
                              <div className="max-w-4xl border border-slate-200 bg-white rounded-md p-4 space-y-3">
                                <div className="flex items-center justify-between">
                                  <div className="text-xs font-semibold text-slate-900">
                                    Split “{r.vehicle}” — {r.liters} L pumped
                                  </div>
                                  <button
                                    type="button"
                                    onClick={() => setOpenSplitId(null)}
                                    className="text-xs text-slate-500 hover:text-slate-900"
                                  >
                                    Close
                                  </button>
                                </div>

                                <table className="w-full text-xs border-collapse">
                                  <thead>
                                    <tr className="border-b border-slate-200 text-slate-600">
                                      <th className="py-1.5 text-left">Used by (Category or Vehicle)</th>
                                      <th className="py-1.5 text-right w-28">Liters</th>
                                      <th className="py-1.5 text-left w-28">Meter Start</th>
                                      <th className="py-1.5 text-left w-28">Meter End</th>
                                      <th className="py-1.5 w-20" />
                                    </tr>
                                  </thead>
                                  <tbody className="divide-y divide-slate-100">
                                    {draft.map((line, idx) => {
                                      const cats = commonCategoriesFor(r.vehicle);
                                      const vehs = commonVehiclesFor(r.vehicle);
                                      const known = new Set([...cats, ...vehs]);
                                      const isOther = !!line.label && !known.has(line.label);
                                      return (
                                        <tr key={idx}>
                                          <td className="py-1.5 pr-2 space-y-1">
                                            <select
                                              value={isOther ? '__other__' : line.label}
                                              onChange={(e) => {
                                                const v = e.target.value;
                                                updateSplitDraft(r.__id, (d) =>
                                                  d.map((item, i) =>
                                                    i === idx
                                                      ? { ...item, label: v === '__other__' ? 'Other' : v }
                                                      : item
                                                  )
                                                );
                                              }}
                                              className="w-full px-2 py-1 border border-slate-300 rounded text-xs bg-white"
                                            >
                                              <option value="">— Select —</option>
                                              <optgroup label="Category">
                                                {cats.map((c) => (
                                                  <option key={c} value={c}>
                                                    {c}
                                                  </option>
                                                ))}
                                              </optgroup>
                                              {vehs.length > 0 && (
                                                <optgroup label="Staff Bike / Vehicle">
                                                  {vehs.map((v) => (
                                                    <option key={v} value={v}>
                                                      {v}
                                                    </option>
                                                  ))}
                                                </optgroup>
                                              )}
                                              <option value="__other__">Other (type manually)</option>
                                            </select>
                                            {isOther && (
                                              <input
                                                type="text"
                                                value={line.label === 'Other' ? '' : line.label}
                                                onChange={(e) =>
                                                  updateSplitDraft(r.__id, (d) =>
                                                    d.map((item, i) =>
                                                      i === idx ? { ...item, label: e.target.value } : item
                                                    )
                                                  )
                                                }
                                                placeholder="Type destination…"
                                                className="w-full px-2 py-1 border border-slate-300 rounded text-xs"
                                              />
                                            )}
                                          </td>
                                          <td className="py-1.5 pr-2">
                                            <input
                                              type="number"
                                              step="0.01"
                                              value={line.liters}
                                              onChange={(e) =>
                                                updateSplitDraft(r.__id, (d) =>
                                                  d.map((item, i) =>
                                                    i === idx ? { ...item, liters: e.target.value } : item
                                                  )
                                                )
                                              }
                                              className="w-full px-2 py-1 border border-slate-300 rounded text-xs text-right font-data"
                                            />
                                          </td>
                                          <td className="py-1.5 pr-2">
                                            <input
                                              type="text"
                                              value={line.meter_start}
                                              onChange={(e) =>
                                                updateSplitDraft(r.__id, (d) =>
                                                  d.map((item, i) =>
                                                    i === idx ? { ...item, meter_start: e.target.value } : item
                                                  )
                                                )
                                              }
                                              className="w-full px-2 py-1 border border-slate-300 rounded text-xs font-data"
                                            />
                                          </td>
                                          <td className="py-1.5 pr-2">
                                            <input
                                              type="text"
                                              value={line.meter_end}
                                              onChange={(e) =>
                                                updateSplitDraft(r.__id, (d) =>
                                                  d.map((item, i) =>
                                                    i === idx ? { ...item, meter_end: e.target.value } : item
                                                  )
                                                )
                                              }
                                              className="w-full px-2 py-1 border border-slate-300 rounded text-xs font-data"
                                            />
                                          </td>
                                          <td className="py-1.5 text-right">
                                            <button
                                              type="button"
                                              onClick={() =>
                                                updateSplitDraft(r.__id, (d) => d.filter((_, i) => i !== idx))
                                              }
                                              className="text-red-700 hover:underline"
                                            >
                                              Remove
                                            </button>
                                          </td>
                                        </tr>
                                      );
                                    })}
                                  </tbody>
                                </table>

                                <div className="flex flex-wrap items-center justify-between gap-3 pt-2 border-t border-slate-100">
                                  <div className="flex items-center gap-3">
                                    <button
                                      type="button"
                                      onClick={() =>
                                        updateSplitDraft(r.__id, (d) => [
                                          ...d,
                                          { label: '', liters: '', meter_start: '', meter_end: '' },
                                        ])
                                      }
                                      className="inline-flex items-center gap-1 px-2.5 py-1 border border-slate-300 rounded text-xs font-medium hover:bg-slate-50"
                                    >
                                      <Plus className="w-3 h-3" /> Add Line
                                    </button>
                                    <span
                                      className={`font-data text-xs font-semibold ${
                                        splitBalanced ? 'text-emerald-700' : 'text-red-700'
                                      }`}
                                    >
                                      Split {splitSum.toFixed(2)} / {splitTarget.toFixed(2)} L{' '}
                                      {splitBalanced ? '(Balanced)' : ''}
                                    </span>
                                  </div>
                                  <button
                                    type="button"
                                    disabled={!splitBalanced}
                                    onClick={() => confirmSplit(r)}
                                    className="px-3.5 py-1.5 bg-[#1F5143] disabled:opacity-40 text-white text-xs font-medium rounded-md"
                                  >
                                    Confirm Split → Create Rows
                                  </button>
                                </div>
                              </div>
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>
        </div>
      )}

      {/* SUBTAB 3: APPLIED VOUCHERS */}
      {subTab === 'applied' && (
        <div className="space-y-4">
          {appliedRows.length === 0 ? (
            <div className="bg-white border border-slate-200 rounded-md p-12 text-center text-xs text-slate-400">
              No applied vouchers yet. Select ready transactions on the Pending tab and apply a Voucher No.
            </div>
          ) : (
            <>
              {/* Date & Voucher Filter Grid */}
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
                <div className="lg:col-span-5 bg-white border border-slate-200 rounded-md overflow-hidden">
                  <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between text-xs">
                    <span className="font-semibold text-slate-800">
                      Usage Date ({selectedAppliedDays.length || 'All'} / {appliedDayKeysList.length})
                    </span>
                    <button
                      type="button"
                      onClick={() => setSelectedAppliedDays([])}
                      className="text-slate-600 hover:text-slate-900 font-medium"
                    >
                      Clear
                    </button>
                  </div>
                  <div className="p-3 flex flex-wrap gap-1.5 max-h-44 overflow-y-auto">
                    {appliedDayKeysList.map((k) => {
                      const active = selectedAppliedDays.includes(k);
                      const sample = appliedRows.find((r) => dayKey(r.usage_date) === k);
                      return (
                        <button
                          key={k}
                          type="button"
                          onClick={() =>
                            setSelectedAppliedDays((prev) =>
                              prev.includes(k) ? prev.filter((x) => x !== k) : [...prev, k]
                            )
                          }
                          className={`px-2.5 py-1 rounded text-xs font-data font-medium border transition-colors ${
                            active
                              ? 'bg-[#1F5143] text-white border-[#1F5143]'
                              : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
                          }`}
                        >
                          {sample ? dayLabel(sample.usage_date) : k}
                        </button>
                      );
                    })}
                  </div>
                </div>

                <div className="lg:col-span-7 bg-white border border-slate-200 rounded-md overflow-hidden">
                  <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between text-xs">
                    <span className="font-semibold text-slate-800">
                      Voucher No ({selectedAppliedVouchers.length || 'All'} / {allVouchersList.length})
                    </span>
                    <button
                      type="button"
                      onClick={() => setSelectedAppliedVouchers([])}
                      className="text-slate-600 hover:text-slate-900 font-medium"
                    >
                      Clear
                    </button>
                  </div>
                  <div className="p-3 flex flex-wrap gap-1.5 max-h-44 overflow-y-auto">
                    {allVouchersList.map((v) => {
                      const active = selectedAppliedVouchers.includes(v);
                      const matches = vouchersMatchingDays.has(v);
                      return (
                        <button
                          key={v}
                          type="button"
                          disabled={!matches && !active}
                          onClick={() =>
                            setSelectedAppliedVouchers((prev) =>
                              prev.includes(v) ? prev.filter((x) => x !== v) : [...prev, v]
                            )
                          }
                          className={`px-2.5 py-1 rounded text-xs font-data font-medium border transition-colors ${
                            active
                              ? 'bg-[#1F5143] text-white border-[#1F5143]'
                              : matches
                                ? 'bg-slate-50 text-slate-800 border-slate-300 hover:bg-slate-100'
                                : 'bg-slate-50/40 text-slate-300 border-slate-200 cursor-not-allowed'
                          }`}
                        >
                          {v}
                        </button>
                      );
                    })}
                  </div>
                </div>
              </div>

              {/* Search & Export Toolbar */}
              <div className="bg-white border border-slate-200 rounded-md p-4 flex flex-wrap items-center justify-between gap-3">
                <div className="relative">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-2.5" />
                  <input
                    type="text"
                    value={appliedSearch}
                    onChange={(e) => setAppliedSearch(e.target.value)}
                    placeholder="Search Vehicle / Voucher / Area…"
                    className="pl-8 pr-3 py-1.5 border border-slate-300 rounded-md text-xs w-64 focus:outline-none focus:border-[#1F5143]"
                  />
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={() => openPrintPreview('selection')}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-[#1F5143] hover:bg-[#173D32] text-white text-xs font-medium rounded-md whitespace-nowrap"
                  >
                    <Printer className="w-3.5 h-3.5" /> Print Preview (Selection)
                  </button>
                  <button
                    type="button"
                    onClick={() => downloadAppliedExcel('selection')}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-white hover:bg-slate-50 text-slate-800 border border-slate-300 text-xs font-medium rounded-md whitespace-nowrap"
                  >
                    <Download className="w-3.5 h-3.5" /> Excel (Selection)
                  </button>
                  <button
                    type="button"
                    onClick={() => downloadAppliedPdf('selection')}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-white hover:bg-slate-50 text-slate-800 border border-slate-300 text-xs font-medium rounded-md whitespace-nowrap"
                  >
                    <Download className="w-3.5 h-3.5" /> PDF (Selection)
                  </button>
                  <button
                    type="button"
                    onClick={() => downloadAppliedExcel('full')}
                    className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-medium rounded-md whitespace-nowrap"
                  >
                    Excel (Full)
                  </button>
                  <button
                    type="button"
                    onClick={() => downloadAppliedPdf('full')}
                    className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-medium rounded-md whitespace-nowrap"
                  >
                    PDF (Full)
                  </button>
                  {onPublishCompletedFile && (
                    <button
                      type="button"
                      disabled={publishing}
                      onClick={() => void publishVoucherFiles('selection')}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-900 hover:bg-slate-800 text-white text-xs font-medium rounded-md whitespace-nowrap disabled:opacity-60"
                    >
                      <Send className="w-3.5 h-3.5" />{' '}
                      {publishing ? 'Publishing...' : 'Publish for Fuel Officer'}
                    </button>
                  )}
                </div>
              </div>
              {managementComments.length > 0 && (
                <div className="bg-white border border-slate-200 rounded-md p-4 space-y-1.5">
                  <div className="text-xs font-semibold text-slate-800">
                    Management Comments ({managementComments.length})
                  </div>
                  {managementComments.slice(0, 3).map((c) => (
                    <div key={c.id} className="text-xs text-slate-600 flex flex-wrap gap-2">
                      <span className="font-semibold text-slate-900">{c.author}</span>
                      <span>· {c.status}</span>
                      {c.reportDate && <span className="font-data">· {c.reportDate}</span>}
                      <span>— {c.comment}</span>
                    </div>
                  ))}
                </div>
              )}

              {/* 16-Column Applied Ledger Table */}
              <div className="bg-white border border-slate-200 rounded-md overflow-hidden">
                <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-2 text-xs">
                  <span className="font-semibold text-slate-900">
                    Saved Voucher Record ({visibleAppliedRows.length} rows)
                  </span>
                  <span className="font-data font-semibold text-slate-800 tabular-nums">
                    USAGE {fmt2(appliedTotals.l)} L · TOTAL Rs. {fmt2(appliedTotals.c)}
                  </span>
                </div>
                <div className="overflow-x-auto max-h-[60vh]">
                  <table className="w-full text-xs text-left border-collapse min-w-[1400px]">
                    <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 sticky top-0 z-10">
                      <tr>
                        {APPLIED_COLS.map((c) => {
                          const isNum = ['liters', 'price', 'total_cost', 'mileage'].includes(c.key);
                          return (
                            <th
                              key={c.key}
                              className={`px-3 py-2.5 font-semibold whitespace-nowrap ${isNum ? 'text-right' : ''}`}
                            >
                              {c.label}
                            </th>
                          );
                        })}
                        <th className="px-3 py-2.5 font-semibold text-right whitespace-nowrap">Actions</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {visibleAppliedRows.map((r) => (
                        <tr key={r.__id} className="hover:bg-slate-50">
                          {APPLIED_COLS.map((c) => {
                            const isNum = ['liters', 'price', 'total_cost', 'mileage'].includes(c.key);
                            let val: React.ReactNode;
                            if (c.key === '__serial') val = r.__serial ?? '';
                            else if (c.key === 'shift')
                              val = (
                                <span
                                  className={`font-medium ${
                                    r.shift === 'DAY' ? 'text-blue-700' : 'text-purple-700'
                                  }`}
                                >
                                  {r.shift}
                                </span>
                              );
                            else if (c.key === 'price') val = fmtPrice(r.price);
                            else if (c.key === 'total_cost') val = fmtCost(r.total_cost);
                            else val = String(r[c.key as keyof VoucherRow] ?? '');

                            return (
                              <td
                                key={c.key}
                                className={`px-3 py-2 whitespace-nowrap font-data tabular-nums ${
                                  isNum ? 'text-right' : ''
                                } ${c.key === 'vehicle' ? 'font-semibold text-slate-900' : ''}`}
                              >
                                {val}
                              </td>
                            );
                          })}
                          <td className="px-3 py-2 text-right whitespace-nowrap space-x-1">
                            <button
                              type="button"
                              onClick={() => openEditModal(r)}
                              className="inline-flex items-center gap-1 px-2 py-1 border border-slate-300 rounded text-slate-700 hover:bg-slate-100"
                            >
                              <Edit3 className="w-3 h-3" /> Edit
                            </button>
                            <button
                              type="button"
                              onClick={() => deleteApplied(r.__id)}
                              className="inline-flex items-center gap-1 px-2 py-1 border border-red-200 rounded text-red-700 hover:bg-red-50"
                            >
                              <Trash2 className="w-3 h-3" /> Delete
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot className="bg-slate-100 border-t border-slate-300 font-semibold text-slate-900">
                      <tr>
                        <td colSpan={5} className="px-3 py-2.5 text-right">
                          TOTALS
                        </td>
                        <td className="px-3 py-2.5 text-right font-data tabular-nums">{fmt2(appliedTotals.l)}</td>
                        <td colSpan={2} />
                        <td className="px-3 py-2.5 text-right font-data tabular-nums">{fmt2(appliedTotals.c)}</td>
                        <td colSpan={8} />
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </div>
            </>
          )}
        </div>
      )}

      {/* Sticky Bottom Bulk Voucher Bar (only when pending rows are ticked) */}
      {subTab === 'pending' && selectedPending.length > 0 && (
        <div className="fixed bottom-0 left-0 right-0 lg:left-64 bg-slate-900 text-white px-6 py-3.5 border-t-2 border-emerald-400 shadow-2xl z-40 flex flex-wrap items-center justify-between gap-4">
          <div className="space-y-0.5">
            <div className="font-data text-sm font-semibold tabular-nums">
              {selectedPending.length} selected · USAGE {selectedLitersSum.toFixed(2)} L
            </div>
            {existingVoucherMatches.length > 0 && (
              <div className="text-xs text-emerald-300">
                Voucher already holds {existingVoucherMatches.length} record(s) (
                {existingVoucherMatches[0].fuel_type || 'fuel'} @ Rs.{existingVoucherMatches[0].price || '—'})
              </div>
            )}
          </div>
          <div className="flex flex-wrap items-center gap-3">
            <input
              type="text"
              value={bulkVoucher}
              onChange={(e) => setBulkVoucher(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && applyVoucher()}
              placeholder="Voucher No (e.g. 8484)"
              className="px-3 py-2 bg-slate-800 border border-slate-600 rounded text-sm font-data text-white placeholder-slate-400 focus:outline-none focus:border-emerald-400 w-48"
            />
            <button
              type="button"
              onClick={clearSelection}
              className="px-3.5 py-2 border border-slate-600 text-slate-200 hover:bg-slate-800 rounded text-xs font-medium"
            >
              Clear
            </button>
            <button
              type="button"
              onClick={applyVoucher}
              className="px-4 py-2 bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-semibold rounded text-xs whitespace-nowrap"
            >
              Apply Voucher &amp; Move
            </button>
          </div>
        </div>
      )}

      {/* Edit Applied Record Modal */}
      {editingId && (
        <div className="fixed inset-0 bg-slate-900/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white border border-slate-200 rounded-md w-full max-w-md p-6 space-y-4 shadow-xl">
            <h3 className="text-base font-semibold text-slate-900">Edit Applied Record</h3>
            <div className="space-y-3 text-xs">
              <div>
                <label className="block font-medium text-slate-600 mb-1">Voucher No</label>
                <input
                  type="text"
                  value={editForm.voucher}
                  onChange={(e) => setEditForm((f) => ({ ...f, voucher: e.target.value }))}
                  className="w-full px-3 py-2 border border-slate-300 rounded font-data"
                />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-medium text-slate-600 mb-1">Fuel Type</label>
                  <select
                    value={editForm.fuel_type}
                    onChange={(e) => setEditForm((f) => ({ ...f, fuel_type: e.target.value }))}
                    className="w-full px-3 py-2 border border-slate-300 rounded bg-white"
                  >
                    <option value="">— Select —</option>
                    <option>Diesel</option>
                    <option>Petrol 92 Octan</option>
                    <option>Petrol 95 Octan</option>
                    <option>Kerosine Oil</option>
                    <option>Super Diesel</option>
                  </select>
                </div>
                <div>
                  <label className="block font-medium text-slate-600 mb-1">SHIFT</label>
                  <select
                    value={editForm.shift}
                    onChange={(e) => setEditForm((f) => ({ ...f, shift: e.target.value as 'DAY' | 'NIGHT' }))}
                    className="w-full px-3 py-2 border border-slate-300 rounded bg-white"
                  >
                    <option value="DAY">DAY</option>
                    <option value="NIGHT">NIGHT</option>
                  </select>
                </div>
              </div>
              <div>
                <label className="block font-medium text-slate-600 mb-1">Price Per Liter</label>
                <input
                  type="number"
                  step="0.01"
                  value={editForm.price}
                  onChange={(e) => setEditForm((f) => ({ ...f, price: e.target.value }))}
                  className="w-full px-3 py-2 border border-slate-300 rounded font-data"
                />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-medium text-slate-600 mb-1">Meter Reading Start</label>
                  <input
                    type="text"
                    value={editForm.meter_start}
                    onChange={(e) => setEditForm((f) => ({ ...f, meter_start: e.target.value }))}
                    className="w-full px-3 py-2 border border-slate-300 rounded font-data"
                  />
                </div>
                <div>
                  <label className="block font-medium text-slate-600 mb-1">Meter Reading End</label>
                  <input
                    type="text"
                    value={editForm.meter_end}
                    onChange={(e) => setEditForm((f) => ({ ...f, meter_end: e.target.value }))}
                    className="w-full px-3 py-2 border border-slate-300 rounded font-data"
                  />
                </div>
              </div>
              <div className="px-3 py-2 bg-slate-50 border border-slate-200 rounded font-data font-semibold text-emerald-800">
                KMS (End − Start): {editMileagePreview}
              </div>
            </div>
            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setEditingId(null)}
                className="px-4 py-2 border border-slate-300 rounded text-xs font-medium text-slate-700 hover:bg-slate-50"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={saveEditModal}
                className="px-4 py-2 bg-[#1F5143] hover:bg-[#173D32] text-white rounded text-xs font-medium"
              >
                Save Changes
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
