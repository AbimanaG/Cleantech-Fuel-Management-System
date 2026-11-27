import React, { useMemo, useState } from 'react';
import { Upload, Download, Printer, Trash2, RefreshCw, AlertTriangle, Send } from 'lucide-react';
import type { CompletedFileModule, ConsolidatedDay, ManagementComment } from '../types/fuel';
import { extractPdfText } from '../utils/pdfReader';
import {
  buildConsolidatedWorkbook,
  buildMatrices,
  buildModel,
  f2,
  f2s,
  parseDailyReport,
} from '../utils/consolidationParser';
import { saveBlob } from '../utils/voucherParser';
import { esc } from '../utils/gpsMatcherParsers';

interface ConsolidationViewProps {
  days: Record<string, ConsolidatedDay>;
  onUpdateDays: (next: Record<string, ConsolidatedDay>) => void;
  canSyncFromGps: boolean;
  onSyncFromGps: () => void;
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

export const ConsolidationView: React.FC<ConsolidationViewProps> = ({
  days,
  onUpdateDays,
  canSyncFromGps,
  onSyncFromGps,
  onPublishCompletedFile,
  managementComments = [],
}) => {
  const [publishing, setPublishing] = useState(false);
  const [status, setStatus] = useState<string>('Upload daily reconciliation PDFs or sync directly from GPS Reconciliation.');
  const [rebuildMultiples, setRebuildMultiples] = useState(false);

  const sortedDates = useMemo(() => Object.keys(days).sort(), [days]);
  const model = useMemo(() => buildModel(days, { multiples: rebuildMultiples }), [days, rebuildMultiples]);
  const matrices = useMemo(() => buildMatrices(model), [model]);

  const handleUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (!files.length) return;
    const msgs: string[] = [];
    const next = { ...days };

    for (let i = 0; i < files.length; i++) {
      setStatus(`Reading ${i + 1}/${files.length}: ${files[i].name}…`);
      try {
        const text = await extractPdfText(files[i]);
        const day = parseDailyReport(text, files[i].name);
        if ('error' in day) {
          msgs.push(`✖ ${files[i].name}: ${day.error}`);
          continue;
        }
        if (next[day.date]) msgs.push(`↻ ${day.date}: replaced earlier file`);
        next[day.date] = day;
      } catch (err: any) {
        msgs.push(`✖ ${files[i].name}: ${err.message}`);
      }
    }
    e.target.value = '';
    onUpdateDays(next);
    const count = Object.keys(next).length;
    setStatus(`${count} daily report(s) loaded.` + (msgs.length ? '  ' + msgs.join(' | ') : ''));
  };

  const removeDay = (d: string) => {
    const next = { ...days };
    delete next[d];
    onUpdateDays(next);
  };

  const clearAll = () => {
    onUpdateDays({});
    setStatus('Cleared all consolidated days.');
  };

  const downloadXlsx = async () => {
    if (!sortedDates.length) return;
    const wb = await buildConsolidatedWorkbook(model);
    const buf = await wb.xlsx.writeBuffer();
    saveBlob(
      new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      `Fuel_Reconciliation_${sortedDates[0]}_to_${sortedDates[sortedDates.length - 1]}.xlsx`
    );
  };

  const publishXlsx = async () => {
    if (!sortedDates.length || !onPublishCompletedFile) return;
    setPublishing(true);
    try {
      const wb = await buildConsolidatedWorkbook(model);
      const buf = await wb.xlsx.writeBuffer();
      const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      const blob = new Blob([buf], { type: mimeType });
      const range =
        sortedDates.length === 1
          ? sortedDates[0]
          : `${sortedDates[0]} to ${sortedDates[sortedDates.length - 1]}`;
      const totalL = model.s1.reduce((s, x) => s + x.total, 0);
      await onPublishCompletedFile({
        title: `Multi-Day Fuel Consolidation (${range})`,
        fileName: `Fuel_Reconciliation_${sortedDates[0]}_to_${sortedDates[sortedDates.length - 1]}.xlsx`,
        module: 'consolidation',
        format: 'xlsx',
        mimeType,
        blob,
        summary: `${sortedDates.length} day(s) · ${totalL.toFixed(2)} L · ${model.s2.length} variance`,
        reportDate: range,
      });
      setStatus(`Published completed Excel workbook (${range}) for Fuel Officer & Management.`);
    } catch (err: any) {
      setStatus(err?.message || 'Failed to publish completed Excel workbook.');
    } finally {
      setPublishing(false);
    }
  };

  const downloadPdf = () => {
    if (!sortedDates.length) return;
    const w = window.open('', '_blank');
    if (!w) {
      setStatus('Please allow pop-ups to open the print PDF dialog.');
      return;
    }
    const sm = (k: 'diesel' | 'petrol' | 'superD' | 'total' | 'amt') =>
      model.s1.reduce((s, x) => s + x[k], 0);
    const c2 = model.s2.reduce((s, x) => s + x.cardL, 0);
    const g2 = model.s2.reduce((s, x) => s + x.gpsL, 0);
    const l3 = model.s3.reduce((s, x) => s + x.litres, 0);

    const s1Rows = model.s1
      .map(
        (x) =>
          `<tr><td><b>${x.date}</b></td><td class="num">${f2(x.diesel)}</td><td class="num">${f2(
            x.petrol
          )}</td><td class="num">${f2(x.superD)}</td><td class="num">${f2(x.total)}</td><td class="num">${f2(
            x.amt
          )}</td></tr>`
      )
      .join('');

    const s2Rows =
      model.s2
        .map((x) => {
          const v = x.gpsL - x.cardL;
          return `<tr class="${x.unusual ? 'unu' : ''}"><td>${x.date} ${x.time}</td><td><b>${esc(
            x.vehicle
          )}</b></td><td>${esc(x.type)}</td><td class="num">${f2(x.cardL)}</td><td class="num">${f2(
            x.gpsL
          )}</td><td class="num">${f2s(v)}</td><td class="num">${
            x.cardL ? ((v / x.cardL) * 100 > 0 ? '+' : '') + ((v / x.cardL) * 100).toFixed(1) : '-'
          }%</td><td></td></tr>`;
        })
        .join('') || '<tr><td colspan="8" class="none">None.</td></tr>';

    const s4Rows =
      model.s4
        .map(
          (x) =>
            `<tr><td><b>${esc(x.vehicle)}</b></td><td>${x.cardDT}</td><td>${x.gpsDT}</td><td class="num">${x.mins}</td><td>${x.risk}</td></tr>`
        )
        .join('') || '<tr><td colspan="5" class="none">None.</td></tr>';

    const s3Rows =
      model.s3
        .map(
          (x) =>
            `<tr><td><b>${esc(x.vehicle)}</b></td><td>${x.dt}</td><td class="num">${f2(
              x.litres
            )}</td><td>${esc(x.category)}</td></tr>`
        )
        .join('') || '<tr><td colspan="4" class="none">None.</td></tr>';

    const s5Rows =
      model.s5
        .map(
          (x) =>
            `<tr><td><b>${esc(x.vehicle)}</b></td><td>${x.dt}</td><td class="num">${
              x.variance == null ? '' : f2s(x.variance)
            }</td><td></td><td>${esc(x.issue)}${x.action ? ' — ' + esc(x.action) : ''}</td></tr>`
        )
        .join('') || '<tr><td colspan="5" class="none">None.</td></tr>';

    const CH = 14;
    let mxHtml = '';
    matrices.forEach((mx) => {
      const M = mx.mat;
      const tot = (a?: number[]) => (a || []).reduce((s, v) => s + v, 0);
      if (!M.keys.length) {
        mxHtml += `<div class="report-section"><div class="sec-title">${esc(
          mx.title
        )}</div><table class="simple"><tbody><tr><td class="none">None.</td></tr></tbody></table></div>`;
        return;
      }
      mxHtml += `<div class="report-section"><div class="sec-title">${esc(mx.title)}</div><div class="sec-sub">${esc(
        mx.sub
      )} ${M.keys.length} vehicles.</div>`;
      for (let i = 0; i < M.keys.length; i += CH) {
        const ks = M.keys.slice(i, i + CH);
        const nm = M.names.slice(i, i + CH);
        mxHtml += `<table class="simple" style="margin-bottom:8px"><thead><tr><th>Date</th>${nm
          .map((n) => `<th>${esc(n)}</th>`)
          .join('')}<th class="num">Day total</th></tr></thead><tbody>`;
        M.dates.forEach((d) => {
          const dt = M.keys.reduce((s, k) => s + tot((M.cells[d] || {})[k]), 0);
          mxHtml += `<tr><td><b>${d}</b></td>${ks
            .map((k) => {
              const arr = (M.cells[d] || {})[k];
              const txt = arr
                ? arr.map((v) => (mx.signed ? f2s(v) : mx.unit === 'min' ? String(v) : f2(v))).join(' / ')
                : '';
              return `<td class="num">${txt}</td>`;
            })
            .join('')}<td class="num"><b>${
            mx.signed ? f2s(dt) : mx.unit === 'min' ? dt : f2(dt)
          }</b></td></tr>`;
        });
        mxHtml += `<tr class="subtotal-row"><td><b>Total (${mx.unit})</b></td>${ks
          .map((k) => {
            const t = M.dates.reduce((s, d) => s + tot((M.cells[d] || {})[k]), 0);
            return `<td class="num"><b>${mx.signed ? f2s(t) : mx.unit === 'min' ? t : f2(t)}</b></td>`;
          })
          .join('')}<td></td></tr></tbody></table>`;
      }
      mxHtml += '</div>';
    });

    w.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Fuel Reconciliation ${sortedDates[0]} to ${sortedDates[sortedDates.length - 1]}</title>
    <style>
      @page{size:A4 landscape;margin:10mm 8mm}
      body{font-family:Arial,Helvetica,sans-serif;color:#1b2430;margin:0;padding:6px;font-size:9pt;background:#fff}
      .report-heading{text-align:center;margin-bottom:10px}.report-heading .l1{font-size:15pt;font-weight:700;text-transform:uppercase}.report-heading .l2{font-size:10.5pt;color:#333;margin-top:3px}
      .report-section{margin-bottom:14px}
      .sec-title{font-size:11pt;font-weight:700;margin:0 0 5px;padding:6px 10px;background:#e9e5d8;border-left:4px solid #2c5f4f}
      .sec-sub{font-size:9pt;color:#5a6470;margin:0 0 6px 2px}
      table.simple{border-collapse:collapse;width:100%;font-size:8pt}
      table.simple th,table.simple td{border:1px solid #bdb8aa;padding:4px 5px;text-align:left;vertical-align:top}
      table.simple th{background:#efece2;font-weight:700}
      table.simple td.num,table.simple th.num{text-align:right;white-space:nowrap}
      tr.subtotal-row td{background:#f2f0e6;font-weight:700}
      tr.unu td{background:#fbeee0}
      td.none{color:#94a3b8;padding:10px}
      thead{display:table-header-group} tr{page-break-inside:avoid}
    </style></head><body>
    <div class="report-heading"><div class="l1">CLEANTECH (PVT) LTD — FUEL RECONCILIATION REPORT</div>
    <div class="l2">Fuel card (Mobitel OneFuel) vs GPS fills · ${sortedDates[0]}${sortedDates.length > 1 ? ' → ' + sortedDates[sortedDates.length - 1] : ''}</div></div>
    <div class="report-section"><div class="sec-title">1. Mobitel Bill Summary</div>
      <table class="simple"><thead><tr><th>Date</th><th class="num">Diesel (L)</th><th class="num">Petrol (L)</th><th class="num">Super Diesel (L)</th><th class="num">Total Ltr</th><th class="num">Total Amount</th></tr></thead>
      <tbody>${s1Rows}<tr class="subtotal-row"><td><b>Total</b></td><td class="num"><b>${f2(sm('diesel'))}</b></td><td class="num"><b>${f2(sm('petrol'))}</b></td><td class="num"><b>${f2(sm('superD'))}</b></td><td class="num"><b>${f2(sm('total'))}</b></td><td class="num"><b>${f2(sm('amt'))}</b></td></tr></tbody></table></div>
    <div class="report-section"><div class="sec-title">2. Filling variance +1.5</div>
      <table class="simple"><thead><tr><th>Date</th><th>Vehicle</th><th>Vehicle Type</th><th class="num">Pumped (Mobitel)</th><th class="num">Pumped (GPS)</th><th class="num">Variance</th><th class="num">Vari %</th><th>Op. Comment</th></tr></thead>
      <tbody>${s2Rows}<tr class="subtotal-row"><td colspan="3"><b>Total — ${model.s2.length} entries</b></td><td class="num"><b>${f2(c2)}</b></td><td class="num"><b>${f2(g2)}</b></td><td class="num"><b>${f2s(g2 - c2)}</b></td><td class="num"><b>${c2 ? (((g2 - c2) / c2) * 100 > 0 ? '+' : '') + (((g2 - c2) / c2) * 100).toFixed(1) : '-'}%</b></td><td></td></tr></tbody></table></div>
    <div class="report-section"><div class="sec-title">4. Fuel Pump Time Variance GPS &amp; Pump Report — more than 5 min</div>
      <table class="simple"><thead><tr><th>Vehicle</th><th>Card Time</th><th>GPS Time</th><th class="num">Variance (min)</th><th>Risk</th></tr></thead><tbody>${s4Rows}</tbody></table></div>
    <div class="report-section"><div class="sec-title">3. GPS Unavailable &amp; Offline Vehicle List</div>
      <table class="simple"><thead><tr><th>Vehicle</th><th>Date &amp; Time</th><th class="num">Pumped (Mobitel)</th><th>Category</th></tr></thead>
      <tbody>${s3Rows}<tr class="subtotal-row"><td colspan="2"><b>Total — ${model.s3.length} entries</b></td><td class="num"><b>${f2(l3)}</b></td><td></td></tr></tbody></table></div>
    <div class="report-section"><div class="sec-title">5. Abnormal Highlights &amp; Recommendations</div>
      <table class="simple"><thead><tr><th>Vehicle</th><th>Date &amp; Time</th><th class="num">Variance</th><th>Op. Comment</th><th>Observed issue / Action</th></tr></thead><tbody>${s5Rows}</tbody></table></div>
    ${mxHtml}
    <script>window.onload=function(){setTimeout(function(){window.focus();window.print();},300);};<\/script>
    </body></html>`);
    w.document.close();
  };

  const sm = (k: 'diesel' | 'petrol' | 'superD' | 'total' | 'amt') =>
    model.s1.reduce((s, x) => s + x[k], 0);
  const c2 = model.s2.reduce((s, x) => s + x.cardL, 0);
  const g2 = model.s2.reduce((s, x) => s + x.gpsL, 0);
  const l3 = model.s3.reduce((s, x) => s + x.litres, 0);

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-end justify-between gap-4 border-b border-slate-200 pb-5">
        <div>
          <h1 className="text-xl font-bold text-slate-900">Multi-Day Consolidation (Excel Format)</h1>
          <p className="text-sm text-slate-600 mt-1">
            Combine daily Fuel Card vs GPS Reconciliation PDFs or live reconciliation data into the multi-sheet GPS.xlsx format with Vehicle × Day variance matrices.
          </p>
        </div>
        {sortedDates.length > 0 && (
          <div className="flex flex-wrap items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={downloadXlsx}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-[#1F5143] hover:bg-[#173D32] text-white text-xs font-medium rounded-md transition-colors whitespace-nowrap"
            >
              <Download className="w-3.5 h-3.5" /> Download Excel (.xlsx)
            </button>
            {onPublishCompletedFile && (
              <button
                type="button"
                disabled={publishing}
                onClick={() => void publishXlsx()}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-slate-900 hover:bg-slate-800 text-white text-xs font-medium rounded-md transition-colors whitespace-nowrap disabled:opacity-60"
              >
                <Send className="w-3.5 h-3.5" />{' '}
                {publishing ? 'Publishing...' : 'Publish for Fuel Officer'}
              </button>
            )}
            <button
              type="button"
              onClick={downloadPdf}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-white hover:bg-slate-50 text-slate-800 border border-slate-300 text-xs font-medium rounded-md transition-colors whitespace-nowrap"
            >
              <Printer className="w-3.5 h-3.5" /> Print / Save PDF
            </button>
            <button
              type="button"
              onClick={clearAll}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-white hover:bg-red-50 text-red-700 border border-red-200 text-xs font-medium rounded-md transition-colors whitespace-nowrap"
            >
              <Trash2 className="w-3.5 h-3.5" /> Clear All
            </button>
          </div>
        )}
      </div>

      {managementComments.length > 0 && (
        <div className="bg-white border border-slate-200 rounded-md p-4 space-y-2">
          <div className="text-xs font-semibold text-slate-800">
            Recent Management Comments ({managementComments.length})
          </div>
          <div className="space-y-1.5">
            {managementComments.slice(0, 3).map((c) => (
              <div key={c.id} className="text-xs text-slate-600 flex flex-wrap gap-2">
                <span className="font-semibold text-slate-900">{c.author}</span>
                <span>· {c.status}</span>
                {c.reportDate && <span className="font-data">· {c.reportDate}</span>}
                <span>— {c.comment}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Upload & Direct Sync Bar */}
      <div className="bg-white border border-slate-200 rounded-md p-5 space-y-4">
        <div className="flex flex-wrap items-center gap-3">
          <label className="inline-flex items-center gap-2 px-4 py-2 bg-[#1F5143] hover:bg-[#173D32] text-white text-xs font-medium rounded-md cursor-pointer transition-colors whitespace-nowrap">
            <Upload className="w-4 h-4" /> Choose Daily Report PDFs (multi-select)
            <input type="file" accept="application/pdf,.pdf" multiple className="hidden" onChange={handleUpload} />
          </label>

          {canSyncFromGps && (
            <button
              type="button"
              onClick={onSyncFromGps}
              className="inline-flex items-center gap-1.5 px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-medium rounded-md transition-colors whitespace-nowrap"
            >
              <RefreshCw className="w-3.5 h-3.5" /> Import Active GPS Reconciliation
            </button>
          )}

          <span className="text-xs text-slate-500">{status}</span>
        </div>

        {/* Loaded Days Verification Table */}
        {sortedDates.length > 0 && (
          <div className="pt-3 border-t border-slate-100 space-y-3">
            <div className="overflow-x-auto">
              <table className="w-full text-xs text-left border-collapse">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200 text-slate-600">
                    <th className="px-3 py-2 font-semibold">Date</th>
                    <th className="px-3 py-2 font-semibold">Source</th>
                    <th className="px-3 py-2 font-semibold text-right">Unusual</th>
                    <th className="px-3 py-2 font-semibold text-right">&gt;±1.5 L</th>
                    <th className="px-3 py-2 font-semibold text-right">GPS Unavail.</th>
                    <th className="px-3 py-2 font-semibold text-right">Time &gt;5m</th>
                    <th className="px-3 py-2 font-semibold text-right">Variance OK</th>
                    <th className="px-3 py-2 font-semibold text-right">Abnormal</th>
                    <th className="px-3 py-2 font-semibold text-right">Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {sortedDates.map((d) => {
                    const day = days[d];
                    const renderCheck = (label: string) => {
                      const c = day.checks.find((x) => x.label === label);
                      if (!c) return <td className="px-3 py-2 text-right text-slate-400">–</td>;
                      return (
                        <td
                          className={`px-3 py-2 text-right font-data tabular-nums font-medium ${
                            c.ok ? 'text-emerald-700' : 'text-red-700'
                          }`}
                        >
                          {c.got}
                          {c.ok ? ' ✓' : ` / ${c.want} ⚠`}
                        </td>
                      );
                    };
                    return (
                      <tr key={d} className="hover:bg-slate-50">
                        <td className="px-3 py-2 font-data font-semibold text-slate-900">{d}</td>
                        <td className="px-3 py-2 text-slate-600 max-w-xs truncate">{day.fileName}</td>
                        {renderCheck('Unusual filling variance')}
                        {renderCheck('More than ±1.5 L variance')}
                        {renderCheck('GPS unavailable & offline')}
                        {renderCheck('Time variance > 5 min')}
                        {renderCheck('Variance OK')}
                        {day.abnormalMissing ? (
                          <td className="px-3 py-2 text-right text-amber-700 font-medium">missing ⚠</td>
                        ) : (
                          renderCheck('Abnormal highlights')
                        )}
                        <td className="px-3 py-2 text-right">
                          <button
                            type="button"
                            onClick={() => removeDay(d)}
                            className="text-red-700 hover:underline font-medium"
                          >
                            Remove
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div className="flex flex-wrap items-center justify-between gap-3 text-xs text-slate-500">
              <span>✓ confirms parsed row counts match the section heading totals.</span>
              {sortedDates.some((d) => days[d].abnormalMissing) && (
                <label className="inline-flex items-center gap-2 text-slate-700 font-medium cursor-pointer">
                  <input
                    type="checkbox"
                    checked={rebuildMultiples}
                    onChange={(e) => setRebuildMultiples(e.target.checked)}
                    className="rounded border-slate-300 text-[#1F5143] focus:ring-[#1F5143]"
                  />
                  Rebuild “multiple same-day fillings” when Section 7 is missing
                </label>
              )}
            </div>
          </div>
        )}
      </div>

      {sortedDates.length === 0 ? (
        <div className="bg-white border border-slate-200 rounded-md p-12 text-center space-y-3">
          <div className="text-sm font-medium text-slate-800">No consolidated daily reports loaded</div>
          <p className="text-xs text-slate-500 max-w-md mx-auto">
            Upload daily Fuel Card vs GPS Reconciliation PDFs above, or complete a reconciliation in the GPS Reconciliation module and import it directly.
          </p>
        </div>
      ) : (
        <div className="space-y-8">
          {/* 1. Mobitel Bill Summary */}
          <section className="bg-white border border-slate-200 rounded-md overflow-hidden">
            <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
              <h2 className="text-sm font-semibold text-slate-900">1. Mobitel Bill Summary</h2>
              <span className="text-xs text-slate-500 font-data">
                {sortedDates[0]} → {sortedDates[sortedDates.length - 1]} ({sortedDates.length} day{sortedDates.length > 1 ? 's' : ''})
              </span>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-xs text-left border-collapse">
                <thead className="bg-slate-50/60 border-b border-slate-200 text-slate-600">
                  <tr>
                    <th className="px-4 py-2.5 font-semibold">Date</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Diesel (L)</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Petrol (L)</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Super Diesel (L)</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Total Ltr</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Total Amount</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {model.s1.map((x) => (
                    <tr key={x.date} className="hover:bg-slate-50">
                      <td className="px-4 py-2 font-data font-medium text-slate-900">{x.date}</td>
                      <td className="px-4 py-2 font-data text-right tabular-nums">{f2(x.diesel)}</td>
                      <td className="px-4 py-2 font-data text-right tabular-nums">{f2(x.petrol)}</td>
                      <td className="px-4 py-2 font-data text-right tabular-nums">{f2(x.superD)}</td>
                      <td className="px-4 py-2 font-data text-right tabular-nums font-medium">{f2(x.total)}</td>
                      <td className="px-4 py-2 font-data text-right tabular-nums">{f2(x.amt)}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot className="bg-slate-100 border-t border-slate-300 font-semibold text-slate-900">
                  <tr>
                    <td className="px-4 py-2.5">Total</td>
                    <td className="px-4 py-2.5 font-data text-right tabular-nums">{f2(sm('diesel'))}</td>
                    <td className="px-4 py-2.5 font-data text-right tabular-nums">{f2(sm('petrol'))}</td>
                    <td className="px-4 py-2.5 font-data text-right tabular-nums">{f2(sm('superD'))}</td>
                    <td className="px-4 py-2.5 font-data text-right tabular-nums">{f2(sm('total'))}</td>
                    <td className="px-4 py-2.5 font-data text-right tabular-nums">{f2(sm('amt'))}</td>
                  </tr>
                </tfoot>
              </table>
            </div>
          </section>

          {/* 2. Filling variance +1.5 */}
          <section className="bg-white border border-slate-200 rounded-md overflow-hidden">
            <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
              <h2 className="text-sm font-semibold text-slate-900">2. Filling Variance &gt; ±1.5 L</h2>
              <span className="text-xs text-slate-500">Highlighted rows indicate UNUSUAL fillings</span>
            </div>
            <div className="overflow-x-auto max-h-96">
              <table className="w-full text-xs text-left border-collapse">
                <thead className="bg-slate-50/60 border-b border-slate-200 text-slate-600 sticky top-0">
                  <tr>
                    <th className="px-4 py-2.5 font-semibold">Date &amp; Time</th>
                    <th className="px-4 py-2.5 font-semibold">Vehicle</th>
                    <th className="px-4 py-2.5 font-semibold">Vehicle Type</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Pumped (Mobitel)</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Pumped (GPS)</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Variance</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Vari %</th>
                    <th className="px-4 py-2.5 font-semibold">Flag</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {model.s2.length === 0 ? (
                    <tr>
                      <td colSpan={8} className="px-4 py-8 text-center text-slate-400">
                        None.
                      </td>
                    </tr>
                  ) : (
                    model.s2.map((x, idx) => {
                      const v = x.gpsL - x.cardL;
                      const pct = x.cardL ? (v / x.cardL) * 100 : 0;
                      return (
                        <tr key={idx} className={x.unusual ? 'bg-amber-50/70' : 'hover:bg-slate-50'}>
                          <td className="px-4 py-2 font-data whitespace-nowrap">
                            {x.date} {x.time}
                          </td>
                          <td className="px-4 py-2 font-data font-semibold text-slate-900">{x.vehicle}</td>
                          <td className="px-4 py-2 text-slate-600">{x.type}</td>
                          <td className="px-4 py-2 font-data text-right tabular-nums">{f2(x.cardL)}</td>
                          <td className="px-4 py-2 font-data text-right tabular-nums">{f2(x.gpsL)}</td>
                          <td className="px-4 py-2 font-data text-right tabular-nums font-medium text-red-700">
                            {f2s(v)}
                          </td>
                          <td className="px-4 py-2 font-data text-right tabular-nums">
                            {x.cardL ? `${pct > 0 ? '+' : ''}${pct.toFixed(1)}%` : '-'}
                          </td>
                          <td className="px-4 py-2">
                            {x.unusual ? (
                              <span className="inline-flex items-center gap-1 text-amber-800 font-medium">
                                <AlertTriangle className="w-3.5 h-3.5" /> Unusual
                              </span>
                            ) : (
                              <span className="text-slate-500">Mismatch</span>
                            )}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
                {model.s2.length > 0 && (
                  <tfoot className="bg-slate-100 border-t border-slate-300 font-semibold text-slate-900">
                    <tr>
                      <td colSpan={3} className="px-4 py-2.5">
                        Total — {model.s2.length} entries
                      </td>
                      <td className="px-4 py-2.5 font-data text-right tabular-nums">{f2(c2)}</td>
                      <td className="px-4 py-2.5 font-data text-right tabular-nums">{f2(g2)}</td>
                      <td className="px-4 py-2.5 font-data text-right tabular-nums">{f2s(g2 - c2)}</td>
                      <td className="px-4 py-2.5 font-data text-right tabular-nums">
                        {c2 ? `${((g2 - c2) / c2) * 100 > 0 ? '+' : ''}${(((g2 - c2) / c2) * 100).toFixed(1)}%` : '-'}
                      </td>
                      <td />
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>
          </section>

          {/* 4. Time Variance & 3. Offline List */}
          <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
            <section className="bg-white border border-slate-200 rounded-md overflow-hidden">
              <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                <h2 className="text-sm font-semibold text-slate-900">4. Fuel Pump Time Variance (&gt; 5 min)</h2>
                <span className="text-xs text-slate-500">{model.s4.length} entries</span>
              </div>
              <div className="overflow-x-auto max-h-80">
                <table className="w-full text-xs text-left border-collapse">
                  <thead className="bg-slate-50/60 border-b border-slate-200 text-slate-600 sticky top-0">
                    <tr>
                      <th className="px-3 py-2 font-semibold">Vehicle</th>
                      <th className="px-3 py-2 font-semibold">Card Time</th>
                      <th className="px-3 py-2 font-semibold">GPS Time</th>
                      <th className="px-3 py-2 font-semibold text-right">Mins</th>
                      <th className="px-3 py-2 font-semibold">Risk</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {model.s4.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="px-4 py-8 text-center text-slate-400">
                          None.
                        </td>
                      </tr>
                    ) : (
                      model.s4.map((x, idx) => (
                        <tr key={idx} className="hover:bg-slate-50">
                          <td className="px-3 py-2 font-data font-semibold text-slate-900">{x.vehicle}</td>
                          <td className="px-3 py-2 font-data whitespace-nowrap">{x.cardDT}</td>
                          <td className="px-3 py-2 font-data whitespace-nowrap">{x.gpsDT}</td>
                          <td className="px-3 py-2 font-data text-right tabular-nums">{x.mins}</td>
                          <td
                            className={`px-3 py-2 font-medium ${
                              x.risk === 'Critical'
                                ? 'text-red-700'
                                : x.risk === 'High'
                                  ? 'text-amber-700'
                                  : 'text-emerald-700'
                            }`}
                          >
                            {x.risk}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </section>

            <section className="bg-white border border-slate-200 rounded-md overflow-hidden">
              <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                <h2 className="text-sm font-semibold text-slate-900">3. GPS Unavailable &amp; Offline List</h2>
                <span className="text-xs text-slate-500 font-data">
                  {model.s3.length} entries · {f2(l3)} L
                </span>
              </div>
              <div className="overflow-x-auto max-h-80">
                <table className="w-full text-xs text-left border-collapse">
                  <thead className="bg-slate-50/60 border-b border-slate-200 text-slate-600 sticky top-0">
                    <tr>
                      <th className="px-3 py-2 font-semibold">Vehicle</th>
                      <th className="px-3 py-2 font-semibold">Date &amp; Time</th>
                      <th className="px-3 py-2 font-semibold text-right">Pumped (Mobitel)</th>
                      <th className="px-3 py-2 font-semibold">Category</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {model.s3.length === 0 ? (
                      <tr>
                        <td colSpan={4} className="px-4 py-8 text-center text-slate-400">
                          None.
                        </td>
                      </tr>
                    ) : (
                      model.s3.map((x, idx) => (
                        <tr key={idx} className="hover:bg-slate-50">
                          <td className="px-3 py-2 font-data font-semibold text-slate-900">{x.vehicle}</td>
                          <td className="px-3 py-2 font-data whitespace-nowrap">{x.dt}</td>
                          <td className="px-3 py-2 font-data text-right tabular-nums">{f2(x.litres)}</td>
                          <td className="px-3 py-2 text-slate-600">{x.category}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </section>
          </div>

          {/* 5. Abnormal Highlights */}
          <section className="bg-white border border-slate-200 rounded-md overflow-hidden">
            <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
              <h2 className="text-sm font-semibold text-slate-900">5. Abnormal Highlights &amp; Recommendations</h2>
              <span className="text-xs text-slate-500">{model.s5.length} items</span>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-xs text-left border-collapse">
                <thead className="bg-slate-50/60 border-b border-slate-200 text-slate-600">
                  <tr>
                    <th className="px-4 py-2.5 font-semibold">Vehicle</th>
                    <th className="px-4 py-2.5 font-semibold">Date &amp; Time</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Variance</th>
                    <th className="px-4 py-2.5 font-semibold">Observed Issue / Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {model.s5.length === 0 ? (
                    <tr>
                      <td colSpan={4} className="px-4 py-8 text-center text-slate-400">
                        None.
                      </td>
                    </tr>
                  ) : (
                    model.s5.map((x, idx) => (
                      <tr key={idx} className="hover:bg-slate-50">
                        <td className="px-4 py-2 font-data font-semibold text-slate-900">{x.vehicle}</td>
                        <td className="px-4 py-2 font-data whitespace-nowrap">{x.dt}</td>
                        <td className="px-4 py-2 font-data text-right tabular-nums">
                          {x.variance == null ? '—' : f2s(x.variance)}
                        </td>
                        <td className="px-4 py-2 text-slate-700">
                          {x.issue}
                          {x.action ? ` — ${x.action}` : ''}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </section>

          {/* 6, 7, 8 Vehicle x Day Matrices */}
          {matrices.map((mx) => {
            const M = mx.mat;
            const tot = (a?: number[]) => (a || []).reduce((s, v) => s + v, 0);
            return (
              <section key={mx.title} className="bg-white border border-slate-200 rounded-md overflow-hidden">
                <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200">
                  <h2 className="text-sm font-semibold text-slate-900">{mx.title}</h2>
                  <p className="text-xs text-slate-500 mt-0.5">
                    {mx.sub} ({M.keys.length} vehicles)
                  </p>
                </div>
                {M.keys.length === 0 ? (
                  <div className="px-5 py-8 text-center text-xs text-slate-400">None.</div>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs text-left border-collapse">
                      <thead className="bg-slate-50/60 border-b border-slate-200 text-slate-600">
                        <tr>
                          <th className="px-3 py-2 font-semibold sticky left-0 bg-slate-50 z-10 border-r border-slate-200">
                            Date
                          </th>
                          {M.names.map((n) => (
                            <th key={n} className="px-3 py-2 font-semibold font-data whitespace-nowrap text-right">
                              {n}
                            </th>
                          ))}
                          <th className="px-3 py-2 font-semibold text-right bg-slate-100 border-l border-slate-200">
                            Day Total
                          </th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {M.dates.map((d) => {
                          const dt = M.keys.reduce((s, k) => s + tot((M.cells[d] || {})[k]), 0);
                          return (
                            <tr key={d} className="hover:bg-slate-50">
                              <td className="px-3 py-2 font-data font-medium text-slate-900 sticky left-0 bg-white border-r border-slate-200 whitespace-nowrap">
                                {d}
                              </td>
                              {M.keys.map((k) => {
                                const arr = (M.cells[d] || {})[k];
                                const txt = arr
                                  ? arr
                                      .map((v) => (mx.signed ? f2s(v) : mx.unit === 'min' ? String(v) : f2(v)))
                                      .join(' / ')
                                  : '';
                                return (
                                  <td key={k} className="px-3 py-2 font-data text-right tabular-nums whitespace-nowrap">
                                    {txt}
                                  </td>
                                );
                              })}
                              <td className="px-3 py-2 font-data text-right tabular-nums font-semibold bg-slate-50 border-l border-slate-200">
                                {mx.signed ? f2s(dt) : mx.unit === 'min' ? dt : f2(dt)}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                      <tfoot className="bg-slate-100 border-t border-slate-300 font-semibold text-slate-900">
                        <tr>
                          <td className="px-3 py-2 sticky left-0 bg-slate-100 border-r border-slate-200">
                            Total ({mx.unit})
                          </td>
                          {M.keys.map((k) => {
                            const t = M.dates.reduce((s, d) => s + tot((M.cells[d] || {})[k]), 0);
                            return (
                              <td key={k} className="px-3 py-2 font-data text-right tabular-nums">
                                {mx.signed ? f2s(t) : mx.unit === 'min' ? t : f2(t)}
                              </td>
                            );
                          })}
                          <td className="px-3 py-2 bg-slate-100 border-l border-slate-200" />
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                )}
              </section>
            );
          })}
        </div>
      )}
    </div>
  );
};
