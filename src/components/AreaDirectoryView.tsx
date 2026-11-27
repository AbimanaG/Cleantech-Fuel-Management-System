import React, { useState } from 'react';
import ExcelJS from 'exceljs';
import { Download, Upload, Plus, RotateCcw, Trash2, Search } from 'lucide-react';
import { parseCsvText, saveBlob, todayIso, xlCellText } from '../utils/voucherParser';

interface AreaDirectoryViewProps {
  vehicleAreas: Record<string, string[]>;
  onUpdateVehicleAreas: (next: Record<string, string[]>) => void;
  onResetVehicleAreas: () => void;
}

export const AreaDirectoryView: React.FC<AreaDirectoryViewProps> = ({
  vehicleAreas,
  onUpdateVehicleAreas,
  onResetVehicleAreas,
}) => {
  const [search, setSearch] = useState('');
  const [newVeh, setNewVeh] = useState('');
  const [newArea, setNewArea] = useState('');
  const [uploadMode, setUploadMode] = useState<'merge' | 'replace'>('merge');
  const [status, setStatus] = useState<{ msg: string; error?: boolean } | null>(null);

  const splitAreas = (s: string) =>
    String(s || '')
      .split(';')
      .map((x) => x.trim())
      .filter(Boolean);

  const sortedVehicles = Object.keys(vehicleAreas)
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }))
    .filter((v) => {
      if (!search.trim()) return true;
      const q = search.trim().toLowerCase();
      return (
        v.toLowerCase().includes(q) ||
        (vehicleAreas[v] || []).some((a) => a.toLowerCase().includes(q))
      );
    });

  const handleEdit = (veh: string, value: string) => {
    const arr = splitAreas(value);
    const next = { ...vehicleAreas };
    if (arr.length) next[veh] = arr;
    else delete next[veh];
    onUpdateVehicleAreas(next);
    setStatus({ msg: `Saved — ${veh} updated.` });
  };

  const handleDelete = (veh: string) => {
    const next = { ...vehicleAreas };
    delete next[veh];
    onUpdateVehicleAreas(next);
    setStatus({ msg: `Removed ${veh}.` });
  };

  const handleAdd = () => {
    const v = newVeh.trim();
    const a = splitAreas(newArea);
    if (!v || !a.length) {
      setStatus({ msg: 'Enter both a Vehicle Number and an Area.', error: true });
      return;
    }
    const next = { ...vehicleAreas, [v]: a };
    onUpdateVehicleAreas(next);
    setNewVeh('');
    setNewArea('');
    setStatus({ msg: `Saved — ${v} added.` });
  };

  const handleDownload = async (fmt: 'xlsx' | 'csv') => {
    const pairs: [string, string][] = [];
    Object.keys(vehicleAreas)
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }))
      .forEach((v) => (vehicleAreas[v] || []).forEach((a) => pairs.push([v, a])));

    if (!pairs.length) {
      setStatus({ msg: 'The list is empty.', error: true });
      return;
    }
    const name = 'Vehicle_Area_List_' + todayIso();
    if (fmt === 'csv') {
      const q = (v: string) => (/[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v);
      saveBlob(
        new Blob(['\ufeffVehicle Number,Area\n' + pairs.map((p) => q(p[0]) + ',' + q(p[1])).join('\n')], {
          type: 'text/csv;charset=utf-8;',
        }),
        name + '.csv'
      );
      return;
    }
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Vehicle Areas', { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = [
      { header: 'Vehicle Number', key: 'v', width: 24 },
      { header: 'Area', key: 'a', width: 42 },
    ];
    const g = { style: 'thin' as const, color: { argb: 'FF999999' } };
    const border = { top: g, left: g, bottom: g, right: g };
    ws.getRow(1).eachCell((c) => {
      c.font = { bold: true };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEEEEE' } };
      c.border = border;
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    pairs.forEach((p) => {
      const r = ws.addRow({ v: p[0], a: p[1] });
      r.eachCell((c) => {
        c.border = border;
      });
    });
    ws.autoFilter = { from: 'A1', to: 'B1' };
    saveBlob(
      new Blob([await wb.xlsx.writeBuffer()], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }),
      name + '.xlsx'
    );
  };

  const handleUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      let rows: string[][] = [];
      if (/\.csv$/i.test(file.name) || file.type === 'text/csv') {
        rows = parseCsvText(new TextDecoder('utf-8').decode(await file.arrayBuffer()));
      } else {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await file.arrayBuffer());
        const ws = wb.worksheets[0];
        if (!ws) throw new Error('The workbook has no sheet.');
        ws.eachRow({ includeEmpty: false }, (row) => {
          const vals: string[] = [];
          row.eachCell({ includeEmpty: true }, (c, i) => {
            vals[i - 1] = xlCellText(c.value);
          });
          rows.push(vals);
        });
      }
      rows = rows.map((r) => (r || []).map((x) => String(x ?? '').trim())).filter((r) => r.some(Boolean));
      if (!rows.length) throw new Error('The file is empty.');

      let vi = 0;
      let ai = 1;
      let start = 0;
      const head = rows[0].map((x) => x.toLowerCase());
      const hv = head.findIndex((x) => /vehicle/.test(x));
      const ha = head.findIndex((x) => /area/.test(x));
      if (hv >= 0 && ha >= 0) {
        vi = hv;
        ai = ha;
        start = 1;
      }

      const incoming: Record<string, string[]> = {};
      rows.slice(start).forEach((r) => {
        const v = (r[vi] || '').trim();
        const a = (r[ai] || '').trim();
        if (!v || !a) return;
        splitAreas(a).forEach((x) => {
          incoming[v] = incoming[v] || [];
          if (!incoming[v].includes(x)) incoming[v].push(x);
        });
      });

      const vehCount = Object.keys(incoming).length;
      if (!vehCount) throw new Error('No "Vehicle Number" + "Area" rows found.');

      let added = 0;
      let updated = 0;
      let next: Record<string, string[]>;
      if (uploadMode === 'replace') {
        next = incoming;
        added = vehCount;
      } else {
        next = { ...vehicleAreas };
        Object.keys(incoming).forEach((v) => {
          if (next[v]) {
            if (next[v].join('|') !== incoming[v].join('|')) updated++;
          } else {
            added++;
          }
          next[v] = incoming[v];
        });
      }
      onUpdateVehicleAreas(next);
      setStatus({
        msg: `${vehCount} vehicle${vehCount === 1 ? '' : 's'} read from "${file.name}" — ${
          uploadMode === 'replace' ? 'list replaced' : `${added} added, ${updated} updated`
        }. Saved in browser.`,
      });
    } catch (err: any) {
      setStatus({ msg: `Could not read file: ${err.message}`, error: true });
    }
    e.target.value = '';
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-end justify-between gap-4 border-b border-slate-200 pb-5">
        <div>
          <h1 className="text-xl font-bold text-slate-900">Vehicle → Area Directory</h1>
          <p className="text-sm text-slate-600 mt-1">
            Default area suggestions for {Object.keys(vehicleAreas).length} fleet vehicles, saved automatically in this browser. Separate multiple areas with semicolons (;).
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2 shrink-0">
          <button
            type="button"
            onClick={() => handleDownload('xlsx')}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-[#1F5143] hover:bg-[#173D32] text-white text-xs font-medium rounded-md transition-colors whitespace-nowrap"
          >
            <Download className="w-3.5 h-3.5" /> Download Excel
          </button>
          <button
            type="button"
            onClick={() => handleDownload('csv')}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 text-xs font-medium rounded-md transition-colors whitespace-nowrap"
          >
            <Download className="w-3.5 h-3.5" /> Download CSV
          </button>
          <button
            type="button"
            onClick={() => {
              onResetVehicleAreas();
              setStatus({ msg: 'Built-in Cleantech vehicle directory restored.' });
            }}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-white hover:bg-red-50 text-red-700 border border-red-200 text-xs font-medium rounded-md transition-colors whitespace-nowrap"
          >
            <RotateCcw className="w-3.5 h-3.5" /> Reset Built-In
          </button>
        </div>
      </div>

      {/* Upload / Add Controls */}
      <div className="bg-white border border-slate-200 rounded-md p-5 space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div className="flex flex-wrap items-center gap-3">
            <label className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-medium rounded-md cursor-pointer transition-colors whitespace-nowrap">
              <Upload className="w-3.5 h-3.5" /> Upload Area List (.xlsx / .csv)
              <input type="file" accept=".xlsx,.csv,text/csv" className="hidden" onChange={handleUpload} />
            </label>
            <select
              value={uploadMode}
              onChange={(e) => setUploadMode(e.target.value as 'merge' | 'replace')}
              className="px-3 py-2 bg-white border border-slate-300 rounded-md text-xs text-slate-700 focus:outline-none focus:border-[#1F5143]"
            >
              <option value="merge">Merge — update &amp; add vehicles</option>
              <option value="replace">Replace the whole list</option>
            </select>
          </div>

          {status && (
            <div className={`text-xs font-medium ${status.error ? 'text-red-700' : 'text-emerald-700'}`}>
              {status.msg}
            </div>
          )}
        </div>

        <div className="grid grid-cols-1 md:grid-cols-12 gap-3 pt-4 border-t border-slate-100 items-end">
          <div className="md:col-span-5">
            <label className="block text-xs font-medium text-slate-500 mb-1">Find Vehicle or Area</label>
            <div className="relative">
              <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="e.g. LM-1328 or D5"
                className="w-full pl-9 pr-3 py-2 border border-slate-300 rounded-md text-sm focus:outline-none focus:border-[#1F5143]"
              />
            </div>
          </div>
          <div className="md:col-span-3">
            <label className="block text-xs font-medium text-slate-500 mb-1">New Vehicle Number</label>
            <input
              type="text"
              value={newVeh}
              onChange={(e) => setNewVeh(e.target.value)}
              placeholder="e.g. LP-8800"
              className="w-full px-3 py-2 border border-slate-300 rounded-md text-sm focus:outline-none focus:border-[#1F5143]"
            />
          </div>
          <div className="md:col-span-3">
            <label className="block text-xs font-medium text-slate-500 mb-1">Assigned Area</label>
            <input
              type="text"
              value={newArea}
              onChange={(e) => setNewArea(e.target.value)}
              placeholder="e.g. D5"
              onKeyDown={(e) => e.key === 'Enter' && handleAdd()}
              className="w-full px-3 py-2 border border-slate-300 rounded-md text-sm focus:outline-none focus:border-[#1F5143]"
            />
          </div>
          <div className="md:col-span-1">
            <button
              type="button"
              onClick={handleAdd}
              className="w-full inline-flex items-center justify-center gap-1 px-3 py-2 bg-[#1F5143] hover:bg-[#173D32] text-white text-sm font-medium rounded-md transition-colors whitespace-nowrap"
            >
              <Plus className="w-4 h-4" /> Add
            </button>
          </div>
        </div>
      </div>

      {/* Directory Table */}
      <div className="bg-white border border-slate-200 rounded-md overflow-hidden">
        <div className="overflow-x-auto max-h-[60vh]">
          <table className="w-full text-sm text-left border-collapse">
            <thead className="bg-slate-50 border-b border-slate-200 sticky top-0 z-10">
              <tr>
                <th className="px-4 py-2.5 text-xs font-semibold text-slate-600 w-16">#</th>
                <th className="px-4 py-2.5 text-xs font-semibold text-slate-600 w-56">Vehicle Number</th>
                <th className="px-4 py-2.5 text-xs font-semibold text-slate-600">
                  Area (first = suggested; separate multiple with ;)
                </th>
                <th className="px-4 py-2.5 text-xs font-semibold text-slate-600 text-right w-28">Action</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {sortedVehicles.length === 0 ? (
                <tr>
                  <td colSpan={4} className="px-4 py-10 text-center text-slate-400">
                    No vehicles match your search.
                  </td>
                </tr>
              ) : (
                sortedVehicles.map((v, idx) => (
                  <tr key={v} className="hover:bg-slate-50/80">
                    <td className="px-4 py-2 font-data text-xs text-slate-400 tabular-nums">{idx + 1}</td>
                    <td className="px-4 py-2 font-data font-medium text-slate-900">{v}</td>
                    <td className="px-4 py-1.5">
                      <input
                        type="text"
                        defaultValue={(vehicleAreas[v] || []).join('; ')}
                        key={v + ':' + (vehicleAreas[v] || []).join(';')}
                        onBlur={(e) => {
                          const nextVal = e.target.value;
                          if (nextVal !== (vehicleAreas[v] || []).join('; ')) {
                            handleEdit(v, nextVal);
                          }
                        }}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') {
                            e.currentTarget.blur();
                          }
                        }}
                        className="w-full px-2.5 py-1.5 border border-slate-200 rounded text-sm bg-white focus:outline-none focus:border-[#1F5143]"
                      />
                    </td>
                    <td className="px-4 py-2 text-right">
                      <button
                        type="button"
                        onClick={() => handleDelete(v)}
                        className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-red-700 hover:bg-red-50 rounded transition-colors"
                      >
                        <Trash2 className="w-3.5 h-3.5" /> Remove
                      </button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
