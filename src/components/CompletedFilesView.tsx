import React, { useMemo, useState } from 'react';
import { Download, RefreshCw, Search, Trash2, FileText } from 'lucide-react';
import type { CompletedFileModule, CompletedFileRecord, UserRole } from '../types/fuel';
import { completedFilesApi } from '../utils/apiClient';

interface CompletedFilesViewProps {
  files: CompletedFileRecord[];
  userRole: UserRole;
  onRefresh: () => Promise<void>;
  onDeleteFile?: (id: string) => Promise<void>;
}

const MODULE_LABELS: Record<CompletedFileModule, string> = {
  gps: 'Fuel Card vs GPS Reconciliation',
  consolidation: 'Multi-Day Consolidation (Excel Format)',
  voucher: 'Fuel Voucher Analysis',
};

function formatBytes(bytes: number): string {
  if (!bytes || bytes < 1024) return `${bytes || 0} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

export const CompletedFilesView: React.FC<CompletedFilesViewProps> = ({
  files,
  userRole,
  onRefresh,
  onDeleteFile,
}) => {
  const [moduleFilter, setModuleFilter] = useState<'all' | CompletedFileModule>('all');
  const [search, setSearch] = useState('');
  const [downloadingId, setDownloadingId] = useState<string | null>(null);
  const [refreshing, setRefreshing] = useState(false);
  const [notice, setNotice] = useState<{ msg: string; error?: boolean } | null>(null);

  const visibleFiles = useMemo(() => {
    const q = search.trim().toLowerCase();
    return files.filter((f) => {
      if (moduleFilter !== 'all' && f.module !== moduleFilter) return false;
      if (q) {
        const hay = `${f.title} ${f.fileName} ${f.summary} ${f.uploadedBy} ${f.reportDate}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }, [files, moduleFilter, search]);

  const handleDownload = async (file: CompletedFileRecord) => {
    setDownloadingId(file.id);
    setNotice(null);
    try {
      await completedFilesApi.download(file);
    } catch (err: any) {
      setNotice({ msg: err?.message || 'Could not download file.', error: true });
    } finally {
      setDownloadingId(null);
    }
  };

  const handleRefresh = async () => {
    setRefreshing(true);
    try {
      await onRefresh();
    } finally {
      setRefreshing(false);
    }
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-end justify-between gap-4 border-b border-slate-200 pb-5">
        <div>
          <h1 className="text-xl font-bold text-slate-900">Completed Fuel Files</h1>
          <p className="text-sm text-slate-600 mt-1">
            Download finalized Excel workbooks, PDF reconciliation reports, and voucher exports published by Controllers.
          </p>
        </div>
        <button
          type="button"
          onClick={() => void handleRefresh()}
          disabled={refreshing}
          className="inline-flex items-center gap-2 px-3.5 py-2 bg-white hover:bg-slate-50 text-slate-800 border border-slate-300 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0"
        >
          <RefreshCw className={`w-3.5 h-3.5 ${refreshing ? 'animate-spin' : ''}`} />
          <span>{refreshing ? 'Refreshing...' : 'Refresh List'}</span>
        </button>
      </div>

      {/* Filters */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
        <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-md w-fit">
          {(
            [
              { id: 'all', label: 'All Completed' },
              { id: 'gps', label: 'GPS Reconciliation' },
              { id: 'consolidation', label: 'Multi-Day Excel' },
              { id: 'voucher', label: 'Voucher Analysis' },
            ] as const
          ).map((tab) => (
            <button
              key={tab.id}
              type="button"
              onClick={() => setModuleFilter(tab.id)}
              className={`px-3 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap ${
                moduleFilter === tab.id
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              {tab.label}
            </button>
          ))}
        </div>

        <div className="relative w-full md:w-72">
          <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search file name, date, or controller..."
            className="w-full pl-9 pr-3 py-1.5 text-sm border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-[#1F5143]/25"
          />
        </div>
      </div>

      {notice && (
        <div
          className={`p-3 rounded-md border text-xs ${
            notice.error
              ? 'border-red-200 bg-red-50 text-red-800'
              : 'border-emerald-200 bg-emerald-50 text-emerald-900'
          }`}
        >
          {notice.msg}
        </div>
      )}

      {/* Completed Files Table */}
      <div className="bg-white border border-slate-200 rounded-md overflow-hidden">
        {visibleFiles.length === 0 ? (
          <div className="p-10 text-center space-y-2">
            <FileText className="w-8 h-8 text-slate-300 mx-auto" />
            <div className="text-sm font-medium text-slate-700">No completed files published yet</div>
            <p className="text-xs text-slate-500 max-w-md mx-auto">
              Completed files appear here once Controller 1 or Controller 2 publishes a finalized report or Excel workbook from their workspace.
            </p>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-sm">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-600">
                  <th className="px-4 py-3">Report Title &amp; File</th>
                  <th className="px-4 py-3">Module</th>
                  <th className="px-4 py-3">Period / Summary</th>
                  <th className="px-4 py-3">Published By</th>
                  <th className="px-4 py-3 text-right">Size</th>
                  <th className="px-4 py-3 text-right">Action</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {visibleFiles.map((f) => (
                  <tr key={f.id} className="hover:bg-slate-50">
                    <td className="px-4 py-3">
                      <div className="font-medium text-slate-900">{f.title}</div>
                      <div className="text-xs text-slate-500 font-data mt-0.5">
                        {f.fileName} · {f.format.toUpperCase()}
                      </div>
                    </td>
                    <td className="px-4 py-3 text-xs text-slate-700 whitespace-nowrap">
                      {MODULE_LABELS[f.module]}
                    </td>
                    <td className="px-4 py-3 text-xs text-slate-600">
                      {f.reportDate && (
                        <div className="font-medium text-slate-800 font-data">{f.reportDate}</div>
                      )}
                      {f.summary && <div className="text-slate-500 mt-0.5">{f.summary}</div>}
                    </td>
                    <td className="px-4 py-3 text-xs text-slate-600 whitespace-nowrap">
                      <div className="font-medium text-slate-800">{f.uploadedBy}</div>
                      <div className="text-slate-400 font-data">
                        {new Date(f.createdAt).toLocaleString()}
                      </div>
                    </td>
                    <td className="px-4 py-3 text-xs text-right font-data tabular-nums text-slate-600 whitespace-nowrap">
                      {formatBytes(f.size)}
                    </td>
                    <td className="px-4 py-3 text-right whitespace-nowrap">
                      <div className="inline-flex items-center gap-2">
                        <button
                          type="button"
                          disabled={downloadingId === f.id}
                          onClick={() => void handleDownload(f)}
                          className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-[#1F5143] hover:bg-[#173D32] text-white text-xs font-medium rounded-md transition-colors disabled:opacity-60"
                        >
                          <Download className="w-3.5 h-3.5" />
                          <span>{downloadingId === f.id ? 'Downloading...' : 'Download'}</span>
                        </button>
                        {userRole === 'Admin' && onDeleteFile && (
                          <button
                            type="button"
                            onClick={() => void onDeleteFile(f.id)}
                            className="p-1.5 border border-red-200 text-red-700 hover:bg-red-50 rounded-md transition-colors"
                            title="Delete completed file"
                            aria-label={`Delete ${f.fileName}`}
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};
