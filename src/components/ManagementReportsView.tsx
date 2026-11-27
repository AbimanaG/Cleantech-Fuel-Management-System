import React, { useMemo, useState } from 'react';
import { Download, MessageSquarePlus, RefreshCw, Trash2 } from 'lucide-react';
import type {
  CommentStatus,
  CompletedFileRecord,
  ConsolidatedDay,
  GpsSummarySnapshot,
  ManagementComment,
  UserRole,
  VoucherSummarySnapshot,
} from '../types/fuel';
import { completedFilesApi } from '../utils/apiClient';

interface ManagementReportsViewProps {
  consolidatedDays: Record<string, ConsolidatedDay>;
  gpsSummary: GpsSummarySnapshot | null;
  voucherSummary: VoucherSummarySnapshot | null;
  completedFiles: CompletedFileRecord[];
  comments: ManagementComment[];
  currentUsername: string;
  userRole: UserRole;
  onRefresh: () => Promise<void>;
  onAddComment: (payload: {
    module: ManagementComment['module'];
    reportDate: string;
    status: CommentStatus;
    comment: string;
  }) => Promise<void>;
  onUpdateCommentStatus: (id: string, status: CommentStatus) => Promise<void>;
  onDeleteComment: (id: string) => Promise<void>;
}

const STATUS_STYLE: Record<CommentStatus, string> = {
  Approved: 'text-emerald-700 font-semibold',
  Reviewed: 'text-slate-700 font-semibold',
  'Action Required': 'text-red-700 font-semibold',
};

export const ManagementReportsView: React.FC<ManagementReportsViewProps> = ({
  consolidatedDays,
  gpsSummary,
  voucherSummary,
  completedFiles,
  comments,
  currentUsername,
  userRole,
  onRefresh,
  onAddComment,
  onUpdateCommentStatus,
  onDeleteComment,
}) => {
  const [moduleChoice, setModuleChoice] = useState<ManagementComment['module']>('General');
  const [reportDate, setReportDate] = useState('');
  const [statusChoice, setStatusChoice] = useState<CommentStatus>('Reviewed');
  const [commentText, setCommentText] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [refreshing, setRefreshing] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const sortedDays = useMemo(
    () =>
      Object.keys(consolidatedDays)
        .sort()
        .map((k) => consolidatedDays[k]),
    [consolidatedDays]
  );

  const consolidatedTotals = useMemo(() => {
    let diesel = 0;
    let petrol = 0;
    let superD = 0;
    let totalL = 0;
    let totalAmt = 0;
    let varianceEvents = 0;
    let unusualEvents = 0;
    let offlineEvents = 0;

    sortedDays.forEach((d) => {
      if (d.summary) {
        diesel += d.summary.diesel || 0;
        petrol += d.summary.petrol || 0;
        superD += d.summary.superD || 0;
        totalL += d.summary.total || 0;
        totalAmt += d.summary.totalAmt || 0;
      }
      varianceEvents += d.flagged?.length || 0;
      unusualEvents += d.unusual?.length || 0;
      offlineEvents += d.cardOnly?.length || 0;
    });

    return {
      diesel,
      petrol,
      superD,
      totalL,
      totalAmt,
      varianceEvents,
      unusualEvents,
      offlineEvents,
    };
  }, [sortedDays]);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    if (!commentText.trim()) {
      setError('Enter a management review comment.');
      return;
    }
    setSubmitting(true);
    try {
      await onAddComment({
        module: moduleChoice,
        reportDate: reportDate.trim(),
        status: statusChoice,
        comment: commentText.trim(),
      });
      setCommentText('');
      setReportDate('');
    } catch (err: any) {
      setError(err?.message || 'Failed to save comment.');
    } finally {
      setSubmitting(false);
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
    <div className="space-y-8">
      <div className="flex flex-col md:flex-row md:items-end justify-between gap-4 border-b border-slate-200 pb-5">
        <div>
          <h1 className="text-xl font-bold text-slate-900">Management Overview Reports &amp; Comments</h1>
          <p className="text-sm text-slate-600 mt-1">
            Review consolidated reconciliation totals, GPS variances, fuel voucher totals, published reports, and record management remarks.
          </p>
        </div>
        <button
          type="button"
          onClick={() => void handleRefresh()}
          disabled={refreshing}
          className="inline-flex items-center gap-2 px-3.5 py-2 bg-white hover:bg-slate-50 text-slate-800 border border-slate-300 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0"
        >
          <RefreshCw className={`w-3.5 h-3.5 ${refreshing ? 'animate-spin' : ''}`} />
          <span>{refreshing ? 'Refreshing...' : 'Refresh Reports'}</span>
        </button>
      </div>

      {/* Top Executive Metrics Strip */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="bg-white border border-slate-200 rounded-md p-5">
          <div className="text-xs font-medium text-slate-500">Multi-Day Consolidated Volume</div>
          <div className="mt-2 font-data text-2xl font-semibold text-slate-900 tabular-nums">
            {consolidatedTotals.totalL.toFixed(2)} L
          </div>
          <div className="mt-1.5 text-xs text-slate-500 font-data tabular-nums">
            {sortedDays.length} day(s) · Rs.{' '}
            {consolidatedTotals.totalAmt.toLocaleString('en-US', {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            })}
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-md p-5">
          <div className="text-xs font-medium text-slate-500">Active GPS Reconciliation</div>
          <div className="mt-2 font-data text-2xl font-semibold text-slate-900 tabular-nums">
            {gpsSummary ? `${gpsSummary.matchedOkCount} OK` : 'No Snapshot'}
          </div>
          <div className="mt-1.5 text-xs text-slate-500 font-data tabular-nums">
            {gpsSummary
              ? `${gpsSummary.varianceCount} variance · ${gpsSummary.unusualCount} unusual · Net ${gpsSummary.grandVarianceL.toFixed(2)} L`
              : 'Synced when Controller 1 runs GPS Matcher'}
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-md p-5">
          <div className="text-xs font-medium text-slate-500">Fuel Voucher Ledger</div>
          <div className="mt-2 font-data text-2xl font-semibold text-slate-900 tabular-nums">
            {voucherSummary ? `${voucherSummary.appliedCount} Applied` : 'No Snapshot'}
          </div>
          <div className="mt-1.5 text-xs text-slate-500 font-data tabular-nums">
            {voucherSummary
              ? `${voucherSummary.appliedLiters.toFixed(2)} L · Rs. ${voucherSummary.appliedCost.toLocaleString('en-US', {
                  minimumFractionDigits: 2,
                  maximumFractionDigits: 2,
                })}`
              : 'Synced when Controller 2 updates vouchers'}
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-md p-5">
          <div className="text-xs font-medium text-slate-500">Published Files &amp; Reviews</div>
          <div className="mt-2 font-data text-2xl font-semibold text-slate-900 tabular-nums">
            {completedFiles.length}{' '}
            <span className="text-sm font-normal text-slate-400">/ {comments.length} notes</span>
          </div>
          <div className="mt-1.5 text-xs text-slate-500">
            {consolidatedTotals.varianceEvents + consolidatedTotals.unusualEvents} variance flags ·{' '}
            {consolidatedTotals.offlineEvents} offline
          </div>
        </div>
      </div>

      {/* Consolidated Daily Summary Report Table */}
      <div className="bg-white border border-slate-200 rounded-md overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200 flex items-center justify-between">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              Daily Fuel Summary Report (Mobitel &amp; GPS Variance)
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Breakdown across all days loaded in Multi-Day Consolidation.
            </p>
          </div>
        </div>

        {sortedDays.length === 0 ? (
          <div className="p-8 text-center text-xs text-slate-500">
            No multi-day consolidation reports have been synced yet.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-sm">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-600">
                  <th className="px-4 py-2.5">Date</th>
                  <th className="px-4 py-2.5 text-right">Diesel (L)</th>
                  <th className="px-4 py-2.5 text-right">Petrol (L)</th>
                  <th className="px-4 py-2.5 text-right">Super D (L)</th>
                  <th className="px-4 py-2.5 text-right">Total (L)</th>
                  <th className="px-4 py-2.5 text-right">Amount (Rs.)</th>
                  <th className="px-4 py-2.5 text-right">±1.5 L Flags</th>
                  <th className="px-4 py-2.5 text-right">Offline</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 font-data tabular-nums text-xs">
                {sortedDays.map((d) => (
                  <tr key={d.date} className="hover:bg-slate-50">
                    <td className="px-4 py-2.5 font-semibold text-slate-900">{d.date}</td>
                    <td className="px-4 py-2.5 text-right">{(d.summary?.diesel || 0).toFixed(2)}</td>
                    <td className="px-4 py-2.5 text-right">{(d.summary?.petrol || 0).toFixed(2)}</td>
                    <td className="px-4 py-2.5 text-right">{(d.summary?.superD || 0).toFixed(2)}</td>
                    <td className="px-4 py-2.5 text-right font-semibold">
                      {(d.summary?.total || 0).toFixed(2)}
                    </td>
                    <td className="px-4 py-2.5 text-right">
                      {(d.summary?.totalAmt || 0).toLocaleString('en-US', {
                        minimumFractionDigits: 2,
                        maximumFractionDigits: 2,
                      })}
                    </td>
                    <td className="px-4 py-2.5 text-right text-amber-700">
                      {(d.flagged?.length || 0) + (d.unusual?.length || 0)}
                    </td>
                    <td className="px-4 py-2.5 text-right text-slate-600">
                      {d.cardOnly?.length || 0}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Bottom Split: Management Comments + Published Reports Quick Access */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* Management Comments */}
        <div className="lg:col-span-7 bg-white border border-slate-200 rounded-md p-6 space-y-5">
          <div className="border-b border-slate-100 pb-3">
            <h2 className="text-base font-semibold text-slate-900">Management Comments &amp; Sign-Off</h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Post review feedback, approvals, or action items visible to Controllers and Admin.
            </p>
          </div>

          <form onSubmit={handleSubmit} className="space-y-3 bg-slate-50 border border-slate-200 rounded-md p-4">
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">Report Area</label>
                <select
                  value={moduleChoice}
                  onChange={(e) => setModuleChoice(e.target.value as ManagementComment['module'])}
                  className="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded bg-white text-slate-900"
                >
                  <option value="General">General</option>
                  <option value="GPS Reconciliation">GPS Reconciliation</option>
                  <option value="Multi-Day Consolidation">Multi-Day Consolidation</option>
                  <option value="Fuel Voucher Analysis">Fuel Voucher Analysis</option>
                </select>
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Report Date / Period
                </label>
                <input
                  type="text"
                  value={reportDate}
                  onChange={(e) => setReportDate(e.target.value)}
                  placeholder="e.g. 2026-11-27 or Nov Week 4"
                  className="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded bg-white text-slate-900"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">Decision Status</label>
                <select
                  value={statusChoice}
                  onChange={(e) => setStatusChoice(e.target.value as CommentStatus)}
                  className="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded bg-white text-slate-900"
                >
                  <option value="Reviewed">Reviewed</option>
                  <option value="Approved">Approved</option>
                  <option value="Action Required">Action Required</option>
                </select>
              </div>
            </div>

            <div>
              <label className="block text-xs font-medium text-slate-700 mb-1">Comment</label>
              <textarea
                rows={3}
                value={commentText}
                onChange={(e) => setCommentText(e.target.value)}
                placeholder="Write management observations, variance queries, or approval remarks..."
                className="w-full px-3 py-2 text-sm border border-slate-300 rounded bg-white text-slate-900 focus:outline-none focus:ring-2 focus:ring-[#1F5143]/25"
              />
            </div>

            {error && <div className="text-xs text-red-700">{error}</div>}

            <div className="flex justify-end">
              <button
                type="submit"
                disabled={submitting}
                className="inline-flex items-center gap-1.5 px-4 py-2 bg-[#1F5143] hover:bg-[#173D32] text-white text-xs font-medium rounded-md transition-colors disabled:opacity-60"
              >
                <MessageSquarePlus className="w-3.5 h-3.5" />
                <span>{submitting ? 'Posting...' : 'Post Comment'}</span>
              </button>
            </div>
          </form>

          <div className="divide-y divide-slate-100 max-h-[420px] overflow-y-auto">
            {comments.length === 0 ? (
              <div className="py-8 text-center text-xs text-slate-500">
                No management comments recorded yet.
              </div>
            ) : (
              comments.map((c) => (
                <div key={c.id} className="py-3.5 space-y-1.5">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <div className="text-xs text-slate-600 flex flex-wrap items-center gap-2">
                      <span className="font-semibold text-slate-900">{c.author}</span>
                      <span aria-hidden="true">·</span>
                      <span>{c.module}</span>
                      {c.reportDate && (
                        <>
                          <span aria-hidden="true">·</span>
                          <span className="font-data text-slate-700">{c.reportDate}</span>
                        </>
                      )}
                      <span aria-hidden="true">·</span>
                      <span className={STATUS_STYLE[c.status]}>{c.status}</span>
                    </div>
                    <div className="flex items-center gap-2">
                      {(userRole === 'Admin' || c.author === currentUsername) && (
                        <>
                          <select
                            value={c.status}
                            onChange={(e) =>
                              void onUpdateCommentStatus(c.id, e.target.value as CommentStatus)
                            }
                            aria-label="Update comment status"
                            className="text-xs border border-slate-300 rounded px-2 py-0.5 bg-white text-slate-700"
                          >
                            <option value="Reviewed">Reviewed</option>
                            <option value="Approved">Approved</option>
                            <option value="Action Required">Action Required</option>
                          </select>
                          <button
                            type="button"
                            onClick={() => void onDeleteComment(c.id)}
                            className="p-1 text-slate-400 hover:text-red-600 rounded"
                            aria-label="Delete comment"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </>
                      )}
                    </div>
                  </div>
                  <p className="text-sm text-slate-800 whitespace-pre-wrap">{c.comment}</p>
                  <div className="text-[11px] text-slate-400 font-data">
                    {new Date(c.createdAt).toLocaleString()}
                  </div>
                </div>
              ))
            )}
          </div>
        </div>

        {/* Published Completed Files Quick Access */}
        <div className="lg:col-span-5 bg-white border border-slate-200 rounded-md p-6 space-y-4">
          <div className="border-b border-slate-100 pb-3">
            <h2 className="text-base font-semibold text-slate-900">Published Report Files</h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Download completed spreadsheets and PDFs for audit review.
            </p>
          </div>

          {completedFiles.length === 0 ? (
            <div className="py-8 text-center text-xs text-slate-500">
              No completed files published yet.
            </div>
          ) : (
            <div className="divide-y divide-slate-100 max-h-[480px] overflow-y-auto">
              {completedFiles.map((f) => (
                <div key={f.id} className="py-3 flex items-center justify-between gap-3">
                  <div className="min-w-0">
                    <div className="text-sm font-medium text-slate-900 truncate">{f.title}</div>
                    <div className="text-xs text-slate-500 mt-0.5">
                      {f.format.toUpperCase()} · {f.uploadedBy}
                      {f.reportDate ? ` · ${f.reportDate}` : ''}
                    </div>
                  </div>
                  <button
                    type="button"
                    onClick={() => void completedFilesApi.download(f)}
                    className="inline-flex items-center gap-1 px-2.5 py-1.5 border border-slate-300 hover:bg-slate-50 text-slate-800 text-xs font-medium rounded transition-colors shrink-0"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>Download</span>
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
