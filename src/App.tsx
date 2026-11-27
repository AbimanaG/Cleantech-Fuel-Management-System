import { useCallback, useEffect, useMemo, useState } from 'react';
import { LogOut, Menu } from 'lucide-react';
import { DEFAULT_VEHICLE_AREAS } from './data/defaultAreas';
import type {
  AbnormalItem,
  AuthUser,
  CardRecord,
  CommentStatus,
  CompletedFileModule,
  CompletedFileRecord,
  ConsolidatedDay,
  GpsFillRecord,
  GpsSummarySnapshot,
  ManagedUser,
  ManagementComment,
  ModuleTab,
  OfflineReason,
  UnusualFlag,
  VoucherRow,
  VoucherSummarySnapshot,
} from './types/fuel';
import { buildAutoAbnormalRows, buildGpsReportData } from './utils/gpsMatcherParsers';
import { convertReportToConsolidatedDays } from './utils/consolidationParser';
import { calcTotal, cleanMileage, determineShift, normFuelName } from './utils/voucherParser';
import {
  authApi,
  blobToBase64,
  commentsApi,
  completedFilesApi,
  setAuthToken,
  usersApi,
  workspaceApi,
} from './utils/apiClient';
import { LoginView } from './components/LoginView';
import { ROLE_ALLOWED_TABS, Sidebar } from './components/Sidebar';
import { GpsReconciliationView } from './components/GpsReconciliationView';
import { ConsolidationView } from './components/ConsolidationView';
import { VoucherAnalysisView } from './components/VoucherAnalysisView';
import { CompletedFilesView } from './components/CompletedFilesView';
import { ManagementReportsView } from './components/ManagementReportsView';
import { AreaDirectoryView } from './components/AreaDirectoryView';
import { UserManagementView } from './components/UserManagementView';

const AREA_STORE_KEY = 'fuelVoucherAnalysis.vehicleAreas.v1';
const VOUCHER_STORE_KEY = 'fuelVoucherAnalysis.v17';

const TAB_TITLES: Record<ModuleTab, string> = {
  gps: 'Fuel Card vs GPS Reconciliation',
  consolidation: 'Multi-Day Consolidation (Excel Format)',
  voucher: 'Fuel Voucher Analysis',
  files: 'Completed Files',
  management: 'Overview Reports & Comments',
  directory: 'Vehicle → Area Directory',
  users: 'User & Role Management',
};

export default function App() {
  const [bootState, setBootState] = useState<{
    checked: boolean;
    needsSetup: boolean;
    envAdminConfigured: boolean;
  }>({ checked: false, needsSetup: false, envAdminConfigured: false });
  const [authUser, setAuthUser] = useState<AuthUser | null>(null);
  const [activeTab, setActiveTab] = useState<ModuleTab>('gps');
  const [mobileNavOpen, setMobileNavOpen] = useState(false);

  // Shared server-backed state
  const [completedFiles, setCompletedFiles] = useState<CompletedFileRecord[]>([]);
  const [comments, setComments] = useState<ManagementComment[]>([]);
  const [managedUsers, setManagedUsers] = useState<ManagedUser[]>([]);
  const [remoteGpsSummary, setRemoteGpsSummary] = useState<GpsSummarySnapshot | null>(null);
  const [remoteVoucherSummary, setRemoteVoucherSummary] = useState<VoucherSummarySnapshot | null>(
    null
  );

  // Shared Vehicle -> Area directory
  const [vehicleAreas, setVehicleAreas] = useState<Record<string, string[]>>(() => {
    try {
      const raw = localStorage.getItem(AREA_STORE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') return parsed;
      }
    } catch {
      // ignore
    }
    return JSON.parse(JSON.stringify(DEFAULT_VEHICLE_AREAS));
  });

  // Module 1: GPS Reconciliation State
  const [cardRows, setCardRows] = useState<CardRecord[]>([]);
  const [gpsRows, setGpsRows] = useState<GpsFillRecord[]>([]);
  const [allocations, setAllocations] = useState<Record<string, string[]>>({});
  const [offlineReasons, setOfflineReasons] = useState<Record<string, OfflineReason>>({});
  const [skipped, setSkipped] = useState<Record<string, boolean>>({});
  const [unusualMap, setUnusualMap] = useState<Record<string, UnusualFlag>>({});
  const [abnormalItems, setAbnormalItems] = useState<AbnormalItem[] | null>(null);
  const [abnormalEdited, setAbnormalEdited] = useState(false);

  const gpsReport = useMemo(() => {
    if (!cardRows.length || !gpsRows.length) return null;
    return buildGpsReportData(cardRows, gpsRows, allocations, offlineReasons, skipped, unusualMap);
  }, [cardRows, gpsRows, allocations, offlineReasons, skipped, unusualMap]);

  // Module 2: Multi-Day Consolidation State
  const [consolidatedDays, setConsolidatedDays] = useState<Record<string, ConsolidatedDay>>({});

  // Module 3: Fuel Voucher Analysis State
  const initialVoucherState = useMemo(() => {
    try {
      const raw = localStorage.getItem(VOUCHER_STORE_KEY);
      if (raw) return JSON.parse(raw);
    } catch {
      // ignore
    }
    return null;
  }, []);

  const [pendingVouchers, setPendingVouchers] = useState<VoucherRow[]>(
    () => initialVoucherState?.pendingRows || []
  );
  const [appliedVouchers, setAppliedVouchers] = useState<VoucherRow[]>(
    () => initialVoucherState?.appliedRows || []
  );
  const [rowSeq, setRowSeq] = useState<number>(() => initialVoucherState?.rowSeq || 0);
  const [loadedSeqs, setLoadedSeqs] = useState<string[]>(() => initialVoucherState?.loadedSeqs || []);
  const [currentUser, setCurrentUser] = useState<string>(() => initialVoucherState?.currentUser || '');
  const [userId, setUserId] = useState<string>(() => initialVoucherState?.userId || '9362');
  const [selectedAppliedDays, setSelectedAppliedDays] = useState<string[]>(
    () => initialVoucherState?.selectedAppliedDays || []
  );
  const [selectedAppliedVouchers, setSelectedAppliedVouchers] = useState<string[]>(
    () => initialVoucherState?.selectedAppliedVouchers || []
  );

  useEffect(() => {
    authApi
      .bootstrapStatus()
      .then((res) =>
        setBootState({
          checked: true,
          needsSetup: res.needsSetup,
          envAdminConfigured: res.envAdminConfigured,
        })
      )
      .catch(() =>
        setBootState({ checked: true, needsSetup: false, envAdminConfigured: false })
      );
  }, []);

  const refreshServerData = useCallback(async (user: AuthUser) => {
    const tasks: Promise<void>[] = [
      completedFilesApi.list().then((r) => setCompletedFiles(r.files)),
    ];

    if (user.role !== 'Fuel Officer') {
      tasks.push(
        workspaceApi.get().then((w) => {
          if (w.vehicleAreas && typeof w.vehicleAreas === 'object') {
            setVehicleAreas(w.vehicleAreas);
          }
          if (w.consolidatedDays && typeof w.consolidatedDays === 'object') {
            setConsolidatedDays(w.consolidatedDays);
          }
          if (w.gpsSummary) setRemoteGpsSummary(w.gpsSummary);
          if (w.voucherSummary) setRemoteVoucherSummary(w.voucherSummary);
        })
      );
      tasks.push(commentsApi.list().then((r) => setComments(r.comments)));
    }

    if (user.role === 'Admin') {
      tasks.push(usersApi.list().then((r) => setManagedUsers(r.users)));
    }

    await Promise.allSettled(tasks);
  }, []);

  const handleAuthenticated = (user: AuthUser) => {
    setAuthUser(user);
    setBootState((prev) => ({ ...prev, needsSetup: false }));
    const allowed = ROLE_ALLOWED_TABS[user.role];
    setActiveTab(allowed[0]);
    if (!currentUser) {
      setCurrentUser(user.username);
    }
    void refreshServerData(user);
  };

  const handleSignOut = () => {
    setAuthToken(null);
    setAuthUser(null);
    setCompletedFiles([]);
    setComments([]);
    setManagedUsers([]);
    authApi
      .bootstrapStatus()
      .then((res) =>
        setBootState({
          checked: true,
          needsSetup: res.needsSetup,
          envAdminConfigured: res.envAdminConfigured,
        })
      )
      .catch(() => undefined);
  };

  // Persist voucher draft locally for operator continuity
  useEffect(() => {
    try {
      localStorage.setItem(
        VOUCHER_STORE_KEY,
        JSON.stringify({
          pendingRows: pendingVouchers,
          appliedRows: appliedVouchers,
          rowSeq,
          loadedSeqs,
          selectedAppliedDays,
          selectedAppliedVouchers,
          currentUser,
          userId,
        })
      );
    } catch {
      // ignore
    }
  }, [
    pendingVouchers,
    appliedVouchers,
    rowSeq,
    loadedSeqs,
    selectedAppliedDays,
    selectedAppliedVouchers,
    currentUser,
    userId,
  ]);

  // Sync GPS summary snapshot to server for Management reports
  useEffect(() => {
    if (!authUser || (authUser.role !== 'Controller 1' && authUser.role !== 'Admin')) return;
    if (!gpsReport || !cardRows.length) return;
    const snapshot: GpsSummarySnapshot = {
      updatedAt: new Date().toISOString(),
      updatedBy: authUser.username,
      cardCount: cardRows.length,
      gpsCount: gpsRows.length,
      matchedOkCount: gpsReport.matchedOk.length,
      varianceCount: gpsReport.matchedFlagged.length,
      unusualCount: gpsReport.matchedUnusual.length,
      cardOnlyCount: gpsReport.cardOnly.length,
      grandCardL: gpsReport.grandCard,
      grandGpsL: gpsReport.grandGps,
      grandVarianceL: gpsReport.grandVar,
    };
    setRemoteGpsSummary(snapshot);
    void workspaceApi.save({ gpsSummary: snapshot }).catch(() => undefined);
  }, [authUser, gpsReport, cardRows.length, gpsRows.length]);

  // Sync Voucher summary snapshot to server for Management reports
  useEffect(() => {
    if (!authUser || (authUser.role !== 'Controller 2' && authUser.role !== 'Admin')) return;
    const pendingLiters = pendingVouchers.reduce((s, r) => s + (parseFloat(r.liters) || 0), 0);
    const appliedLiters = appliedVouchers.reduce((s, r) => s + (parseFloat(r.liters) || 0), 0);
    const appliedCost = appliedVouchers.reduce((s, r) => s + (parseFloat(r.total_cost) || 0), 0);
    const voucherNumbers = [...new Set(appliedVouchers.map((r) => r.voucher).filter(Boolean))];
    const snapshot: VoucherSummarySnapshot = {
      updatedAt: new Date().toISOString(),
      updatedBy: authUser.username,
      pendingCount: pendingVouchers.length,
      appliedCount: appliedVouchers.length,
      pendingLiters,
      appliedLiters,
      appliedCost,
      voucherNumbers,
    };
    setRemoteVoucherSummary(snapshot);
    void workspaceApi.save({ voucherSummary: snapshot }).catch(() => undefined);
  }, [authUser, pendingVouchers, appliedVouchers]);

  const updateVehicleAreas = (next: Record<string, string[]>) => {
    setVehicleAreas(next);
    try {
      localStorage.setItem(AREA_STORE_KEY, JSON.stringify(next));
    } catch {
      // ignore
    }
    if (authUser?.role === 'Admin') {
      void workspaceApi.save({ vehicleAreas: next }).catch(() => undefined);
    }
  };

  const resetVehicleAreas = () => {
    const fresh = JSON.parse(JSON.stringify(DEFAULT_VEHICLE_AREAS));
    updateVehicleAreas(fresh);
  };

  const updateConsolidatedDays = (next: Record<string, ConsolidatedDay>) => {
    setConsolidatedDays(next);
    if (authUser && (authUser.role === 'Controller 1' || authUser.role === 'Admin')) {
      void workspaceApi.save({ consolidatedDays: next }).catch(() => undefined);
    }
  };

  const clearVoucherWorkspace = () => {
    setPendingVouchers([]);
    setAppliedVouchers([]);
    setLoadedSeqs([]);
    setSelectedAppliedDays([]);
    setSelectedAppliedVouchers([]);
    try {
      localStorage.removeItem(VOUCHER_STORE_KEY);
    } catch {
      // ignore
    }
  };

  // Cross-module sync: Controller 1 GPS Reconciliation -> Multi-Day Consolidation
  const handleSyncGpsToConsolidation = () => {
    if (!gpsReport || !cardRows.length) return;
    const activeAbnormal =
      abnormalItems !== null && abnormalEdited
        ? abnormalItems
        : buildAutoAbnormalRows(gpsReport).map((x, i) => ({ ...x, id: i + 1 }));
    const converted = convertReportToConsolidatedDays(cardRows, gpsReport, activeAbnormal);
    if (!converted.length) return;
    const next = { ...consolidatedDays };
    converted.forEach((day) => {
      next[day.date] = day;
    });
    updateConsolidatedDays(next);
    setActiveTab('consolidation');
  };

  // Cross-module sync (Admin only since Admin has both GPS & Voucher):
  const hasMobitelCardRows = useMemo(
    () => cardRows.some((c) => c.sheetName === 'Mobitel OneFuel'),
    [cardRows]
  );

  const handleSyncCardToVouchers = () => {
    if (!authUser || (authUser.role !== 'Admin' && authUser.role !== 'Controller 2')) return;
    const mobitelRows = cardRows.filter((c) => c.sheetName === 'Mobitel OneFuel');
    if (!mobitelRows.length) return;
    const seqSet = new Set(loadedSeqs);
    let seqCounter = rowSeq;
    const imported: VoucherRow[] = [];

    mobitelRows.forEach((c) => {
      const seq = String(c.no || '');
      if (seq && seqSet.has(seq)) return;
      const dt = `${c.date} ${c.time || '00:00'}`.trim();
      const row: VoucherRow = {
        __id: 'r' + seqCounter++,
        __seq: seq,
        __selected: false,
        __areaConfirmed: false,
        __isSplitChild: false,
        usage_id: '',
        vehicle: c.vehicle,
        fuel_type: c.fuel ? normFuelName(c.fuel) : '',
        usage_date: dt,
        liters: String(c.liters),
        shift: determineShift(dt),
        price: c.ppl != null && !isNaN(c.ppl) ? String(c.ppl) : '',
        total_cost: '',
        mileage: '',
        voucher: '',
        area: '',
        created_date: '',
        created_by: '',
        meter_start: '',
        meter_end: c.mileage != null ? cleanMileage(String(c.mileage)) : '',
      };
      calcTotal(row);
      if (seq) seqSet.add(seq);
      imported.push(row);
    });

    if (imported.length) {
      setRowSeq(seqCounter);
      setLoadedSeqs([...seqSet]);
      setPendingVouchers((prev) => [...prev, ...imported]);
    }
    setActiveTab('voucher');
  };

  const handlePublishCompletedFile = async (payload: {
    title: string;
    fileName: string;
    module: CompletedFileModule;
    format: 'xlsx' | 'pdf' | 'csv' | 'html';
    mimeType: string;
    blob: Blob;
    summary: string;
    reportDate: string;
  }) => {
    const contentBase64 = await blobToBase64(payload.blob);
    const res = await completedFilesApi.publish({
      title: payload.title,
      fileName: payload.fileName,
      module: payload.module,
      format: payload.format,
      mimeType: payload.mimeType,
      contentBase64,
      summary: payload.summary,
      reportDate: payload.reportDate,
    });
    setCompletedFiles((prev) => [res.file, ...prev]);
  };

  const handleDeleteCompletedFile = async (id: string) => {
    await completedFilesApi.remove(id);
    setCompletedFiles((prev) => prev.filter((f) => f.id !== id));
  };

  const handleAddComment = async (payload: {
    module: ManagementComment['module'];
    reportDate: string;
    status: CommentStatus;
    comment: string;
  }) => {
    const res = await commentsApi.create(payload);
    setComments((prev) => [res.comment, ...prev]);
  };

  const handleUpdateCommentStatus = async (id: string, status: CommentStatus) => {
    const res = await commentsApi.updateStatus(id, status);
    setComments((prev) => prev.map((c) => (c.id === id ? res.comment : c)));
  };

  const handleDeleteComment = async (id: string) => {
    await commentsApi.remove(id);
    setComments((prev) => prev.filter((c) => c.id !== id));
  };

  if (!bootState.checked) {
    return (
      <div className="min-h-screen bg-[#F8F9FA] flex items-center justify-center text-sm text-slate-500">
        Loading workspace...
      </div>
    );
  }

  if (!authUser) {
    return (
      <LoginView
        needsSetup={bootState.needsSetup}
        envAdminConfigured={bootState.envAdminConfigured}
        onAuthenticated={handleAuthenticated}
      />
    );
  }

  const allowedTabs = ROLE_ALLOWED_TABS[authUser.role];
  const currentTab: ModuleTab = allowedTabs.includes(activeTab) ? activeTab : allowedTabs[0];

  const gpsComments = comments.filter(
    (c) => c.module === 'General' || c.module === 'GPS Reconciliation'
  );
  const consolidationComments = comments.filter(
    (c) => c.module === 'General' || c.module === 'Multi-Day Consolidation'
  );
  const voucherComments = comments.filter(
    (c) => c.module === 'General' || c.module === 'Fuel Voucher Analysis'
  );

  return (
    <div className="min-h-screen flex bg-[#F8F9FA] text-[#0F172A]">
      <Sidebar
        activeTab={currentTab}
        onSelectTab={(t) => {
          if (allowedTabs.includes(t)) setActiveTab(t);
        }}
        userRole={authUser.role}
        counts={{
          cardCount: cardRows.length,
          consolidatedDays: Object.keys(consolidatedDays).length,
          voucherPending: pendingVouchers.length,
          voucherApplied: appliedVouchers.length,
          completedFiles: completedFiles.length,
          commentsCount: comments.length,
          directoryVehicles: Object.keys(vehicleAreas).length,
        }}
        mobileOpen={mobileNavOpen}
        onCloseMobile={() => setMobileNavOpen(false)}
      />

      <div className="flex-1 flex flex-col min-w-0">
        {/* Top Header Bar */}
        <header className="h-14 bg-white border-b border-slate-200 px-6 flex items-center justify-between gap-6 sticky top-0 z-30">
          <div className="flex items-center gap-3 min-w-0">
            <button
              type="button"
              onClick={() => setMobileNavOpen(true)}
              className="lg:hidden p-1.5 -ml-1.5 text-slate-600 hover:text-slate-900 rounded-md"
              aria-label="Open navigation"
            >
              <Menu className="w-5 h-5" />
            </button>
            <div className="flex items-center gap-2 text-sm truncate">
              <span className="text-slate-400 hidden sm:inline">Cleantech Fuel</span>
              <span className="text-slate-300 hidden sm:inline">/</span>
              <span className="font-semibold text-slate-900 truncate">{TAB_TITLES[currentTab]}</span>
            </div>
          </div>

          <div className="flex items-center gap-4 shrink-0 text-xs text-slate-600">
            <span className="hidden sm:inline whitespace-nowrap">
              <strong className="font-semibold text-slate-900">{authUser.username}</strong>
              <span className="text-slate-400 mx-1.5">·</span>
              <span>{authUser.role}</span>
            </span>
            <button
              type="button"
              onClick={handleSignOut}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-medium border border-slate-300 bg-white text-slate-800 hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0"
            >
              <LogOut className="w-3.5 h-3.5" />
              <span>Sign Out</span>
            </button>
          </div>
        </header>

        {/* Main Workspace Canvas */}
        <main className="flex-1 p-6 lg:p-8 max-w-[1440px] w-full mx-auto">
          {currentTab === 'gps' && (
            <GpsReconciliationView
              cardRows={cardRows}
              setCardRows={setCardRows}
              gpsRows={gpsRows}
              setGpsRows={setGpsRows}
              allocations={allocations}
              setAllocations={setAllocations}
              offlineReasons={offlineReasons}
              setOfflineReasons={setOfflineReasons}
              skipped={skipped}
              setSkipped={setSkipped}
              unusualMap={unusualMap}
              setUnusualMap={setUnusualMap}
              abnormalItems={abnormalItems}
              setAbnormalItems={setAbnormalItems}
              abnormalEdited={abnormalEdited}
              setAbnormalEdited={setAbnormalEdited}
              onPushToConsolidation={handleSyncGpsToConsolidation}
              onPushToVouchers={authUser.role === 'Admin' ? handleSyncCardToVouchers : undefined}
              onPublishCompletedFile={handlePublishCompletedFile}
              managementComments={gpsComments}
            />
          )}

          {currentTab === 'consolidation' && (
            <ConsolidationView
              days={consolidatedDays}
              onUpdateDays={updateConsolidatedDays}
              canSyncFromGps={Boolean(gpsReport && cardRows.length)}
              onSyncFromGps={handleSyncGpsToConsolidation}
              onPublishCompletedFile={handlePublishCompletedFile}
              managementComments={consolidationComments}
            />
          )}

          {currentTab === 'voucher' && (
            <VoucherAnalysisView
              pendingRows={pendingVouchers}
              setPendingRows={setPendingVouchers}
              appliedRows={appliedVouchers}
              setAppliedRows={setAppliedVouchers}
              loadedSeqs={loadedSeqs}
              setLoadedSeqs={setLoadedSeqs}
              rowSeq={rowSeq}
              setRowSeq={setRowSeq}
              currentUser={currentUser}
              setCurrentUser={setCurrentUser}
              userId={userId}
              setUserId={setUserId}
              selectedAppliedDays={selectedAppliedDays}
              setSelectedAppliedDays={setSelectedAppliedDays}
              selectedAppliedVouchers={selectedAppliedVouchers}
              setSelectedAppliedVouchers={setSelectedAppliedVouchers}
              vehicleAreas={vehicleAreas}
              canImportFromGps={authUser.role === 'Admin' && hasMobitelCardRows}
              onImportFromGps={handleSyncCardToVouchers}
              onClearWorkspace={clearVoucherWorkspace}
              onPublishCompletedFile={handlePublishCompletedFile}
              managementComments={voucherComments}
            />
          )}

          {currentTab === 'files' && (
            <CompletedFilesView
              files={completedFiles}
              userRole={authUser.role}
              onRefresh={() => refreshServerData(authUser)}
              onDeleteFile={authUser.role === 'Admin' ? handleDeleteCompletedFile : undefined}
            />
          )}

          {currentTab === 'management' && (
            <ManagementReportsView
              consolidatedDays={consolidatedDays}
              gpsSummary={remoteGpsSummary}
              voucherSummary={remoteVoucherSummary}
              completedFiles={completedFiles}
              comments={comments}
              currentUsername={authUser.username}
              userRole={authUser.role}
              onRefresh={() => refreshServerData(authUser)}
              onAddComment={handleAddComment}
              onUpdateCommentStatus={handleUpdateCommentStatus}
              onDeleteComment={handleDeleteComment}
            />
          )}

          {currentTab === 'directory' && (
            <AreaDirectoryView
              vehicleAreas={vehicleAreas}
              onUpdateVehicleAreas={updateVehicleAreas}
              onResetVehicleAreas={resetVehicleAreas}
            />
          )}

          {currentTab === 'users' && authUser.role === 'Admin' && (
            <UserManagementView
              users={managedUsers}
              currentUserId={authUser.id}
              onUsersChange={setManagedUsers}
            />
          )}
        </main>
      </div>
    </div>
  );
}
