export type ModuleTab =
  | 'gps'
  | 'consolidation'
  | 'voucher'
  | 'files'
  | 'management'
  | 'directory'
  | 'users';

/* ============================================================
   MODULE 1: GPS vs FUEL CARD MATCHER
   ============================================================ */
export interface CardRecord {
  id: string;
  idx: number;
  depot: string;
  no: string | number;
  vehicle: string;
  vehicleKey: string;
  fuel: string;
  liters: number;
  shift: string;
  area: string;
  voucher: string;
  date: string;
  time: string;
  station?: string;
  status?: string;
  mileage?: number | null;
  amount?: number;
  payable?: number;
  ppl?: number;
  cardNo?: string;
  empName?: string;
  vehicleType?: string;
  division?: string;
  sourceFile: string;
  sheetName: string;
}

export interface GpsFillRecord {
  id: string;
  idx: number;
  no: string;
  vehicle: string;
  vehicleKey: string;
  description: string;
  date: string;
  time: string;
  location: string;
  liters: number;
  sourceFile: string;
}

export interface OfflineReason {
  category: string;
  detail: string;
}

export interface UnusualFlag {
  note: string;
  off?: boolean;
  auto?: boolean;
}

export interface MatchedEntry {
  unusual: UnusualFlag | null;
  card: CardRecord;
  gpsList: GpsFillRecord[];
  cardL: number;
  gpsL: number;
  variance: number;
  pct: number | null;
  dt: string;
  locations: string;
  status: 'ok' | 'mismatch';
}

export interface CardOnlyEntry {
  card: CardRecord;
  hasGps: boolean;
  skipped: boolean;
  note: string;
  category: string;
  detail: string;
  auto: boolean;
}

export interface GpsOnlyEntry {
  gps: GpsFillRecord;
}

export interface TimeVarianceRow {
  r: MatchedEntry;
  first: GpsFillRecord;
  mins: number;
}

export interface AbnormalItem {
  id: number;
  cat: string;
  veh: string;
  issue: string;
  litres: number;
  action: string;
}

export interface GpsReportData {
  matched: MatchedEntry[];
  matchedNormal: MatchedEntry[];
  matchedUnusual: MatchedEntry[];
  matchedFlagged: MatchedEntry[];
  matchedOk: MatchedEntry[];
  cardOnly: CardOnlyEntry[];
  gpsOnlyOnSite: GpsOnlyEntry[];
  gpsOnlyOffSite: GpsOnlyEntry[];
  grandCard: number;
  grandGps: number;
  grandVar: number;
  grandPct: number | null;
}

/* ============================================================
   MODULE 2: MULTI-DAY CONSOLIDATION
   ============================================================ */
export interface ConsolidatedDaySummary {
  date: string;
  diesel: number;
  dieselAmt: number;
  petrol: number;
  petrolAmt: number;
  superD: number;
  superAmt: number;
  other: number;
  otherAmt: number;
  total: number;
  totalAmt: number;
}

export interface ConsolidatedMatchedRow {
  date: string;
  cardTime: string;
  gpsTimes: string;
  vehicle: string;
  type: string;
  area: string;
  cardL: number;
  gpsL: number;
  variance: number;
  pct: number;
  status: string;
  note: string;
}

export interface ConsolidatedCardOnlyRow {
  date: string;
  cardTime: string;
  vehicle: string;
  type: string;
  shift: string;
  area: string;
  cardL: number;
  note: string;
  category: string;
}

export interface ConsolidatedTimeRow {
  vehicle: string;
  type: string;
  date: string;
  cardTime: string;
  gpsTime: string;
  mins: number;
  location: string;
  risk: string;
}

export interface ConsolidatedAbnormalRow {
  cat: string;
  vehicle: string;
  issue: string;
  litres: number;
  action: string;
  derived?: boolean;
}

export interface ConsolidatedCheck {
  label: string;
  got: number;
  want: number | null;
  ok: boolean;
}

export interface ConsolidatedDay {
  date: string;
  fileName: string;
  summary: ConsolidatedDaySummary | null;
  checks: ConsolidatedCheck[];
  warnings: string[];
  unusual: ConsolidatedMatchedRow[];
  flagged: ConsolidatedMatchedRow[];
  ok: ConsolidatedMatchedRow[];
  cardOnly: ConsolidatedCardOnlyRow[];
  timeVar: ConsolidatedTimeRow[];
  abnormal: ConsolidatedAbnormalRow[];
  counts: { s5: number | null; s51: number | null; ok: number | null };
  grand: { cardL: number; gpsL: number; variance: number; pct: number } | null;
  abnormalMissing: boolean;
}

/* ============================================================
   MODULE 3: FUEL VOUCHER ANALYSIS
   ============================================================ */
export interface SplitLineDraft {
  label: string;
  liters: string;
  meter_start: string;
  meter_end: string;
}

export interface VoucherRow {
  __id: string;
  __seq: string;
  __serial?: number;
  __selected: boolean;
  __areaConfirmed: boolean;
  __isSplitChild: boolean;
  __splitParentVehicle?: string;
  __splitDraft?: SplitLineDraft[];
  usage_id: string;
  vehicle: string;
  fuel_type: string;
  usage_date: string;
  liters: string;
  shift: 'DAY' | 'NIGHT';
  price: string;
  total_cost: string;
  mileage: string;
  voucher: string;
  area: string;
  created_date: string;
  created_by: string;
  meter_start: string;
  meter_end: string;
}

export interface UploadSummaryInfo {
  taken: number;
  dups: number;
  bad: { seq: string; why: string }[];
  missing: string[];
  reportCount: number | null;
  reportLitres: number | null;
  reportAmount: number | null;
  sumLitres: number;
  sumAmount: number;
  fuelCounts: Record<string, number>;
  commonCount: number;
  noMileage: number;
}

export type UserRole =
  | 'Fuel Officer'
  | 'Controller 1'
  | 'Controller 2'
  | 'Admin'
  | 'Management';

export interface AuthUser {
  id: string;
  username: string;
  role: UserRole;
}

export interface ManagedUser extends AuthUser {
  active: boolean;
  createdAt: string;
}

export type CompletedFileModule = 'gps' | 'consolidation' | 'voucher';

export interface CompletedFileRecord {
  id: string;
  title: string;
  fileName: string;
  module: CompletedFileModule;
  format: 'xlsx' | 'pdf' | 'csv' | 'html';
  mimeType: string;
  size: number;
  summary: string;
  reportDate: string;
  uploadedBy: string;
  uploadedByRole: UserRole;
  createdAt: string;
}

export type CommentStatus = 'Reviewed' | 'Approved' | 'Action Required';

export interface ManagementComment {
  id: string;
  module: 'General' | 'GPS Reconciliation' | 'Multi-Day Consolidation' | 'Fuel Voucher Analysis';
  reportDate: string;
  status: CommentStatus;
  comment: string;
  author: string;
  authorRole: UserRole;
  createdAt: string;
  updatedAt?: string;
}

export interface GpsSummarySnapshot {
  updatedAt: string;
  updatedBy: string;
  cardCount: number;
  gpsCount: number;
  matchedOkCount: number;
  varianceCount: number;
  unusualCount: number;
  cardOnlyCount: number;
  grandCardL: number;
  grandGpsL: number;
  grandVarianceL: number;
}

export interface VoucherSummarySnapshot {
  updatedAt: string;
  updatedBy: string;
  pendingCount: number;
  appliedCount: number;
  pendingLiters: number;
  appliedLiters: number;
  appliedCost: number;
  voucherNumbers: string[];
}

