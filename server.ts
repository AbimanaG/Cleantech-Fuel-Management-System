import 'dotenv/config';
import express, { type NextFunction, type Request, type Response } from 'express';
import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

type UserRole = 'Fuel Officer' | 'Controller 1' | 'Controller 2' | 'Admin' | 'Management';
type CompletedFileModule = 'gps' | 'consolidation' | 'voucher';
type FileFormat = 'xlsx' | 'pdf' | 'csv' | 'html';
type CommentStatus = 'Reviewed' | 'Approved' | 'Action Required';

interface StoredUser {
  id: string;
  username: string;
  role: UserRole;
  salt: string;
  passwordHash: string;
  active: boolean;
  createdAt: string;
}

interface CompletedFileMeta {
  id: string;
  title: string;
  fileName: string;
  storedName: string;
  module: CompletedFileModule;
  format: FileFormat;
  mimeType: string;
  size: number;
  summary: string;
  reportDate: string;
  uploadedBy: string;
  uploadedByRole: UserRole;
  createdAt: string;
}

interface ManagementCommentRecord {
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

interface StoreData {
  users: StoredUser[];
  completedFiles: CompletedFileMeta[];
  comments: ManagementCommentRecord[];
  workspace: {
    vehicleAreas: Record<string, string[]> | null;
    consolidatedDays: Record<string, unknown>;
    gpsSummary: Record<string, unknown> | null;
    voucherSummary: Record<string, unknown> | null;
  };
}

interface TokenPayload {
  sub: string;
  username: string;
  role: UserRole;
  exp: number;
}

interface AuthedRequest extends Request {
  user?: TokenPayload;
}

const VALID_ROLES: UserRole[] = [
  'Fuel Officer',
  'Controller 1',
  'Controller 2',
  'Admin',
  'Management',
];
const VALID_MODULES: CompletedFileModule[] = ['gps', 'consolidation', 'voucher'];
const VALID_FORMATS: FileFormat[] = ['xlsx', 'pdf', 'csv', 'html'];
const VALID_COMMENT_MODULES = [
  'General',
  'GPS Reconciliation',
  'Multi-Day Consolidation',
  'Fuel Voucher Analysis',
];
const VALID_COMMENT_STATUSES: CommentStatus[] = ['Reviewed', 'Approved', 'Action Required'];

const DATA_DIR = path.resolve(__dirname, '.data');
const FILES_DIR = path.resolve(DATA_DIR, 'files');
const DB_PATH = path.resolve(DATA_DIR, 'cleantech-fuel-db.json');

fs.mkdirSync(FILES_DIR, { recursive: true });

const SESSION_SECRET =
  (process.env.SESSION_SECRET || '').trim() || crypto.randomBytes(32).toString('hex');

function getEnvAdminPasscode(): string {
  return (process.env.ADMIN_PASSCODE || '').trim();
}

function defaultStore(): StoreData {
  return {
    users: [],
    completedFiles: [],
    comments: [],
    workspace: {
      vehicleAreas: null,
      consolidatedDays: {},
      gpsSummary: null,
      voucherSummary: null,
    },
  };
}

let writeQueue = Promise.resolve();

function loadStore(): StoreData {
  try {
    if (!fs.existsSync(DB_PATH)) return defaultStore();
    const raw = fs.readFileSync(DB_PATH, 'utf8');
    const parsed = JSON.parse(raw) as Partial<StoreData>;
    return {
      users: Array.isArray(parsed.users) ? parsed.users : [],
      completedFiles: Array.isArray(parsed.completedFiles) ? parsed.completedFiles : [],
      comments: Array.isArray(parsed.comments) ? parsed.comments : [],
      workspace: {
        vehicleAreas:
          parsed.workspace?.vehicleAreas && typeof parsed.workspace.vehicleAreas === 'object'
            ? parsed.workspace.vehicleAreas
            : null,
        consolidatedDays:
          parsed.workspace?.consolidatedDays &&
          typeof parsed.workspace.consolidatedDays === 'object'
            ? parsed.workspace.consolidatedDays
            : {},
        gpsSummary:
          parsed.workspace?.gpsSummary && typeof parsed.workspace.gpsSummary === 'object'
            ? parsed.workspace.gpsSummary
            : null,
        voucherSummary:
          parsed.workspace?.voucherSummary && typeof parsed.workspace.voucherSummary === 'object'
            ? parsed.workspace.voucherSummary
            : null,
      },
    };
  } catch {
    return defaultStore();
  }
}

function saveStore(store: StoreData): Promise<void> {
  const payload = JSON.stringify(store, null, 2);
  writeQueue = writeQueue.then(async () => {
    const tmp = `${DB_PATH}.${crypto.randomBytes(6).toString('hex')}.tmp`;
    await fs.promises.writeFile(tmp, payload, 'utf8');
    await fs.promises.rename(tmp, DB_PATH);
  });
  return writeQueue;
}

function hashPassword(password: string): { salt: string; hash: string } {
  const salt = crypto.randomBytes(16).toString('hex');
  const hash = crypto.scryptSync(password, salt, 64).toString('hex');
  return { salt, hash };
}

function verifyPassword(password: string, salt: string, expectedHex: string): boolean {
  try {
    const derived = crypto.scryptSync(password, salt, 64);
    const expected = Buffer.from(expectedHex, 'hex');
    if (derived.length !== expected.length) return false;
    return crypto.timingSafeEqual(derived, expected);
  } catch {
    return false;
  }
}

function safeEqualText(a: string, b: string): boolean {
  const da = crypto.createHash('sha256').update(a, 'utf8').digest();
  const db = crypto.createHash('sha256').update(b, 'utf8').digest();
  return crypto.timingSafeEqual(da, db);
}

function signToken(payload: Omit<TokenPayload, 'exp'>): string {
  const full: TokenPayload = {
    ...payload,
    exp: Date.now() + 12 * 60 * 60 * 1000,
  };
  const body = Buffer.from(JSON.stringify(full), 'utf8').toString('base64url');
  const sig = crypto.createHmac('sha256', SESSION_SECRET).update(body).digest('base64url');
  return `${body}.${sig}`;
}

function verifyToken(token: string): TokenPayload | null {
  try {
    const parts = token.split('.');
    if (parts.length !== 2) return null;
    const [body, sig] = parts;
    const expectedSig = crypto
      .createHmac('sha256', SESSION_SECRET)
      .update(body)
      .digest('base64url');
    const a = Buffer.from(sig, 'utf8');
    const b = Buffer.from(expectedSig, 'utf8');
    if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
    const parsed = JSON.parse(Buffer.from(body, 'base64url').toString('utf8')) as TokenPayload;
    if (!parsed || typeof parsed.exp !== 'number' || parsed.exp < Date.now()) return null;
    if (!VALID_ROLES.includes(parsed.role)) return null;
    return parsed;
  } catch {
    return null;
  }
}

function sanitizeUsername(raw: unknown): string | null {
  if (typeof raw !== 'string') return null;
  const trimmed = raw.trim();
  if (trimmed.length < 2 || trimmed.length > 48) return null;
  if (!/^[\w .@-]+$/.test(trimmed)) return null;
  return trimmed;
}

function sanitizePassword(raw: unknown): string | null {
  if (typeof raw !== 'string') return null;
  if (raw.length < 6 || raw.length > 128) return null;
  return raw;
}

function sanitizeFileName(raw: unknown, fallbackExt: string): string {
  const base = path.basename(typeof raw === 'string' ? raw : `report.${fallbackExt}`);
  const cleaned = base.replace(/[^a-zA-Z0-9._-]/g, '_').replace(/^\.+/, '');
  return (cleaned || `report.${fallbackExt}`).slice(0, 120);
}

function requireAuth(req: AuthedRequest, res: Response, next: NextFunction): void {
  const authHeader = req.headers.authorization || '';
  if (!authHeader.startsWith('Bearer ')) {
    res.status(401).json({ error: 'Authentication required.' });
    return;
  }
  const token = authHeader.slice(7).trim();
  const payload = verifyToken(token);
  if (!payload) {
    res.status(401).json({ error: 'Invalid or expired session. Please sign in again.' });
    return;
  }

  // Re-verify active status for stored accounts
  if (payload.sub !== 'env-admin') {
    const store = loadStore();
    const user = store.users.find((u) => u.id === payload.sub);
    if (!user || !user.active) {
      res.status(401).json({ error: 'Account is inactive or no longer exists.' });
      return;
    }
    req.user = {
      sub: user.id,
      username: user.username,
      role: user.role,
      exp: payload.exp,
    };
  } else {
    if (!getEnvAdminPasscode()) {
      res.status(401).json({ error: 'Session expired.' });
      return;
    }
    req.user = payload;
  }
  next();
}

function requireRoles(...roles: UserRole[]) {
  return (req: AuthedRequest, res: Response, next: NextFunction): void => {
    if (!req.user || !roles.includes(req.user.role)) {
      res.status(403).json({ error: 'You do not have permission to perform this action.' });
      return;
    }
    next();
  };
}

async function startServer() {
  const app = express();
  app.disable('x-powered-by');
  app.use(express.json({ limit: '25mb' }));

  /* ---------- AUTH ROUTES ---------- */
  app.get('/api/auth/bootstrap-status', (_req: Request, res: Response) => {
    const store = loadStore();
    const hasAdmin = store.users.some((u) => u.role === 'Admin' && u.active);
    const envAdminConfigured = Boolean(getEnvAdminPasscode());
    res.json({
      needsSetup: !hasAdmin && !envAdminConfigured,
      envAdminConfigured,
    });
  });

  app.post('/api/auth/claim-admin', async (req: Request, res: Response) => {
    const store = loadStore();
    const hasAdmin = store.users.some((u) => u.role === 'Admin' && u.active);
    if (hasAdmin || getEnvAdminPasscode()) {
      res.status(409).json({ error: 'Administrator account is already configured.' });
      return;
    }

    const username = sanitizeUsername(req.body?.username);
    const password = sanitizePassword(req.body?.password);
    if (!username) {
      res.status(400).json({ error: 'Enter a valid User Name (2–48 characters).' });
      return;
    }
    if (!password) {
      res.status(400).json({ error: 'Password must be at least 6 characters long.' });
      return;
    }

    const { salt, hash } = hashPassword(password);
    const adminUser: StoredUser = {
      id: `usr_${crypto.randomUUID().replace(/-/g, '').slice(0, 12)}`,
      username,
      role: 'Admin',
      salt,
      passwordHash: hash,
      active: true,
      createdAt: new Date().toISOString(),
    };
    store.users.push(adminUser);
    await saveStore(store);

    const token = signToken({
      sub: adminUser.id,
      username: adminUser.username,
      role: adminUser.role,
    });
    res.status(201).json({
      token,
      user: { id: adminUser.id, username: adminUser.username, role: adminUser.role },
    });
  });

  app.post('/api/auth/login', (req: Request, res: Response) => {
    const username = typeof req.body?.username === 'string' ? req.body.username.trim() : '';
    const password = typeof req.body?.password === 'string' ? req.body.password : '';
    if (!username || !password) {
      res.status(400).json({ error: 'Enter both User Name and Password.' });
      return;
    }

    const envPass = getEnvAdminPasscode();
    if (envPass && username.toLowerCase() === 'admin' && safeEqualText(password, envPass)) {
      const token = signToken({ sub: 'env-admin', username: 'admin', role: 'Admin' });
      res.json({
        token,
        user: { id: 'env-admin', username: 'admin', role: 'Admin' },
      });
      return;
    }

    const store = loadStore();
    const user = store.users.find(
      (u) => u.username.toLowerCase() === username.toLowerCase()
    );
    if (!user || !verifyPassword(password, user.salt, user.passwordHash)) {
      res.status(401).json({ error: 'Invalid User Name or Password.' });
      return;
    }
    if (!user.active) {
      res.status(403).json({ error: 'This user account is disabled. Contact Admin.' });
      return;
    }

    const token = signToken({
      sub: user.id,
      username: user.username,
      role: user.role,
    });
    res.json({
      token,
      user: { id: user.id, username: user.username, role: user.role },
    });
  });

  app.get('/api/auth/me', requireAuth, (req: AuthedRequest, res: Response) => {
    res.json({
      user: {
        id: req.user!.sub,
        username: req.user!.username,
        role: req.user!.role,
      },
    });
  });

  /* ---------- USER MANAGEMENT (ADMIN ONLY) ---------- */
  app.get(
    '/api/users',
    requireAuth,
    requireRoles('Admin'),
    (_req: AuthedRequest, res: Response) => {
      const store = loadStore();
      res.json({
        users: store.users.map((u) => ({
          id: u.id,
          username: u.username,
          role: u.role,
          active: u.active,
          createdAt: u.createdAt,
        })),
      });
    }
  );

  app.post(
    '/api/users',
    requireAuth,
    requireRoles('Admin'),
    async (req: AuthedRequest, res: Response) => {
      const username = sanitizeUsername(req.body?.username);
      const password = sanitizePassword(req.body?.password);
      const role = req.body?.role as UserRole;

      if (!username) {
        res.status(400).json({ error: 'User Name must be 2–48 valid characters.' });
        return;
      }
      if (!password) {
        res.status(400).json({ error: 'Password must be at least 6 characters.' });
        return;
      }
      if (!VALID_ROLES.includes(role)) {
        res.status(400).json({ error: 'Select a valid role.' });
        return;
      }

      const store = loadStore();
      if (store.users.some((u) => u.username.toLowerCase() === username.toLowerCase())) {
        res.status(409).json({ error: 'A user with that User Name already exists.' });
        return;
      }

      const { salt, hash } = hashPassword(password);
      const created: StoredUser = {
        id: `usr_${crypto.randomUUID().replace(/-/g, '').slice(0, 12)}`,
        username,
        role,
        salt,
        passwordHash: hash,
        active: true,
        createdAt: new Date().toISOString(),
      };
      store.users.push(created);
      await saveStore(store);

      res.status(201).json({
        user: {
          id: created.id,
          username: created.username,
          role: created.role,
          active: created.active,
          createdAt: created.createdAt,
        },
      });
    }
  );

  app.patch(
    '/api/users/:id',
    requireAuth,
    requireRoles('Admin'),
    async (req: AuthedRequest, res: Response) => {
      const { id } = req.params;
      const store = loadStore();
      const target = store.users.find((u) => u.id === id);
      if (!target) {
        res.status(404).json({ error: 'User not found.' });
        return;
      }

      const nextRole = req.body?.role as UserRole | undefined;
      const nextActive = typeof req.body?.active === 'boolean' ? req.body.active : undefined;
      const nextPassword =
        typeof req.body?.password === 'string' && req.body.password.length > 0
          ? sanitizePassword(req.body.password)
          : undefined;

      if (req.body?.password && !nextPassword) {
        res.status(400).json({ error: 'New password must be at least 6 characters.' });
        return;
      }
      if (nextRole !== undefined && !VALID_ROLES.includes(nextRole)) {
        res.status(400).json({ error: 'Invalid role specified.' });
        return;
      }

      // Prevent disabling or demoting the last active Admin
      const activeAdmins = store.users.filter((u) => u.role === 'Admin' && u.active);
      const isTargetLastAdmin =
        target.role === 'Admin' &&
        target.active &&
        activeAdmins.length <= 1 &&
        !getEnvAdminPasscode();
      if (isTargetLastAdmin && ((nextRole && nextRole !== 'Admin') || nextActive === false)) {
        res.status(400).json({ error: 'Cannot disable or demote the last active Admin.' });
        return;
      }

      if (nextRole) target.role = nextRole;
      if (nextActive !== undefined) target.active = nextActive;
      if (nextPassword) {
        const { salt, hash } = hashPassword(nextPassword);
        target.salt = salt;
        target.passwordHash = hash;
      }

      await saveStore(store);
      res.json({
        user: {
          id: target.id,
          username: target.username,
          role: target.role,
          active: target.active,
          createdAt: target.createdAt,
        },
      });
    }
  );

  app.delete(
    '/api/users/:id',
    requireAuth,
    requireRoles('Admin'),
    async (req: AuthedRequest, res: Response) => {
      const { id } = req.params;
      if (req.user?.sub === id) {
        res.status(400).json({ error: 'You cannot delete your own active session account.' });
        return;
      }
      const store = loadStore();
      const idx = store.users.findIndex((u) => u.id === id);
      if (idx === -1) {
        res.status(404).json({ error: 'User not found.' });
        return;
      }
      const target = store.users[idx];
      const activeAdmins = store.users.filter((u) => u.role === 'Admin' && u.active);
      if (
        target.role === 'Admin' &&
        target.active &&
        activeAdmins.length <= 1 &&
        !getEnvAdminPasscode()
      ) {
        res.status(400).json({ error: 'Cannot delete the last active Admin account.' });
        return;
      }
      store.users.splice(idx, 1);
      await saveStore(store);
      res.json({ ok: true });
    }
  );

  /* ---------- COMPLETED FILES LIBRARY ---------- */
  app.get('/api/completed-files', requireAuth, (req: AuthedRequest, res: Response) => {
    const store = loadStore();
    const role = req.user!.role;
    let files = store.completedFiles;

    if (role === 'Controller 1') {
      files = files.filter((f) => f.module === 'gps' || f.module === 'consolidation');
    } else if (role === 'Controller 2') {
      files = files.filter((f) => f.module === 'voucher');
    }
    // Fuel Officer, Management, and Admin see all completed files

    const sorted = [...files]
      .sort((a, b) => b.createdAt.localeCompare(a.createdAt))
      .map(({ storedName: _omit, ...publicMeta }) => publicMeta);
    res.json({ files: sorted });
  });

  app.post(
    '/api/completed-files',
    requireAuth,
    requireRoles('Controller 1', 'Controller 2', 'Admin'),
    async (req: AuthedRequest, res: Response) => {
      const role = req.user!.role;
      const mod = req.body?.module as CompletedFileModule;
      const format = req.body?.format as FileFormat;
      const contentBase64 = typeof req.body?.contentBase64 === 'string' ? req.body.contentBase64 : '';

      if (!VALID_MODULES.includes(mod)) {
        res.status(400).json({ error: 'Invalid report module.' });
        return;
      }
      if (role === 'Controller 1' && mod !== 'gps' && mod !== 'consolidation') {
        res.status(403).json({
          error: 'Controller 1 can only publish GPS Reconciliation or Multi-Day Consolidation files.',
        });
        return;
      }
      if (role === 'Controller 2' && mod !== 'voucher') {
        res.status(403).json({ error: 'Controller 2 can only publish Fuel Voucher Analysis files.' });
        return;
      }
      if (!VALID_FORMATS.includes(format)) {
        res.status(400).json({ error: 'Unsupported file format.' });
        return;
      }
      if (!contentBase64) {
        res.status(400).json({ error: 'File payload is empty.' });
        return;
      }

      const buf = Buffer.from(contentBase64, 'base64');
      if (!buf.length || buf.length > 20 * 1024 * 1024) {
        res.status(400).json({ error: 'File size must be between 1 byte and 20 MB.' });
        return;
      }

      const id = `file_${crypto.randomUUID().replace(/-/g, '').slice(0, 14)}`;
      const fileName = sanitizeFileName(req.body?.fileName, format);
      const storedName = `${id}.${format}`;
      const targetPath = path.resolve(FILES_DIR, storedName);
      if (!targetPath.startsWith(FILES_DIR + path.sep)) {
        res.status(400).json({ error: 'Invalid storage path.' });
        return;
      }

      await fs.promises.writeFile(targetPath, buf);

      const mimeMap: Record<FileFormat, string> = {
        xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        pdf: 'application/pdf',
        csv: 'text/csv;charset=utf-8',
        html: 'text/html;charset=utf-8',
      };

      const title =
        typeof req.body?.title === 'string' && req.body.title.trim()
          ? req.body.title.trim().slice(0, 140)
          : fileName;
      const summary =
        typeof req.body?.summary === 'string' ? req.body.summary.trim().slice(0, 240) : '';
      const reportDate =
        typeof req.body?.reportDate === 'string' ? req.body.reportDate.trim().slice(0, 40) : '';

      const meta: CompletedFileMeta = {
        id,
        title,
        fileName,
        storedName,
        module: mod,
        format,
        mimeType:
          typeof req.body?.mimeType === 'string' && req.body.mimeType
            ? req.body.mimeType.slice(0, 120)
            : mimeMap[format],
        size: buf.length,
        summary,
        reportDate,
        uploadedBy: req.user!.username,
        uploadedByRole: req.user!.role,
        createdAt: new Date().toISOString(),
      };

      const store = loadStore();
      store.completedFiles.push(meta);
      await saveStore(store);

      const { storedName: _omit, ...publicMeta } = meta;
      res.status(201).json({ file: publicMeta });
    }
  );

  app.get('/api/completed-files/:id/download', requireAuth, (req: AuthedRequest, res: Response) => {
    const id = req.params.id;
    if (!/^[a-zA-Z0-9_-]{4,64}$/.test(id)) {
      res.status(400).json({ error: 'Invalid file identifier.' });
      return;
    }

    const store = loadStore();
    const meta = store.completedFiles.find((f) => f.id === id);
    if (!meta) {
      res.status(404).json({ error: 'Completed file not found.' });
      return;
    }

    const role = req.user!.role;
    if (role === 'Controller 1' && meta.module === 'voucher') {
      res.status(403).json({ error: 'Access denied.' });
      return;
    }
    if (role === 'Controller 2' && meta.module !== 'voucher') {
      res.status(403).json({ error: 'Access denied.' });
      return;
    }

    const resolved = path.resolve(FILES_DIR, meta.storedName);
    if (!resolved.startsWith(FILES_DIR + path.sep) || !fs.existsSync(resolved)) {
      res.status(404).json({ error: 'Stored file missing on server.' });
      return;
    }

    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('Content-Type', meta.mimeType || 'application/octet-stream');
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${sanitizeFileName(meta.fileName, meta.format)}"`
    );
    fs.createReadStream(resolved).pipe(res);
  });

  app.delete(
    '/api/completed-files/:id',
    requireAuth,
    requireRoles('Admin', 'Controller 1', 'Controller 2'),
    async (req: AuthedRequest, res: Response) => {
      const id = req.params.id;
      if (!/^[a-zA-Z0-9_-]{4,64}$/.test(id)) {
        res.status(400).json({ error: 'Invalid file identifier.' });
        return;
      }

      const store = loadStore();
      const idx = store.completedFiles.findIndex((f) => f.id === id);
      if (idx === -1) {
        res.status(404).json({ error: 'File not found.' });
        return;
      }
      const meta = store.completedFiles[idx];
      const role = req.user!.role;
      if (role !== 'Admin') {
        if (role === 'Controller 1' && meta.module === 'voucher') {
          res.status(403).json({ error: 'Permission denied.' });
          return;
        }
        if (role === 'Controller 2' && meta.module !== 'voucher') {
          res.status(403).json({ error: 'Permission denied.' });
          return;
        }
      }

      const resolved = path.resolve(FILES_DIR, meta.storedName);
      if (resolved.startsWith(FILES_DIR + path.sep) && fs.existsSync(resolved)) {
        await fs.promises.unlink(resolved).catch(() => undefined);
      }
      store.completedFiles.splice(idx, 1);
      await saveStore(store);
      res.json({ ok: true });
    }
  );

  /* ---------- MANAGEMENT COMMENTS ---------- */
  app.get(
    '/api/comments',
    requireAuth,
    requireRoles('Management', 'Admin', 'Controller 1', 'Controller 2'),
    (req: AuthedRequest, res: Response) => {
      const store = loadStore();
      const role = req.user!.role;
      let list = [...store.comments].sort((a, b) => b.createdAt.localeCompare(a.createdAt));

      if (role === 'Controller 1') {
        list = list.filter(
          (c) =>
            c.module === 'General' ||
            c.module === 'GPS Reconciliation' ||
            c.module === 'Multi-Day Consolidation'
        );
      } else if (role === 'Controller 2') {
        list = list.filter((c) => c.module === 'General' || c.module === 'Fuel Voucher Analysis');
      }

      res.json({ comments: list });
    }
  );

  app.post(
    '/api/comments',
    requireAuth,
    requireRoles('Management', 'Admin'),
    async (req: AuthedRequest, res: Response) => {
      const mod = req.body?.module;
      const status = (req.body?.status || 'Reviewed') as CommentStatus;
      const commentText = typeof req.body?.comment === 'string' ? req.body.comment.trim() : '';
      const reportDate =
        typeof req.body?.reportDate === 'string' ? req.body.reportDate.trim().slice(0, 40) : '';

      if (!VALID_COMMENT_MODULES.includes(mod)) {
        res.status(400).json({ error: 'Select a valid report category.' });
        return;
      }
      if (!VALID_COMMENT_STATUSES.includes(status)) {
        res.status(400).json({ error: 'Invalid status selected.' });
        return;
      }
      if (!commentText || commentText.length > 2000) {
        res.status(400).json({ error: 'Comment must be between 1 and 2,000 characters.' });
        return;
      }

      const item: ManagementCommentRecord = {
        id: `cmt_${crypto.randomUUID().replace(/-/g, '').slice(0, 12)}`,
        module: mod,
        reportDate,
        status,
        comment: commentText,
        author: req.user!.username,
        authorRole: req.user!.role,
        createdAt: new Date().toISOString(),
      };

      const store = loadStore();
      store.comments.push(item);
      await saveStore(store);
      res.status(201).json({ comment: item });
    }
  );

  app.patch(
    '/api/comments/:id',
    requireAuth,
    requireRoles('Management', 'Admin'),
    async (req: AuthedRequest, res: Response) => {
      const id = req.params.id;
      const status = req.body?.status as CommentStatus | undefined;
      if (!status || !VALID_COMMENT_STATUSES.includes(status)) {
        res.status(400).json({ error: 'Invalid status.' });
        return;
      }
      const store = loadStore();
      const target = store.comments.find((c) => c.id === id);
      if (!target) {
        res.status(404).json({ error: 'Comment not found.' });
        return;
      }
      if (req.user!.role !== 'Admin' && target.author !== req.user!.username) {
        res.status(403).json({ error: 'Only the author or Admin can update this comment.' });
        return;
      }
      target.status = status;
      target.updatedAt = new Date().toISOString();
      await saveStore(store);
      res.json({ comment: target });
    }
  );

  app.delete(
    '/api/comments/:id',
    requireAuth,
    requireRoles('Management', 'Admin'),
    async (req: AuthedRequest, res: Response) => {
      const id = req.params.id;
      const store = loadStore();
      const idx = store.comments.findIndex((c) => c.id === id);
      if (idx === -1) {
        res.status(404).json({ error: 'Comment not found.' });
        return;
      }
      if (req.user!.role !== 'Admin' && store.comments[idx].author !== req.user!.username) {
        res.status(403).json({ error: 'Only the author or Admin can delete this comment.' });
        return;
      }
      store.comments.splice(idx, 1);
      await saveStore(store);
      res.json({ ok: true });
    }
  );

  /* ---------- SHARED WORKSPACE SYNC (REPORTS & DIRECTORY) ---------- */
  app.get('/api/workspace', requireAuth, (req: AuthedRequest, res: Response) => {
    const store = loadStore();
    const role = req.user!.role;

    if (role === 'Fuel Officer') {
      res.json({
        vehicleAreas: null,
        consolidatedDays: {},
        gpsSummary: null,
        voucherSummary: null,
      });
      return;
    }

    res.json({
      vehicleAreas:
        role === 'Admin' || role === 'Controller 2' ? store.workspace.vehicleAreas : null,
      consolidatedDays:
        role === 'Admin' || role === 'Controller 1' || role === 'Management'
          ? store.workspace.consolidatedDays
          : {},
      gpsSummary:
        role === 'Admin' || role === 'Controller 1' || role === 'Management'
          ? store.workspace.gpsSummary
          : null,
      voucherSummary:
        role === 'Admin' || role === 'Controller 2' || role === 'Management'
          ? store.workspace.voucherSummary
          : null,
    });
  });

  app.put('/api/workspace', requireAuth, async (req: AuthedRequest, res: Response) => {
    const role = req.user!.role;
    if (role === 'Fuel Officer' || role === 'Management') {
      res.status(403).json({ error: 'Read-only role cannot modify operational workspace state.' });
      return;
    }

    const store = loadStore();
    if (role === 'Admin' && req.body?.vehicleAreas && typeof req.body.vehicleAreas === 'object') {
      store.workspace.vehicleAreas = req.body.vehicleAreas;
    }
    if (
      (role === 'Admin' || role === 'Controller 1') &&
      req.body?.consolidatedDays &&
      typeof req.body.consolidatedDays === 'object'
    ) {
      store.workspace.consolidatedDays = req.body.consolidatedDays;
    }
    if (
      (role === 'Admin' || role === 'Controller 1') &&
      req.body?.gpsSummary !== undefined
    ) {
      store.workspace.gpsSummary = req.body.gpsSummary;
    }
    if (
      (role === 'Admin' || role === 'Controller 2') &&
      req.body?.voucherSummary !== undefined
    ) {
      store.workspace.voucherSummary = req.body.voucherSummary;
    }

    await saveStore(store);
    res.json({ ok: true });
  });

  /* ---------- VITE DEV / STATIC PRODUCTION ---------- */
  if (process.env.NODE_ENV !== 'production') {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distDir = path.resolve(__dirname, 'dist');
    app.use(express.static(distDir));
    app.get('*', (req: Request, res: Response) => {
      if (req.path.startsWith('/api/')) {
        res.status(404).json({ error: 'API route not found.' });
        return;
      }
      res.sendFile(path.resolve(distDir, 'index.html'));
    });
  }

  const port = Number(process.env.PORT) || 3000;
  app.listen(port, '0.0.0.0', () => {
    console.log(`Cleantech Fuel Management Server listening on http://0.0.0.0:${port}`);
  });
}

void startServer();
