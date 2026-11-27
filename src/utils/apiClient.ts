import type {
  AuthUser,
  CommentStatus,
  CompletedFileModule,
  CompletedFileRecord,
  GpsSummarySnapshot,
  ManagedUser,
  ManagementComment,
  UserRole,
  VoucherSummarySnapshot,
} from '../types/fuel';

// Keep auth token strictly in memory (never localStorage or sessionStorage)
let memoryToken: string | null = null;

export function setAuthToken(token: string | null) {
  memoryToken = token;
}

export function getAuthToken(): string | null {
  return memoryToken;
}

async function apiFetch<T>(url: string, options: RequestInit = {}): Promise<T> {
  const headers = new Headers(options.headers || {});
  if (memoryToken) {
    headers.set('Authorization', `Bearer ${memoryToken}`);
  }
  if (options.body && !(options.body instanceof FormData) && !headers.has('Content-Type')) {
    headers.set('Content-Type', 'application/json');
  }

  const res = await fetch(url, { ...options, headers });
  if (!res.ok) {
    let msg = `Request failed (${res.status})`;
    try {
      const err = await res.json();
      if (err?.error) msg = err.error;
    } catch {
      // ignore parse failure
    }
    throw new Error(msg);
  }
  return res.json() as Promise<T>;
}

export function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const result = typeof reader.result === 'string' ? reader.result : '';
      const idx = result.indexOf(',');
      resolve(idx >= 0 ? result.slice(idx + 1) : result);
    };
    reader.onerror = () => reject(new Error('Failed to encode file data.'));
    reader.readAsDataURL(blob);
  });
}

export const authApi = {
  bootstrapStatus: () =>
    apiFetch<{ needsSetup: boolean; envAdminConfigured: boolean }>('/api/auth/bootstrap-status'),
  claimAdmin: (username: string, password: string) =>
    apiFetch<{ token: string; user: AuthUser }>('/api/auth/claim-admin', {
      method: 'POST',
      body: JSON.stringify({ username, password }),
    }),
  login: (username: string, password: string) =>
    apiFetch<{ token: string; user: AuthUser }>('/api/auth/login', {
      method: 'POST',
      body: JSON.stringify({ username, password }),
    }),
  me: () => apiFetch<{ user: AuthUser }>('/api/auth/me'),
};

export const usersApi = {
  list: () => apiFetch<{ users: ManagedUser[] }>('/api/users'),
  create: (username: string, password: string, role: UserRole) =>
    apiFetch<{ user: ManagedUser }>('/api/users', {
      method: 'POST',
      body: JSON.stringify({ username, password, role }),
    }),
  update: (
    id: string,
    patch: { role?: UserRole; active?: boolean; password?: string }
  ) =>
    apiFetch<{ user: ManagedUser }>(`/api/users/${encodeURIComponent(id)}`, {
      method: 'PATCH',
      body: JSON.stringify(patch),
    }),
  remove: (id: string) =>
    apiFetch<{ ok: boolean }>(`/api/users/${encodeURIComponent(id)}`, {
      method: 'DELETE',
    }),
};

export const completedFilesApi = {
  list: () => apiFetch<{ files: CompletedFileRecord[] }>('/api/completed-files'),
  publish: (payload: {
    title: string;
    fileName: string;
    module: CompletedFileModule;
    format: 'xlsx' | 'pdf' | 'csv' | 'html';
    mimeType: string;
    contentBase64: string;
    summary: string;
    reportDate: string;
  }) =>
    apiFetch<{ file: CompletedFileRecord }>('/api/completed-files', {
      method: 'POST',
      body: JSON.stringify(payload),
    }),
  download: async (file: CompletedFileRecord): Promise<void> => {
    const headers = new Headers();
    if (memoryToken) headers.set('Authorization', `Bearer ${memoryToken}`);
    const res = await fetch(`/api/completed-files/${encodeURIComponent(file.id)}/download`, {
      headers,
    });
    if (!res.ok) {
      let msg = 'Download failed.';
      try {
        const err = await res.json();
        if (err?.error) msg = err.error;
      } catch {
        // ignore
      }
      throw new Error(msg);
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = file.fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  },
  remove: (id: string) =>
    apiFetch<{ ok: boolean }>(`/api/completed-files/${encodeURIComponent(id)}`, {
      method: 'DELETE',
    }),
};

export const commentsApi = {
  list: () => apiFetch<{ comments: ManagementComment[] }>('/api/comments'),
  create: (payload: {
    module: ManagementComment['module'];
    reportDate: string;
    status: CommentStatus;
    comment: string;
  }) =>
    apiFetch<{ comment: ManagementComment }>('/api/comments', {
      method: 'POST',
      body: JSON.stringify(payload),
    }),
  updateStatus: (id: string, status: CommentStatus) =>
    apiFetch<{ comment: ManagementComment }>(`/api/comments/${encodeURIComponent(id)}`, {
      method: 'PATCH',
      body: JSON.stringify({ status }),
    }),
  remove: (id: string) =>
    apiFetch<{ ok: boolean }>(`/api/comments/${encodeURIComponent(id)}`, {
      method: 'DELETE',
    }),
};

export const workspaceApi = {
  get: () =>
    apiFetch<{
      vehicleAreas: Record<string, string[]> | null;
      consolidatedDays: Record<string, any>;
      gpsSummary: GpsSummarySnapshot | null;
      voucherSummary: VoucherSummarySnapshot | null;
    }>('/api/workspace'),
  save: (patch: {
    vehicleAreas?: Record<string, string[]>;
    consolidatedDays?: Record<string, any>;
    gpsSummary?: GpsSummarySnapshot | null;
    voucherSummary?: VoucherSummarySnapshot | null;
  }) =>
    apiFetch<{ ok: boolean }>('/api/workspace', {
      method: 'PUT',
      body: JSON.stringify(patch),
    }),
};
