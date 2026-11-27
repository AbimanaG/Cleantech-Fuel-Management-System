import React, { useState } from 'react';
import { UserPlus, Trash2, KeyRound, Check } from 'lucide-react';
import type { ManagedUser, UserRole } from '../types/fuel';
import { usersApi } from '../utils/apiClient';

interface UserManagementViewProps {
  users: ManagedUser[];
  currentUserId: string;
  onUsersChange: (next: ManagedUser[]) => void;
}

const ROLES: UserRole[] = [
  'Fuel Officer',
  'Controller 1',
  'Controller 2',
  'Admin',
  'Management',
];

const ROLE_DESCRIPTIONS: Record<UserRole, string> = {
  'Fuel Officer': 'Only downloads all completed files',
  'Controller 1': 'GPS Reconciliation & Multi-Day Consolidation',
  'Controller 2': 'Fuel Voucher Analysis',
  Admin: 'Full access across all modules & user accounts',
  Management: 'Overview reports & management comments',
};

export const UserManagementView: React.FC<UserManagementViewProps> = ({
  users,
  currentUserId,
  onUsersChange,
}) => {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [role, setRole] = useState<UserRole>('Controller 1');
  const [creating, setCreating] = useState(false);
  const [resetId, setResetId] = useState<string | null>(null);
  const [resetPassword, setResetPassword] = useState('');
  const [notice, setNotice] = useState<{ msg: string; error?: boolean } | null>(null);

  const handleCreate = async (e: React.FormEvent) => {
    e.preventDefault();
    setNotice(null);
    setCreating(true);
    try {
      const res = await usersApi.create(username, password, role);
      onUsersChange([...users, res.user]);
      setUsername('');
      setPassword('');
      setNotice({ msg: `Created user "${res.user.username}" (${res.user.role}).` });
    } catch (err: any) {
      setNotice({ msg: err?.message || 'Could not create user.', error: true });
    } finally {
      setCreating(false);
    }
  };

  const handleRoleChange = async (id: string, nextRole: UserRole) => {
    setNotice(null);
    try {
      const res = await usersApi.update(id, { role: nextRole });
      onUsersChange(users.map((u) => (u.id === id ? res.user : u)));
      setNotice({ msg: `Updated role for ${res.user.username} to ${res.user.role}.` });
    } catch (err: any) {
      setNotice({ msg: err?.message || 'Could not update role.', error: true });
    }
  };

  const handleToggleActive = async (user: ManagedUser) => {
    setNotice(null);
    try {
      const res = await usersApi.update(user.id, { active: !user.active });
      onUsersChange(users.map((u) => (u.id === user.id ? res.user : u)));
      setNotice({
        msg: `${res.user.username} is now ${res.user.active ? 'Active' : 'Disabled'}.`,
      });
    } catch (err: any) {
      setNotice({ msg: err?.message || 'Could not update user status.', error: true });
    }
  };

  const handleSaveResetPassword = async (id: string) => {
    setNotice(null);
    try {
      const res = await usersApi.update(id, { password: resetPassword });
      setResetId(null);
      setResetPassword('');
      setNotice({ msg: `Password updated for ${res.user.username}.` });
    } catch (err: any) {
      setNotice({ msg: err?.message || 'Could not reset password.', error: true });
    }
  };

  const handleDelete = async (user: ManagedUser) => {
    setNotice(null);
    try {
      await usersApi.remove(user.id);
      onUsersChange(users.filter((u) => u.id !== user.id));
      setNotice({ msg: `Removed user ${user.username}.` });
    } catch (err: any) {
      setNotice({ msg: err?.message || 'Could not delete user.', error: true });
    }
  };

  return (
    <div className="space-y-6">
      <div className="border-b border-slate-200 pb-5">
        <h1 className="text-xl font-bold text-slate-900">User &amp; Role Management</h1>
        <p className="text-sm text-slate-600 mt-1">
          Create and manage login credentials for Fuel Officer, Controller 1, Controller 2, Admin, and Management.
        </p>
      </div>

      {/* Add User Card */}
      <form
        onSubmit={handleCreate}
        className="bg-white border border-slate-200 rounded-md p-6 space-y-4"
      >
        <div className="text-sm font-semibold text-slate-900">Add User Account</div>
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4 items-end">
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">User Name</label>
            <input
              type="text"
              required
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              placeholder="e.g. controller1"
              className="w-full px-3 py-2 text-sm border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-[#1F5143]/25"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">Password</label>
            <input
              type="password"
              required
              minLength={6}
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="Min. 6 characters"
              className="w-full px-3 py-2 text-sm border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-[#1F5143]/25"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">User Type / Role</label>
            <select
              value={role}
              onChange={(e) => setRole(e.target.value as UserRole)}
              className="w-full px-3 py-2 text-sm border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-[#1F5143]/25"
            >
              {ROLES.map((r) => (
                <option key={r} value={r}>
                  {r}
                </option>
              ))}
            </select>
          </div>
          <div>
            <button
              type="submit"
              disabled={creating}
              className="w-full inline-flex items-center justify-center gap-2 px-4 py-2 bg-[#1F5143] hover:bg-[#173D32] text-white text-sm font-medium rounded-md transition-colors disabled:opacity-60"
            >
              <UserPlus className="w-4 h-4" />
              <span>{creating ? 'Creating...' : 'Create User'}</span>
            </button>
          </div>
        </div>
        <div className="text-xs text-slate-500">
          Selected permission: <span className="font-medium text-slate-700">{role}</span> —{' '}
          {ROLE_DESCRIPTIONS[role]}
        </div>
      </form>

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

      {/* Users Table */}
      <div className="bg-white border border-slate-200 rounded-md overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-sm">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-600">
                <th className="px-4 py-3">User Name</th>
                <th className="px-4 py-3">Access Type</th>
                <th className="px-4 py-3">Permissions</th>
                <th className="px-4 py-3">Status</th>
                <th className="px-4 py-3 text-right">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {users.length === 0 ? (
                <tr>
                  <td colSpan={5} className="px-4 py-8 text-center text-xs text-slate-500">
                    No stored accounts created yet. Use the form above to add Fuel Officer, Controller 1, Controller 2, or Management accounts.
                  </td>
                </tr>
              ) : (
                users.map((u) => (
                  <tr key={u.id} className="hover:bg-slate-50">
                    <td className="px-4 py-3 font-medium text-slate-900">
                      {u.username}
                      {u.id === currentUserId && (
                        <span className="ml-2 text-xs text-slate-400">(You)</span>
                      )}
                    </td>
                    <td className="px-4 py-3">
                      <select
                        value={u.role}
                        onChange={(e) => void handleRoleChange(u.id, e.target.value as UserRole)}
                        className="px-2.5 py-1 text-xs border border-slate-300 rounded bg-white text-slate-800"
                      >
                        {ROLES.map((r) => (
                          <option key={r} value={r}>
                            {r}
                          </option>
                        ))}
                      </select>
                    </td>
                    <td className="px-4 py-3 text-xs text-slate-600">
                      {ROLE_DESCRIPTIONS[u.role]}
                    </td>
                    <td className="px-4 py-3 text-xs">
                      <button
                        type="button"
                        onClick={() => void handleToggleActive(u)}
                        className={`font-semibold ${
                          u.active ? 'text-emerald-700' : 'text-red-700'
                        }`}
                      >
                        {u.active ? 'Active' : 'Disabled'}
                      </button>
                    </td>
                    <td className="px-4 py-3 text-right whitespace-nowrap">
                      {resetId === u.id ? (
                        <div className="inline-flex items-center gap-1.5">
                          <input
                            type="password"
                            minLength={6}
                            value={resetPassword}
                            onChange={(e) => setResetPassword(e.target.value)}
                            placeholder="New password"
                            className="px-2 py-1 text-xs border border-slate-300 rounded w-32"
                          />
                          <button
                            type="button"
                            onClick={() => void handleSaveResetPassword(u.id)}
                            className="inline-flex items-center gap-1 px-2 py-1 bg-[#1F5143] text-white text-xs rounded"
                          >
                            <Check className="w-3 h-3" /> Save
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              setResetId(null);
                              setResetPassword('');
                            }}
                            className="px-2 py-1 border border-slate-300 text-xs rounded"
                          >
                            Cancel
                          </button>
                        </div>
                      ) : (
                        <div className="inline-flex items-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => {
                              setResetId(u.id);
                              setResetPassword('');
                            }}
                            className="inline-flex items-center gap-1 px-2.5 py-1 border border-slate-300 rounded text-xs text-slate-700 hover:bg-slate-100"
                          >
                            <KeyRound className="w-3 h-3" /> Reset Password
                          </button>
                          {u.id !== currentUserId && (
                            <button
                              type="button"
                              onClick={() => void handleDelete(u)}
                              className="p-1.5 border border-red-200 rounded text-red-700 hover:bg-red-50"
                              aria-label={`Delete ${u.username}`}
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          )}
                        </div>
                      )}
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
