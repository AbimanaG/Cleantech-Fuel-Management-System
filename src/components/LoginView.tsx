import React, { useState } from 'react';
import { Lock, Eye, EyeOff, ShieldCheck } from 'lucide-react';
import type { AuthUser } from '../types/fuel';
import { authApi, setAuthToken } from '../utils/apiClient';

interface LoginViewProps {
  needsSetup: boolean;
  envAdminConfigured: boolean;
  onAuthenticated: (user: AuthUser) => void;
}

export const LoginView: React.FC<LoginViewProps> = ({
  needsSetup,
  envAdminConfigured,
  onAuthenticated,
}) => {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);

    if (needsSetup) {
      if (password !== confirmPassword) {
        setError('Passwords do not match.');
        return;
      }
      setLoading(true);
      try {
        const res = await authApi.claimAdmin(username, password);
        setAuthToken(res.token);
        onAuthenticated(res.user);
      } catch (err: any) {
        setError(err?.message || 'Could not create Admin account.');
      } finally {
        setLoading(false);
      }
      return;
    }

    setLoading(true);
    try {
      const res = await authApi.login(username, password);
      setAuthToken(res.token);
      onAuthenticated(res.user);
    } catch (err: any) {
      setError(err?.message || 'Sign in failed.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#F8F9FA] flex items-center justify-center p-6">
      <div className="w-full max-w-md bg-white border border-slate-200 rounded-md p-8 space-y-6">
        <div className="space-y-2 border-b border-slate-100 pb-5">
          <div className="flex items-center gap-2 text-[#1F5143]">
            {needsSetup ? <ShieldCheck className="w-5 h-5" /> : <Lock className="w-5 h-5" />}
            <span className="text-xs font-semibold tracking-wide text-[#1F5143]">
              Cleantech Fuel Management
            </span>
          </div>
          <h1 className="text-xl font-bold text-slate-900">
            {needsSetup ? 'Initial Administrator Setup' : 'Sign In to Workspace'}
          </h1>
          <p className="text-sm text-slate-600 leading-relaxed">
            {needsSetup
              ? 'Create the primary Administrator account to unlock User Management and configure access for Fuel Officer, Controllers, and Management.'
              : 'Enter your assigned User Name and Password to access your role workspace.'}
          </p>
        </div>

        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label htmlFor="login-username" className="block text-xs font-medium text-slate-700 mb-1.5">
              User Name
            </label>
            <input
              id="login-username"
              type="text"
              autoComplete="username"
              required
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              placeholder={needsSetup ? 'e.g. admin' : 'Enter your User Name'}
              className="w-full px-3 py-2 text-sm border border-slate-300 rounded-md bg-white text-slate-900 focus:outline-none focus:ring-2 focus:ring-[#1F5143]/25 focus:border-[#1F5143]"
            />
          </div>

          <div>
            <label htmlFor="login-password" className="block text-xs font-medium text-slate-700 mb-1.5">
              Password
            </label>
            <div className="relative">
              <input
                id="login-password"
                type={showPassword ? 'text' : 'password'}
                autoComplete={needsSetup ? 'new-password' : 'current-password'}
                required
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder={needsSetup ? 'Minimum 6 characters' : 'Enter your Password'}
                className="w-full px-3 py-2 pr-10 text-sm border border-slate-300 rounded-md bg-white text-slate-900 focus:outline-none focus:ring-2 focus:ring-[#1F5143]/25 focus:border-[#1F5143]"
              />
              <button
                type="button"
                onClick={() => setShowPassword((v) => !v)}
                className="absolute inset-y-0 right-0 px-3 text-slate-400 hover:text-slate-700"
                aria-label={showPassword ? 'Hide password' : 'Show password'}
              >
                {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
              </button>
            </div>
          </div>

          {needsSetup && (
            <div>
              <label
                htmlFor="login-confirm-password"
                className="block text-xs font-medium text-slate-700 mb-1.5"
              >
                Confirm Password
              </label>
              <input
                id="login-confirm-password"
                type={showPassword ? 'text' : 'password'}
                autoComplete="new-password"
                required
                value={confirmPassword}
                onChange={(e) => setConfirmPassword(e.target.value)}
                placeholder="Re-enter password"
                className="w-full px-3 py-2 text-sm border border-slate-300 rounded-md bg-white text-slate-900 focus:outline-none focus:ring-2 focus:ring-[#1F5143]/25 focus:border-[#1F5143]"
              />
            </div>
          )}

          {error && (
            <div className="p-3 rounded-md border border-red-200 bg-red-50 text-xs text-red-800">
              {error}
            </div>
          )}

          <button
            type="submit"
            disabled={loading}
            className="w-full py-2.5 px-4 bg-[#1F5143] hover:bg-[#173D32] text-white text-sm font-medium rounded-md transition-colors disabled:opacity-60"
          >
            {loading
              ? 'Verifying...'
              : needsSetup
              ? 'Create Admin Account & Sign In'
              : 'Sign In'}
          </button>
        </form>

        <div className="pt-4 border-t border-slate-100 space-y-1.5 text-xs text-slate-500">
          <div className="font-medium text-slate-700">Configured Access Types</div>
          <div>Fuel Officer · Controller 1 · Controller 2 · Admin · Management</div>
          {envAdminConfigured && (
            <div className="text-slate-500">
              Environment admin passcode active for user <span className="font-data text-slate-700">admin</span>.
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
