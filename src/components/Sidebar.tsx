import React from 'react';
import {
  Radar,
  FileSpreadsheet,
  ReceiptText,
  MapPin,
  FolderDown,
  ClipboardCheck,
  Users,
  X,
} from 'lucide-react';
import type { ModuleTab, UserRole } from '../types/fuel';

interface SidebarProps {
  activeTab: ModuleTab;
  onSelectTab: (tab: ModuleTab) => void;
  userRole: UserRole;
  counts: {
    cardCount: number;
    consolidatedDays: number;
    voucherPending: number;
    voucherApplied: number;
    completedFiles: number;
    commentsCount: number;
    directoryVehicles: number;
  };
  mobileOpen: boolean;
  onCloseMobile: () => void;
}

export const ROLE_ALLOWED_TABS: Record<UserRole, ModuleTab[]> = {
  'Fuel Officer': ['files'],
  'Controller 1': ['gps', 'consolidation'],
  'Controller 2': ['voucher'],
  Management: ['management'],
  Admin: ['gps', 'consolidation', 'voucher', 'files', 'management', 'directory', 'users'],
};

export const Sidebar: React.FC<SidebarProps> = ({
  activeTab,
  onSelectTab,
  userRole,
  counts,
  mobileOpen,
  onCloseMobile,
}) => {
  const allItems: {
    id: ModuleTab;
    label: string;
    icon: React.ReactNode;
    meta?: string;
  }[] = [
    {
      id: 'gps',
      label: 'GPS Reconciliation',
      icon: <Radar className="w-4 h-4 shrink-0" />,
      meta: counts.cardCount > 0 ? `${counts.cardCount}` : undefined,
    },
    {
      id: 'consolidation',
      label: 'Multi-Day Excel',
      icon: <FileSpreadsheet className="w-4 h-4 shrink-0" />,
      meta: counts.consolidatedDays > 0 ? `${counts.consolidatedDays}d` : undefined,
    },
    {
      id: 'voucher',
      label: 'Voucher Analysis',
      icon: <ReceiptText className="w-4 h-4 shrink-0" />,
      meta:
        counts.voucherPending + counts.voucherApplied > 0
          ? `${counts.voucherPending}/${counts.voucherApplied}`
          : undefined,
    },
    {
      id: 'files',
      label: 'Completed Files',
      icon: <FolderDown className="w-4 h-4 shrink-0" />,
      meta: counts.completedFiles > 0 ? `${counts.completedFiles}` : undefined,
    },
    {
      id: 'management',
      label: 'Reports & Comments',
      icon: <ClipboardCheck className="w-4 h-4 shrink-0" />,
      meta: counts.commentsCount > 0 ? `${counts.commentsCount}` : undefined,
    },
    {
      id: 'directory',
      label: 'Vehicle → Area List',
      icon: <MapPin className="w-4 h-4 shrink-0" />,
      meta: `${counts.directoryVehicles}`,
    },
    {
      id: 'users',
      label: 'User Management',
      icon: <Users className="w-4 h-4 shrink-0" />,
    },
  ];

  const allowed = ROLE_ALLOWED_TABS[userRole];
  const navItems = allItems.filter((item) => allowed.includes(item.id));

  const handleNav = (id: ModuleTab) => {
    onSelectTab(id);
    onCloseMobile();
  };

  return (
    <>
      {mobileOpen && (
        <div
          className="fixed inset-0 bg-slate-900/40 z-40 lg:hidden"
          onClick={onCloseMobile}
        />
      )}
      <aside
        className={`fixed inset-y-0 left-0 z-50 w-64 bg-white border-r border-slate-200 flex flex-col transition-transform duration-150 lg:translate-x-0 lg:static lg:shrink-0 ${
          mobileOpen ? 'translate-x-0' : '-translate-x-full'
        }`}
      >
        <div className="h-14 px-5 border-b border-slate-200 flex items-center justify-between">
          <span className="font-bold text-base tracking-tight text-slate-900 whitespace-nowrap">
            Cleantech Fuel
          </span>
          <button
            type="button"
            onClick={onCloseMobile}
            className="lg:hidden p-1.5 text-slate-500 hover:text-slate-900 rounded-md"
            aria-label="Close navigation"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <nav className="flex-1 px-3 py-4 space-y-1 overflow-y-auto" aria-label="Primary navigation">
          {navItems.map((item) => {
            const active = activeTab === item.id;
            return (
              <button
                key={item.id}
                type="button"
                onClick={() => handleNav(item.id)}
                className={`w-full flex items-center justify-between gap-2 px-3 py-2.5 rounded-md text-sm font-medium transition-colors whitespace-nowrap ${
                  active
                    ? 'bg-[#1F5143] text-white'
                    : 'text-slate-700 hover:bg-slate-100 hover:text-slate-900'
                }`}
              >
                <span className="flex items-center gap-2.5 truncate">
                  {item.icon}
                  <span className="truncate">{item.label}</span>
                </span>
                {item.meta && (
                  <span
                    className={`font-data text-xs shrink-0 ${
                      active ? 'text-emerald-100' : 'text-slate-400'
                    }`}
                  >
                    {item.meta}
                  </span>
                )}
              </button>
            );
          })}
        </nav>

        <div className="px-5 py-4 border-t border-slate-200 text-xs text-slate-500 space-y-1">
          <div className="font-medium text-slate-700">{userRole}</div>
          <div>Sedawatta Station · ±1.5 L Tol.</div>
        </div>
      </aside>
    </>
  );
};
