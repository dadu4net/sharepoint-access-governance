'use client';

import React, { useState } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { NoppaIcon } from './NoppaIcon';
import { ThemeSelector } from './ThemeSelector';
import { TenantSettingsModal } from '../TenantSettingsModal';

export function Header() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  const handleLogin = async () => {
    setIsLoggingIn(true);
    try {
      await instance.loginPopup({
        scopes: ['Sites.Read.All', 'Directory.Read.All', 'User.Read']
      });
    } catch (e) {
      console.error('Login fout', e);
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleLogout = () => {
    instance.logoutPopup();
  };

  return (
    <>
      <header className="sticky top-0 z-40 bg-card/80 backdrop-blur-md border-b border-card shadow-apple-sm transition-colors">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          
          {/* Logo / Title Area */}
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-cyan to-royal flex items-center justify-center shadow-lg shadow-cyan/25 flex-shrink-0 text-white">
              <NoppaIcon name="sharepoint" size={24} />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="font-extrabold text-base tracking-tight bg-gradient-to-r from-cyan to-royal bg-clip-text text-transparent">
                  NOPPA
                </span>
                <span className="text-xs px-2 py-0.5 rounded-full bg-cyan/10 border border-cyan/30 text-cyan font-semibold">
                  Governance Portal
                </span>
              </div>
              <h1 className="text-sm font-semibold text-heading leading-none">
                SharePoint & Teams Access Governance
              </h1>
            </div>
          </div>

          {/* Actions / Auth Area */}
          <div className="flex items-center gap-4">
            
            {isAuthenticated && accounts[0] ? (
              <div className="flex items-center gap-3 bg-hover/50 px-3 py-1.5 rounded-full border border-card">
                <div className="flex flex-col text-right">
                  <span className="text-sm font-bold text-heading leading-tight">{accounts[0].name}</span>
                  <span className="text-[10px] text-muted">{accounts[0].username}</span>
                </div>
                <div className="w-8 h-8 rounded-full bg-royal text-white flex items-center justify-center font-bold text-sm shadow-sm">
                  {accounts[0].name?.charAt(0)}
                </div>
                <button
                  onClick={handleLogout}
                  className="ml-2 p-1.5 text-muted hover:text-rose-500 transition-colors"
                  title="Uitloggen"
                >
                  <NoppaIcon name="close" size={16} />
                </button>
              </div>
            ) : (
              <button
                onClick={handleLogin}
                disabled={isLoggingIn}
                className="btn btn-primary py-2 px-4 text-sm"
              >
                {isLoggingIn ? 'Bezig...' : 'Inloggen M365'}
              </button>
            )}

            <div className="w-px h-6 bg-card" />

            <button
              onClick={() => setIsSettingsOpen(true)}
              className="p-2 text-muted hover:text-cyan transition-colors"
              title="Instellingen"
            >
              <NoppaIcon name="settings" size={20} />
            </button>

            <ThemeSelector />
          </div>
        </div>
      </header>

      {/* Settings Modal */}
      <TenantSettingsModal
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
      />
    </>
  );
}
