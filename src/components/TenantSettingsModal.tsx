'use client';

import React, { useState, useEffect } from 'react';
import { NoppaIcon } from './common/NoppaIcon';

interface TenantSettingsModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSaved?: () => void;
}

interface TenantConfigStatus {
  configured: boolean;
  tenantId?: string;
  clientId?: string;
}

export function TenantSettingsModal({ isOpen, onClose, onSaved }: TenantSettingsModalProps) {
  const [tenantId, setTenantId] = useState('');
  const [clientId, setClientId] = useState('');
  const [status, setStatus] = useState<TenantConfigStatus | null>(null);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState<{ type: 'success' | 'error' | 'info'; text: string } | null>(null);

  useEffect(() => {
    if (isOpen) {
      loadStatus();
    }
  }, [isOpen]);

  const loadStatus = () => {
    try {
      setMessage(null);
      const saved = localStorage.getItem('noppa_graph_config');
      if (saved) {
        const parsed = JSON.parse(saved);
        setStatus({
          configured: !!parsed.clientId,
          tenantId: parsed.tenantId,
          clientId: parsed.clientId,
        });
        if (parsed.tenantId) setTenantId(parsed.tenantId);
        if (parsed.clientId) setClientId(parsed.clientId);
      } else {
        setStatus({ configured: false });
      }
    } catch (e) {
      setMessage({ type: 'error', text: 'Kon huidige configuratie niet ophalen uit de browser.' });
    }
  };

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    if (!tenantId.trim() || !clientId.trim()) {
      setMessage({ type: 'error', text: 'Vul ten minste Tenant ID en Client ID in.' });
      return;
    }

    try {
      setLoading(true);
      setMessage(null);
      
      const config = {
        tenantId: tenantId.trim(),
        clientId: clientId.trim()
      };
      
      localStorage.setItem('noppa_graph_config', JSON.stringify(config));
      
      setMessage({ type: 'success', text: 'Entra ID SPA configuratie lokaal opgeslagen! De applicatie herlaadt...' });
      
      setTimeout(() => {
        if (onSaved) onSaved();
        window.location.reload();
      }, 1500);

    } catch (err: any) {
      setMessage({ type: 'error', text: `Fout bij opslaan: ${err.message}` });
      setLoading(false);
    }
  };

  const handleClear = () => {
    if (!confirm('Weet u zeker dat u de lokaal opgeslagen tenant credentials wilt wissen?')) {
      return;
    }
    try {
      localStorage.removeItem('noppa_graph_config');
      setTenantId('');
      setClientId('');
      setStatus({ configured: false });
      setMessage({ type: 'info', text: 'Credentials gewist. Applicatie herlaadt...' });
      
      setTimeout(() => {
        if (onSaved) onSaved();
        window.location.reload();
      }, 1000);
    } catch (err: any) {
      setMessage({ type: 'error', text: `Fout bij wissen: ${err.message}` });
    }
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm animate-fade-in">
      <div className="bg-card border border-card rounded-2xl w-full max-w-2xl max-h-[90vh] flex flex-col shadow-2xl overflow-hidden animate-scale-up">
        {/* Header */}
        <div className="p-6 border-b border-card flex items-center justify-between bg-card-header">
          <div className="flex items-center gap-3">
            <div className="p-2.5 rounded-xl bg-gradient-to-br from-cyan/20 to-royal/20 border border-cyan/30 text-cyan">
              <NoppaIcon name="security-shield" size={24} />
            </div>
            <div>
              <h2 className="text-xl font-bold text-heading">Microsoft Entra ID SPA Setup</h2>
              <p className="text-xs text-muted">Client-Side configuratie voor SharePoint Governance</p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-2 text-muted hover:text-heading hover:bg-hover rounded-lg transition-colors"
          >
            <NoppaIcon name="close" size={20} />
          </button>
        </div>

        {/* Content */}
        <div className="p-6 overflow-y-auto space-y-6 flex-1">
          {/* Status Badge */}
          <div className="flex items-center justify-between p-3.5 rounded-xl border bg-hover/40">
            <div className="flex items-center gap-2.5">
              <span
                className={`w-3 h-3 rounded-full ${
                  status?.configured
                    ? 'bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.5)]'
                    : 'bg-amber-500 shadow-[0_0_8px_rgba(245,158,11,0.5)]'
                }`}
              />
              <span className="text-sm font-medium text-heading">
                Status: {status?.configured ? 'Klaar voor Single Page App (SPA) Login' : 'Niet geconfigureerd (Geen tenant actief)'}
              </span>
            </div>
          </div>

          {/* Feedback Message */}
          {message && (
            <div
              className={`p-4 rounded-xl border text-sm flex items-start gap-3 ${
                message.type === 'success'
                  ? 'bg-emerald-500/10 border-emerald-500/20 text-emerald-300'
                  : message.type === 'error'
                  ? 'bg-rose-500/10 border-rose-500/20 text-rose-300'
                  : 'bg-cyan/10 border-cyan/20 text-cyan'
              }`}
            >
              <div className="mt-0.5 flex-shrink-0">
                <NoppaIcon
                  name={message.type === 'success' ? 'check' : message.type === 'error' ? 'alert-triangle' : 'info'}
                  size={18}
                />
              </div>
              <div className="flex-1 font-medium">{message.text}</div>
            </div>
          )}

          {/* Form */}
          <form onSubmit={handleSave} className="space-y-4">
            <div>
              <label className="block text-xs font-semibold text-muted uppercase tracking-wider mb-1.5">
                Directory (Tenant) ID *
              </label>
              <input
                type="text"
                value={tenantId}
                onChange={(e) => setTenantId(e.target.value)}
                placeholder="bv. 72f988bf-86f1-41af-91ab-2d7cd011db47"
                className="w-full px-3.5 py-2.5 rounded-xl bg-primary border border-input text-heading placeholder-muted/60 text-sm focus:outline-none focus:border-cyan transition-colors"
                required
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-muted uppercase tracking-wider mb-1.5">
                Application (Client) ID *
              </label>
              <input
                type="text"
                value={clientId}
                onChange={(e) => setClientId(e.target.value)}
                placeholder="bv. 11111111-2222-3333-4444-555555555555"
                className="w-full px-3.5 py-2.5 rounded-xl bg-primary border border-input text-heading placeholder-muted/60 text-sm focus:outline-none focus:border-cyan transition-colors"
                required
              />
              <p className="text-[11px] text-muted mt-1">ID van uw Entra App Registration (als "Single-page application" geconfigureerd).</p>
            </div>

            {/* Quick Entra ID Permissions Summary Card */}
            <div className="p-4 rounded-xl border border-card bg-primary/40 space-y-2.5 text-xs text-muted">
              <div className="flex items-center gap-2 text-heading font-semibold">
                <NoppaIcon name="security-lock" size={16} />
                <span>Vereiste Delegated Permissions (SPA Flow)</span>
              </div>
              <p className="leading-relaxed">
                Aangezien deze app nu volledig lokaal draait in de browser, logt de gebruiker zelf in (Delegated). 
                Zorg dat de App Registration in Entra ID is ingesteld als <strong>Single-page application (SPA)</strong> met de volgende <strong>Delegated</strong> permissies:
              </p>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 pt-1 font-mono text-[11px]">
                <div className="p-2 rounded-lg bg-card border border-card flex items-center gap-2">
                  <span className="text-emerald-400 font-bold">✓</span>
                  <span>Sites.Read.All</span>
                </div>
                <div className="p-2 rounded-lg bg-card border border-card flex items-center gap-2">
                  <span className="text-emerald-400 font-bold">✓</span>
                  <span>User.Read</span>
                </div>
                <div className="p-2 rounded-lg bg-card border border-card flex items-center gap-2">
                  <span className="text-emerald-400 font-bold">✓</span>
                  <span>Directory.Read.All</span>
                </div>
              </div>
            </div>

            {/* Buttons */}
            <div className="flex flex-wrap items-center justify-between gap-3 pt-3 border-t border-card">
              <div>
                {status?.configured && (
                  <button
                    type="button"
                    onClick={handleClear}
                    disabled={loading}
                    className="px-3 py-2 text-xs font-medium text-rose-400 hover:text-rose-300 hover:bg-rose-500/10 rounded-xl transition-colors"
                  >
                    Wissen Configuratie
                  </button>
                )}
              </div>
              <div className="flex items-center gap-3">
                <button
                  type="submit"
                  disabled={loading}
                  className="px-5 py-2 text-xs font-semibold rounded-xl bg-gradient-to-r from-cyan to-royal text-white shadow-lg shadow-cyan/20 hover:opacity-95 transition-opacity disabled:opacity-50 flex items-center gap-2"
                >
                  <NoppaIcon name="check" size={14} />
                  <span>{loading ? 'Opslaan...' : 'Opslaan & Herstarten'}</span>
                </button>
              </div>
            </div>
          </form>
        </div>
      </div>
    </div>
  );
}
