'use client';

import React, { useState, useEffect, useMemo, useCallback } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { GraphSite } from '@/lib/graph/client';
import { fetchSharePointSites } from '@/lib/graph/clientService';
import { SitePermissionsModal } from '@/components/SitePermissionsModal';
import { TenantExternalAuditModal } from '@/components/TenantExternalAuditModal';
import { TenantSettingsModal } from '@/components/TenantSettingsModal';
import { NoppaIcon } from '@/components/common/NoppaIcon';

export default function SharePointGovernanceDashboard() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const [sites, setSites] = useState<GraphSite[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isDemo, setIsDemo] = useState(!isAuthenticated);

  // Search & Filtering
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState<'all' | 'teams' | 'external'>('all');

  // Modals state
  const [selectedSite, setSelectedSite] = useState<GraphSite | null>(null);
  const [isTenantAuditOpen, setIsTenantAuditOpen] = useState(false);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);

  const loadSites = useCallback(async () => {
    if (!isAuthenticated || accounts.length === 0) {
      setIsDemo(true);
      setSites([]); // Or load mock data here
      return;
    }

    try {
      setIsDemo(false);
      setLoading(true);
      setError(null);
      const graphSites = await fetchSharePointSites(instance, accounts);
      setSites(graphSites);
    } catch (err: any) {
      setError(err.message || 'Onbekende fout opgetreden bij het ophalen van sites.');
    } finally {
      setLoading(false);
    }
  }, [isAuthenticated, accounts, instance]);

  useEffect(() => {
    loadSites();
  }, [loadSites]);

  // Filtered sites
  const filteredSites = useMemo(() => {
    return sites.filter((site) => {
      const displayName = site.displayName || site.title || site.name || '';
      const name = site.name || site.title || '';
      const webUrl = site.webUrl || site.url || '';

      const matchesSearch =
        displayName.toLowerCase().includes(searchTerm.toLowerCase()) ||
        name.toLowerCase().includes(searchTerm.toLowerCase()) ||
        webUrl.toLowerCase().includes(searchTerm.toLowerCase());

      if (!matchesSearch) return false;

      if (filterType === 'teams') {
        return Boolean(site.isTeam);
      }
      if (filterType === 'external') {
        return Boolean(site.externalSharingCapability && site.externalSharingCapability !== 'Disabled');
      }
      return true;
    });
  }, [sites, searchTerm, filterType]);

  // KPIs
  const totalSites = sites.length;
  const teamsSitesCount = sites.filter((s) => s.isTeam).length;
  const externalSharingSitesCount = sites.filter(
    (s) => s.externalSharingCapability && s.externalSharingCapability !== 'Disabled'
  ).length;

  return (
    <div className="space-y-8 animate-fade-in">
      {/* Top Banner if Demo Mode */}
      {isDemo && (
        <div className="p-4 rounded-2xl border border-amber-500/30 bg-amber-500/10 text-amber-200 flex flex-col sm:flex-row sm:items-center justify-between gap-4 shadow-apple-sm">
          <div className="flex items-start gap-3">
            <div className="mt-0.5 text-amber-400 flex-shrink-0">
              <NoppaIcon name="alert-triangle" size={20} />
            </div>
            <div>
              <p className="text-sm font-bold text-amber-300">
                Demo Modus Actief &mdash; Geen Microsoft Entra Tenant Gekoppeld
              </p>
              <p className="text-xs text-amber-300/80 mt-0.5">
                U bekijkt voorbeelddata. Koppel uw eigen Microsoft Entra ID tenant om live 4-niveau rechten en externen te auditen.
              </p>
            </div>
          </div>
          <button
            onClick={() => setIsSettingsOpen(true)}
            className="px-4 py-2 text-xs font-semibold rounded-xl bg-amber-500 text-slate-950 hover:bg-amber-400 transition-colors shadow-md flex items-center justify-center gap-2 whitespace-nowrap self-start sm:self-auto"
          >
            <NoppaIcon name="security-lock" size={14} />
            <span>Tenant Instellen</span>
          </button>
        </div>
      )}

      {/* Hero / Header Card with Quick Actions */}
      <div className="p-6 sm:p-8 rounded-3xl bg-card border border-card shadow-apple-md relative overflow-hidden">
        <div className="absolute top-0 right-0 w-96 h-96 bg-gradient-to-br from-cyan/10 to-royal/10 rounded-full blur-3xl pointer-events-none -mr-20 -mt-20" />
        <div className="relative z-10 flex flex-col md:flex-row md:items-center justify-between gap-6">
          <div className="max-w-2xl">
            <span className="text-xs font-bold tracking-wider uppercase text-cyan px-2.5 py-1 rounded-md bg-cyan/10 border border-cyan/20 inline-block mb-3">
              SharePoint &amp; Teams Governance Suite
            </span>
            <h2 className="text-2xl sm:text-3xl font-extrabold text-heading tracking-tight">
              Toegangsrechten &amp; Externe Deelbeheer
            </h2>
            <p className="text-sm text-muted mt-2 leading-relaxed">
              Auditeer autorisaties tot op 4 niveaus diep: <strong>Site &rarr; Bibliotheek &rarr; Lijst &rarr; Item</strong>. 
              Vind direct alle externe gasten, anonieme deellinks en gedeelde locaties binnen uw Microsoft 365 tenant.
            </p>
          </div>

          <div className="flex flex-wrap sm:flex-nowrap items-center gap-3">
            <button
              onClick={() => setIsTenantAuditOpen(true)}
              className="w-full sm:w-auto px-5 py-3 rounded-2xl bg-gradient-to-r from-cyan to-royal text-white font-semibold text-sm shadow-xl shadow-cyan/25 hover:opacity-95 transition-all flex items-center justify-center gap-2.5 group"
            >
              <NoppaIcon name="users" size={18} className="group-hover:scale-110 transition-transform" />
              <span>Externen Zoeken &amp; Vinden</span>
            </button>
            <button
              onClick={() => setIsSettingsOpen(true)}
              className="w-full sm:w-auto px-4 py-3 rounded-2xl border border-input bg-hover/50 hover:bg-hover text-heading font-medium text-sm transition-colors flex items-center justify-center gap-2"
            >
              <NoppaIcon name="settings" size={16} />
              <span>Tenant Verbinding</span>
            </button>
          </div>
        </div>
      </div>

      {/* KPI Stats Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Card 1: Totaal Sites */}
        <div className="p-5 rounded-2xl bg-card border border-card shadow-apple-sm flex items-center justify-between">
          <div>
            <p className="text-xs font-semibold text-muted uppercase tracking-wider">Totaal Sites</p>
            <p className="text-2xl font-bold text-heading mt-1">{loading ? '...' : totalSites}</p>
            <p className="text-xs text-muted mt-0.5">SharePoint &amp; Teamsites</p>
          </div>
          <div className="p-3 rounded-xl bg-cyan/10 border border-cyan/20 text-cyan">
            <NoppaIcon name="sharepoint" size={24} />
          </div>
        </div>

        {/* Card 2: Teams Sites */}
        <div className="p-5 rounded-2xl bg-card border border-card shadow-apple-sm flex items-center justify-between">
          <div>
            <p className="text-xs font-semibold text-muted uppercase tracking-wider">Microsoft Teams</p>
            <p className="text-2xl font-bold text-heading mt-1">{loading ? '...' : teamsSitesCount}</p>
            <p className="text-xs text-muted mt-0.5">Met M365 Unified Group</p>
          </div>
          <div className="p-3 rounded-xl bg-royal/10 border border-royal/20 text-royal">
            <NoppaIcon name="teams" size={24} />
          </div>
        </div>

        {/* Card 3: Externe Deling Actief */}
        <div className="p-5 rounded-2xl bg-card border border-card shadow-apple-sm flex items-center justify-between">
          <div>
            <p className="text-xs font-semibold text-muted uppercase tracking-wider">Externe Deling Actief</p>
            <p className="text-2xl font-bold text-heading mt-1">{loading ? '...' : externalSharingSitesCount}</p>
            <p className="text-xs text-muted mt-0.5">Gasten &amp; Links toegestaan</p>
          </div>
          <div className="p-3 rounded-xl bg-amber-500/10 border border-amber-500/20 text-amber-400">
            <NoppaIcon name="external-link" size={24} />
          </div>
        </div>

        {/* Card 4: Audit Status */}
        <div className="p-5 rounded-2xl bg-card border border-card shadow-apple-sm flex items-center justify-between">
          <div>
            <p className="text-xs font-semibold text-muted uppercase tracking-wider">Audit Diepte</p>
            <p className="text-2xl font-bold text-emerald-400 mt-1">4 Niveaus</p>
            <p className="text-xs text-muted mt-0.5">Site, Bib, Lijst, Item</p>
          </div>
          <div className="p-3 rounded-xl bg-emerald-500/10 border border-emerald-500/20 text-emerald-400">
            <NoppaIcon name="security-shield" size={24} />
          </div>
        </div>
      </div>

      {/* Search & Filter Toolbar */}
      <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-4">
        {/* Search input */}
        <div className="relative flex-1 max-w-md">
          <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-muted">
            <NoppaIcon name="search" size={16} />
          </div>
          <input
            type="text"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder="Zoek site op naam of webadres..."
            className="w-full pl-10 pr-4 py-2.5 rounded-xl bg-card border border-input text-heading placeholder-muted/60 text-sm focus:outline-none focus:border-cyan shadow-apple-sm transition-colors"
          />
        </div>

        {/* Filter pills */}
        <div className="flex items-center gap-1.5 p-1 rounded-xl bg-card border border-card shadow-apple-sm">
          <button
            onClick={() => setFilterType('all')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
              filterType === 'all'
                ? 'bg-gradient-to-r from-cyan to-royal text-white shadow-sm'
                : 'text-muted hover:text-heading hover:bg-hover'
            }`}
          >
            Alle Sites ({sites.length})
          </button>
          <button
            onClick={() => setFilterType('teams')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
              filterType === 'teams'
                ? 'bg-gradient-to-r from-cyan to-royal text-white shadow-sm'
                : 'text-muted hover:text-heading hover:bg-hover'
            }`}
          >
            Teams ({teamsSitesCount})
          </button>
          <button
            onClick={() => setFilterType('external')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
              filterType === 'external'
                ? 'bg-gradient-to-r from-cyan to-royal text-white shadow-sm'
                : 'text-muted hover:text-heading hover:bg-hover'
            }`}
          >
            Met Externe Deling ({externalSharingSitesCount})
          </button>
        </div>
      </div>

      {/* Sites Table Container */}
      <div className="bg-card border border-card rounded-2xl shadow-apple-md overflow-hidden">
        <div className="p-4 sm:p-5 border-b border-card flex items-center justify-between bg-card-header">
          <div className="flex items-center gap-2">
            <NoppaIcon name="folder" size={18} className="text-cyan" />
            <h3 className="font-bold text-heading text-sm sm:text-base">
              SharePoint &amp; Teams Sites ({filteredSites.length})
            </h3>
          </div>
          <button
            onClick={loadSites}
            disabled={loading}
            className="p-1.5 rounded-lg text-muted hover:text-heading hover:bg-hover transition-colors"
            title="Verversen"
          >
            <NoppaIcon name="refresh" size={16} className={loading ? 'animate-spin' : ''} />
          </button>
        </div>

        {loading ? (
          <div className="p-12 text-center text-muted text-sm space-y-3">
            <div className="inline-block animate-spin text-cyan">
              <NoppaIcon name="refresh" size={28} />
            </div>
            <p>SharePoint sites en configuraties ophalen...</p>
          </div>
        ) : error ? (
          <div className="p-8 text-center text-rose-400 space-y-3">
            <p className="font-semibold text-sm">Fout bij laden van SharePoint sites</p>
            <p className="text-xs text-rose-400/80">{error}</p>
            <button
              onClick={loadSites}
              className="px-4 py-2 text-xs font-semibold rounded-xl bg-card border border-rose-500/30 text-rose-300 hover:bg-rose-500/10 transition-colors"
            >
              Opnieuw proberen
            </button>
          </div>
        ) : filteredSites.length === 0 ? (
          <div className="p-12 text-center text-muted text-sm">
            Geen sites gevonden die voldoen aan de zoekcriteria.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left text-sm">
              <thead className="text-[11px] font-semibold text-muted uppercase tracking-wider bg-card-header/60 border-b border-card">
                <tr>
                  <th className="py-3 px-4 sm:px-6">Site Naam &amp; URL</th>
                  <th className="py-3 px-4">Type</th>
                  <th className="py-3 px-4">Externe Deelmogelijkheid</th>
                  <th className="py-3 px-4">Laatst Gewijzigd</th>
                  <th className="py-3 px-4 text-right">Actie</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-card">
                {filteredSites.map((site) => (
                  <tr
                    key={site.id}
                    className="hover:bg-hover/50 transition-colors group"
                  >
                    {/* Site Title & WebUrl */}
                    <td className="py-4 px-4 sm:px-6">
                      <div className="flex items-start gap-3">
                        <div
                          className={`p-2 rounded-xl border flex-shrink-0 mt-0.5 ${
                            site.isTeam
                              ? 'bg-royal/10 border-royal/20 text-royal'
                              : 'bg-cyan/10 border-cyan/20 text-cyan'
                          }`}
                        >
                          <NoppaIcon name={site.isTeam ? 'teams' : 'sharepoint'} size={18} />
                        </div>
                        <div className="min-w-0">
                          <p className="font-semibold text-heading truncate">{site.displayName || site.title || site.name}</p>
                          {(site.webUrl || site.url) && (
                            <a
                              href={site.webUrl || site.url}
                              target="_blank"
                              rel="noreferrer"
                              className="text-xs text-muted hover:text-cyan truncate block transition-colors mt-0.5 flex items-center gap-1"
                            >
                              <span className="truncate">{site.webUrl || site.url}</span>
                              <NoppaIcon name="external-link" size={11} />
                            </a>
                          )}
                        </div>
                      </div>
                    </td>

                    {/* Site Type */}
                    <td className="py-4 px-4 whitespace-nowrap">
                      {site.isTeam ? (
                        <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold bg-royal/10 text-royal border border-royal/20">
                          <NoppaIcon name="teams" size={12} />
                          <span>Teamsite (M365)</span>
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold bg-cyan/10 text-cyan border border-cyan/20">
                          <NoppaIcon name="sharepoint" size={12} />
                          <span>SharePoint Site</span>
                        </span>
                      )}
                    </td>

                    {/* External Sharing Status */}
                    <td className="py-4 px-4 whitespace-nowrap">
                      {site.externalSharingCapability === 'Disabled' ? (
                        <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-md text-xs font-medium bg-slate-500/10 text-muted border border-slate-500/20">
                          <span>Uitgeschakeld</span>
                        </span>
                      ) : site.externalSharingCapability?.includes('Anonymous') ? (
                        <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-md text-xs font-semibold bg-rose-500/10 text-rose-400 border border-rose-500/20">
                          <NoppaIcon name="alert-triangle" size={12} />
                          <span>Anonieme Links Toegestaan</span>
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-md text-xs font-semibold bg-amber-500/10 text-amber-400 border border-amber-500/20">
                          <NoppaIcon name="users" size={12} />
                          <span>Gasten Toegestaan</span>
                        </span>
                      )}
                    </td>

                    {/* Last Modified */}
                    <td className="py-4 px-4 whitespace-nowrap text-xs text-muted">
                      {site.lastModifiedDateTime
                        ? new Date(site.lastModifiedDateTime).toLocaleDateString('nl-NL', {
                            year: 'numeric',
                            month: 'short',
                            day: 'numeric',
                          })
                        : 'Onbekend'}
                    </td>

                    {/* Action button */}
                    <td className="py-4 px-4 text-right whitespace-nowrap">
                      <button
                        onClick={() => setSelectedSite(site)}
                        className="px-3 py-1.5 rounded-xl bg-card border border-input text-heading hover:border-cyan hover:text-cyan font-semibold text-xs transition-colors inline-flex items-center gap-1.5 shadow-sm"
                      >
                        <NoppaIcon name="security-shield" size={14} />
                        <span>Toegang &amp; Externen</span>
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Modal 1: 4-Level Site Permissions & External Auditor */}
      {selectedSite && (
        <SitePermissionsModal
          site={selectedSite}
          isOpen={Boolean(selectedSite)}
          onClose={() => setSelectedSite(null)}
        />
      )}

      {/* Modal 2: Tenant-Wide External Search & Auditor */}
      <TenantExternalAuditModal
        isOpen={isTenantAuditOpen}
        onClose={() => setIsTenantAuditOpen(false)}
        onSelectSiteToInspect={(targetSite) => {
          setIsTenantAuditOpen(false);
          const found = sites.find((s) => s.id === targetSite.id);
          setSelectedSite(found || targetSite);
        }}
      />

      {/* Modal 3: Tenant Connection & Entra ID Credentials */}
      <TenantSettingsModal
        isOpen={isSettingsOpen}
        onClose={() => setIsSettingsOpen(false)}
        onSaved={() => {
          loadSites();
        }}
      />
    </div>
  );
}
