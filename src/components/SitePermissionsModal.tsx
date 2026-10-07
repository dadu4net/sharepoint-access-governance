'use client';

import React, { useState, useEffect, useMemo } from 'react';
import {
  X,
  ExternalLink,
  Shield,
  Users,
  Globe,
  Link2,
  FileText,
  Folder,
  RefreshCw,
  Search,
  Filter,
  AlertTriangle,
  CheckCircle2,
  Lock,
  Unlock,
  ChevronRight,
  Eye,
  Key,
} from 'lucide-react';
import {
  GraphSite,
  SiteAccessAuditSummary,
  SharePointAccessEntity,
  SharePointContainer,
  SharePointItemPermissionSummary,
} from '@/lib/graph/client';
import { NoppaIcon } from '@/components/common/NoppaIcon';

interface SitePermissionsModalProps {
  site: GraphSite | null;
  isOpen: boolean;
  onClose: () => void;
}

import { useMsal } from '@azure/msal-react';
import { fetchSitePermissionsAudit, fetchDriveItemsPermissions } from '@/lib/graph/clientService';

export function SitePermissionsModal({ site, isOpen, onClose }: SitePermissionsModalProps) {
  const { instance, accounts } = useMsal();
  
  const [auditData, setAuditData] = useState<SiteAccessAuditSummary | null>(null);
  const [isLoadingAudit, setIsLoadingAudit] = useState(false);
  const [auditError, setAuditError] = useState<string | null>(null);
  const [isLiveApi, setIsLiveApi] = useState<boolean>(false);

  // Navigatie & Filters
  const [activeTab, setActiveTab] = useState<'site' | 'containers' | 'items'>('site');
  const [onlyExternalFilter, setOnlyExternalFilter] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');

  // Item niveau state
  const [selectedDriveId, setSelectedDriveId] = useState<string>('');
  const [itemsList, setItemsList] = useState<SharePointItemPermissionSummary[]>([]);
  const [isLoadingItems, setIsLoadingItems] = useState(false);
  const [itemsError, setItemsError] = useState<string | null>(null);

  // Container detail inspectie
  const [expandedContainerId, setExpandedContainerId] = useState<string | null>(null);

  useEffect(() => {
    if (isOpen && site) {
      loadAudit();
      setActiveTab('site');
      setOnlyExternalFilter(false);
      setSearchQuery('');
      setExpandedContainerId(null);
    }
  }, [isOpen, site]);

  const loadAudit = async () => {
    if (!site) return;
    setIsLoadingAudit(true);
    setAuditError(null);
    try {
      if (accounts.length > 0) {
        setIsLiveApi(true);
      }
      
      const data = await fetchSitePermissionsAudit(instance, accounts, site.id);
      
      setAuditData(data as any);

      if (data?.containers?.length > 0) {
        const firstLib = data.containers.find((c: any) => c.type === 'documentLibrary') || data.containers[0];
        setSelectedDriveId(firstLib.id);
      }
    } catch (err: any) {
      console.error(err);
      setAuditError(err.message || 'Fout bij het ophalen van permissie data.');
    } finally {
      setIsLoadingAudit(false);
    }
  };

  useEffect(() => {
    if (isOpen && site && activeTab === 'items' && selectedDriveId) {
      loadItems(selectedDriveId);
    }
  }, [isOpen, site, activeTab, selectedDriveId]);

  const loadItems = async (driveId: string) => {
    if (!site) return;
    setIsLoadingItems(true);
    setItemsError(null);
    try {
      const items = await fetchDriveItemsPermissions(instance, accounts, site.id, driveId);
      setItemsList(items as any[]);
    } catch (err: any) {
      console.error(err);
      setItemsError(err.message || 'Fout bij ophalen van items en deellinks.');
    } finally {
      setIsLoadingItems(false);
    }
  };

  const filteredSitePermissions = useMemo(() => {
    if (!auditData?.sitePermissions) return [];
    return auditData.sitePermissions.filter((p) => {
      if (onlyExternalFilter && !p.isExternal) return false;
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const matchName = (p.displayName || p.principalName || '').toLowerCase().includes(q);
        const matchEmail = (p.email || '').toLowerCase().includes(q);
        const matchRole = (p.role || '').toLowerCase().includes(q);
        return matchName || matchEmail || matchRole;
      }
      return true;
    });
  }, [auditData?.sitePermissions, onlyExternalFilter, searchQuery]);

  const filteredContainers = useMemo(() => {
    if (!auditData?.containers) return [];
    return auditData.containers.filter((c) => {
      if (onlyExternalFilter && (!c.externalPermissionsCount || c.externalPermissionsCount === 0)) {
        return false;
      }
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        return c.name.toLowerCase().includes(q);
      }
      return true;
    });
  }, [auditData?.containers, onlyExternalFilter, searchQuery]);

  const filteredItems = useMemo(() => {
    return itemsList.filter((item) => {
      if (onlyExternalFilter && item.externalSharesCount === 0) return false;
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const matchName = item.name.toLowerCase().includes(q);
        const matchPerm = item.permissions.some(
          (p) =>
            (p.displayName || p.principalName || '').toLowerCase().includes(q) ||
            (p.email || '').toLowerCase().includes(q)
        );
        return matchName || matchPerm;
      }
      return true;
    });
  }, [itemsList, onlyExternalFilter, searchQuery]);

  if (!isOpen || !site) return null;

  return (
    <div
      style={{
        position: 'fixed',
        inset: 0,
        backgroundColor: 'rgba(15, 42, 102, 0.55)',
        backdropFilter: 'blur(8px)',
        zIndex: 1100,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        padding: '20px',
      }}
    >
      <div
        className="card-portal"
        style={{
          width: '100%',
          maxWidth: '1140px',
          maxHeight: '92vh',
          display: 'flex',
          flexDirection: 'column',
          backgroundColor: 'var(--bg-card, #FFFFFF)',
          borderRadius: '14px',
          boxShadow: '0 24px 48px -12px rgba(15, 42, 102, 0.25)',
          border: '1px solid var(--border, #E4E9F2)',
          overflow: 'hidden',
          animation: 'fadeIn 0.2s cubic-bezier(0.16, 1, 0.3, 1)',
        }}
      >
        {/* Header */}
        <div
          style={{
            padding: '18px 24px',
            borderBottom: '1px solid var(--border, #E4E9F2)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            background: 'linear-gradient(180deg, var(--bg-primary, #F1F5FB) 0%, var(--bg-card, #FFFFFF) 100%)',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
            <div
              style={{
                width: '42px',
                height: '42px',
                borderRadius: '10px',
                backgroundColor: 'rgba(32, 96, 224, 0.1)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                border: '1px solid rgba(32, 96, 224, 0.2)',
              }}
            >
              <NoppaIcon name="shield" size={24} />
            </div>
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <h2
                  style={{
                    margin: 0,
                    fontSize: '1.1875rem',
                    fontWeight: 700,
                    color: 'var(--text-heading, #0F2A66)',
                    letterSpacing: '-0.02em',
                  }}
                >
                  Toegangsbeheer & Externen Audit
                </h2>
                <span
                  className={`badge ${isLiveApi ? 'badge-cyan' : 'badge-primary'}`}
                  style={{ fontSize: '0.6875rem', padding: '2px 8px' }}
                >
                  {isLiveApi ? '🟢 Live Graph API' : '⚡ Demo Data'}
                </span>
                <span className="badge" style={{ fontSize: '0.6875rem', backgroundColor: 'var(--bg-primary, #F1F5FB)', color: 'var(--text-main, #1A2440)' }}>
                  {site.template === 'Communication' ? 'Communicatiesite' : 'Teamsite'}
                </span>
              </div>
              <div
                style={{
                  fontSize: '0.8125rem',
                  color: 'var(--color-slate, #6B7280)',
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px',
                  marginTop: '2px',
                }}
              >
                <span style={{ fontWeight: 600, color: 'var(--text-main, #1A2440)' }}>{site.title}</span>
                <span>•</span>
                <a
                  href={site.url}
                  target="_blank"
                  rel="noreferrer"
                  style={{
                    color: 'var(--color-royal, #2060E0)',
                    textDecoration: 'none',
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '3px',
                  }}
                >
                  {site.url} <ExternalLink size={11} />
                </a>
              </div>
            </div>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <button
              type="button"
              className="btn btn-secondary"
              onClick={loadAudit}
              disabled={isLoadingAudit}
              title="Permissies opnieuw ophalen"
              style={{ padding: '6px 12px', fontSize: '0.75rem', display: 'flex', alignItems: 'center', gap: '6px' }}
            >
              <RefreshCw size={13} className={isLoadingAudit ? 'spin' : ''} />
              Vernieuwen
            </button>
            <button
              type="button"
              onClick={onClose}
              style={{
                background: 'none',
                border: 'none',
                cursor: 'pointer',
                color: 'var(--color-slate, #6B7280)',
                padding: '6px',
                borderRadius: '6px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
              }}
              aria-label="Sluiten"
            >
              <X size={20} />
            </button>
          </div>
        </div>

        {/* KPI Ribbon */}
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(4, 1fr)',
            gap: '12px',
            padding: '14px 24px',
            backgroundColor: 'var(--bg-primary, #F1F5FB)',
            borderBottom: '1px solid var(--border, #E4E9F2)',
          }}
        >
          <div
            style={{
              backgroundColor: 'var(--bg-card, #FFFFFF)',
              padding: '12px 14px',
              borderRadius: '10px',
              border: '1px solid var(--border, #E4E9F2)',
              display: 'flex',
              alignItems: 'center',
              gap: '12px',
            }}
          >
            <div
              style={{
                width: '36px',
                height: '36px',
                borderRadius: '8px',
                backgroundColor: 'rgba(32, 96, 224, 0.1)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
              }}
            >
              <Users size={18} color="var(--color-royal, #2060E0)" />
            </div>
            <div>
              <div style={{ fontSize: '0.6875rem', fontWeight: 600, color: 'var(--color-slate, #6B7280)', textTransform: 'uppercase' }}>
                Geautoriseerd
              </div>
              <div style={{ fontSize: '1.25rem', fontWeight: 700, color: 'var(--text-heading, #0F2A66)', fontVariantNumeric: 'tabular-nums' }}>
                {auditData?.totalPermissions ?? '-'}
              </div>
            </div>
          </div>

          <div
            style={{
              backgroundColor: 'var(--bg-card, #FFFFFF)',
              padding: '12px 14px',
              borderRadius: '10px',
              border: (auditData?.externalCount || 0) > 0 ? '1px solid rgba(0, 208, 240, 0.4)' : '1px solid var(--border, #E4E9F2)',
              display: 'flex',
              alignItems: 'center',
              gap: '12px',
              position: 'relative',
              overflow: 'hidden',
            }}
          >
            {(auditData?.externalCount || 0) > 0 && (
              <div
                style={{
                  position: 'absolute',
                  top: 0,
                  left: 0,
                  right: 0,
                  height: '2px',
                  background: 'var(--gradient-signature, linear-gradient(90deg, #00D0F0, #2060E0))',
                }}
              />
            )}
            <div
              style={{
                width: '36px',
                height: '36px',
                borderRadius: '8px',
                backgroundColor: (auditData?.externalCount || 0) > 0 ? 'rgba(0, 208, 240, 0.15)' : 'rgba(107, 114, 128, 0.1)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
              }}
            >
              <Globe size={18} color={(auditData?.externalCount || 0) > 0 ? '#0099B8' : '#6B7280'} />
            </div>
            <div>
              <div style={{ fontSize: '0.6875rem', fontWeight: 600, color: 'var(--color-slate, #6B7280)', textTransform: 'uppercase' }}>
                Externe Gasten
              </div>
              <div
                style={{
                  fontSize: '1.25rem',
                  fontWeight: 700,
                  color: (auditData?.externalCount || 0) > 0 ? '#007A93' : 'var(--text-heading, #0F2A66)',
                  fontVariantNumeric: 'tabular-nums',
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px',
                }}
              >
                {auditData?.externalCount ?? '-'}
                {(auditData?.externalCount || 0) > 0 && (
                  <span className="badge badge-cyan" style={{ fontSize: '0.625rem', padding: '1px 6px' }}>
                    Aandacht
                  </span>
                )}
              </div>
            </div>
          </div>

          <div
            style={{
              backgroundColor: 'var(--bg-card, #FFFFFF)',
              padding: '12px 14px',
              borderRadius: '10px',
              border: (auditData?.sharingLinksCount || 0) > 0 ? '1px solid rgba(32, 96, 224, 0.3)' : '1px solid var(--border, #E4E9F2)',
              display: 'flex',
              alignItems: 'center',
              gap: '12px',
            }}
          >
            <div
              style={{
                width: '36px',
                height: '36px',
                borderRadius: '8px',
                backgroundColor: 'rgba(32, 96, 224, 0.1)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
              }}
            >
              <Link2 size={18} color="var(--color-royal, #2060E0)" />
            </div>
            <div>
              <div style={{ fontSize: '0.6875rem', fontWeight: 600, color: 'var(--color-slate, #6B7280)', textTransform: 'uppercase' }}>
                Deellinks
              </div>
              <div style={{ fontSize: '1.25rem', fontWeight: 700, color: 'var(--text-heading, #0F2A66)', fontVariantNumeric: 'tabular-nums' }}>
                {auditData?.sharingLinksCount ?? '-'}
              </div>
            </div>
          </div>

          <div
            style={{
              backgroundColor: 'var(--bg-card, #FFFFFF)',
              padding: '12px 14px',
              borderRadius: '10px',
              border: (auditData?.uniquePermissionItemsCount || 0) > 0 ? '1px solid rgba(245, 158, 11, 0.4)' : '1px solid var(--border, #E4E9F2)',
              display: 'flex',
              alignItems: 'center',
              gap: '12px',
            }}
          >
            <div
              style={{
                width: '36px',
                height: '36px',
                borderRadius: '8px',
                backgroundColor: (auditData?.uniquePermissionItemsCount || 0) > 0 ? 'rgba(245, 158, 11, 0.12)' : 'rgba(107, 114, 128, 0.1)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
              }}
            >
              <Unlock size={18} color={(auditData?.uniquePermissionItemsCount || 0) > 0 ? '#D97706' : '#6B7280'} />
            </div>
            <div>
              <div style={{ fontSize: '0.6875rem', fontWeight: 600, color: 'var(--color-slate, #6B7280)', textTransform: 'uppercase' }}>
                Unieke Rechten
              </div>
              <div
                style={{
                  fontSize: '1.25rem',
                  fontWeight: 700,
                  color: (auditData?.uniquePermissionItemsCount || 0) > 0 ? '#D97706' : 'var(--text-heading, #0F2A66)',
                  fontVariantNumeric: 'tabular-nums',
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px',
                }}
              >
                {auditData?.uniquePermissionItemsCount ?? '-'}
                {(auditData?.uniquePermissionItemsCount || 0) > 0 && (
                  <span style={{ fontSize: '0.625rem', backgroundColor: 'rgba(245, 158, 11, 0.15)', color: '#B45309', padding: '1px 6px', borderRadius: '4px', fontWeight: 600 }}>
                    Uniek
                  </span>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* Toolbar */}
        <div
          style={{
            padding: '12px 24px',
            borderBottom: '1px solid var(--border, #E4E9F2)',
            display: 'flex',
            flexWrap: 'wrap',
            alignItems: 'center',
            justifyContent: 'space-between',
            gap: '12px',
          }}
        >
          <div style={{ display: 'flex', gap: '6px' }}>
            <button
              type="button"
              className={`btn ${activeTab === 'site' ? 'btn-primary' : 'btn-secondary'}`}
              onClick={() => setActiveTab('site')}
              style={{ padding: '7px 14px', fontSize: '0.8125rem', display: 'flex', alignItems: 'center', gap: '6px' }}
            >
              <Users size={14} /> Site Niveau (Groepen & Gebruikers)
            </button>
            <button
              type="button"
              className={`btn ${activeTab === 'containers' ? 'btn-primary' : 'btn-secondary'}`}
              onClick={() => setActiveTab('containers')}
              style={{ padding: '7px 14px', fontSize: '0.8125rem', display: 'flex', alignItems: 'center', gap: '6px' }}
            >
              <Folder size={14} /> Bibliotheken & Lijsten ({auditData?.containers?.length || 0})
            </button>
            <button
              type="button"
              className={`btn ${activeTab === 'items' ? 'btn-primary' : 'btn-secondary'}`}
              onClick={() => setActiveTab('items')}
              style={{ padding: '7px 14px', fontSize: '0.8125rem', display: 'flex', alignItems: 'center', gap: '6px' }}
            >
              <FileText size={14} /> Bestanden & Mappen (Item Niveau)
            </button>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <div style={{ position: 'relative', width: '220px' }}>
              <Search
                size={14}
                style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: '#6B7280' }}
              />
              <input
                type="text"
                className="portal-input"
                placeholder="Zoek gebruiker, rol..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                style={{ paddingLeft: '32px', fontSize: '0.8125rem', height: '34px' }}
              />
            </div>

            <button
              type="button"
              onClick={() => setOnlyExternalFilter(!onlyExternalFilter)}
              style={{
                height: '34px',
                padding: '0 12px',
                borderRadius: '8px',
                fontSize: '0.75rem',
                fontWeight: 600,
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                gap: '6px',
                border: onlyExternalFilter ? '1.5px solid #00D0F0' : '1px solid var(--border, #E4E9F2)',
                backgroundColor: onlyExternalFilter ? 'rgba(0, 208, 240, 0.12)' : 'var(--bg-primary, #F1F5FB)',
                color: onlyExternalFilter ? '#007A93' : 'var(--text-main, #1A2440)',
                transition: 'all 0.15s ease',
              }}
            >
              <Filter size={13} color={onlyExternalFilter ? '#0099B8' : '#6B7280'} />
              <span>Alleen Externen & Deellinks</span>
              {onlyExternalFilter && (
                <span
                  style={{
                    backgroundColor: '#00D0F0',
                    color: '#0F2A66',
                    fontSize: '0.625rem',
                    padding: '1px 6px',
                    borderRadius: '10px',
                    fontWeight: 700,
                  }}
                >
                  Actief
                </span>
              )}
            </button>
          </div>
        </div>

        {/* Content */}
        <div style={{ flex: 1, overflowY: 'auto', padding: '20px 24px' }}>
          {isLoadingAudit ? (
            <div style={{ padding: '60px 0', textAlign: 'center', color: 'var(--color-slate, #6B7280)' }}>
              <RefreshCw size={28} className="spin" style={{ margin: '0 auto 12px auto', color: 'var(--color-royal, #2060E0)' }} />
              <div style={{ fontWeight: 600, color: 'var(--text-heading, #0F2A66)' }}>
                Rechten en machtigingen ophalen via Microsoft Graph...
              </div>
              <div style={{ fontSize: '0.8125rem' }}>Analyseren van groepen, gasten en deellinks...</div>
            </div>
          ) : auditError ? (
            <div
              style={{
                padding: '16px',
                borderRadius: '8px',
                backgroundColor: 'rgba(239, 68, 68, 0.1)',
                border: '1px solid rgba(239, 68, 68, 0.3)',
                color: '#DC2626',
                fontSize: '0.875rem',
                display: 'flex',
                alignItems: 'center',
                gap: '10px',
              }}
            >
              <AlertTriangle size={18} />
              <span>{auditError}</span>
            </div>
          ) : (
            <>
              {/* TAB 1: SITE NIVEAU */}
              {activeTab === 'site' && (
                <div>
                  <div style={{ marginBottom: '14px' }}>
                    <h3 style={{ margin: 0, fontSize: '0.9375rem', fontWeight: 700, color: 'var(--text-heading, #0F2A66)' }}>
                      Site Machtigingstoewijzingen ({filteredSitePermissions.length})
                    </h3>
                    <p style={{ margin: '2px 0 0 0', fontSize: '0.75rem', color: 'var(--color-slate, #6B7280)' }}>
                      Groepen, directe accounts en gasten die toegang hebben tot de hoofdlocatie van de SharePoint site.
                    </p>
                  </div>

                  <div className="table-container" style={{ borderRadius: '10px', border: '1px solid var(--border, #E4E9F2)' }}>
                    <table className="portal-table" style={{ width: '100%', fontSize: '0.8125rem' }}>
                      <thead>
                        <tr>
                          <th>Identiteit / Naam</th>
                          <th>Type</th>
                          <th>Toegewezen Rol</th>
                          <th>Status & Externe Toegang</th>
                          <th>Overerving</th>
                          <th>Details</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredSitePermissions.length === 0 ? (
                          <tr>
                            <td colSpan={6} style={{ textAlign: 'center', padding: '36px', color: 'var(--color-slate, #6B7280)' }}>
                              Geen toewijzingen gevonden die voldoen aan het huidige filter.
                            </td>
                          </tr>
                        ) : (
                          filteredSitePermissions.map((perm) => (
                            <tr
                              key={perm.id}
                              style={{
                                backgroundColor: perm.isExternal ? 'rgba(0, 208, 240, 0.03)' : undefined,
                              }}
                            >
                              <td>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                  <div
                                    style={{
                                      width: '28px',
                                      height: '28px',
                                      borderRadius: '6px',
                                      backgroundColor: perm.isExternal
                                        ? 'rgba(0, 208, 240, 0.15)'
                                        : 'rgba(32, 96, 224, 0.1)',
                                      display: 'flex',
                                      alignItems: 'center',
                                      justifyContent: 'center',
                                      flexShrink: 0,
                                    }}
                                  >
                                    {perm.userType === 'sharingLink' ? (
                                      <Link2 size={14} color={perm.isExternal ? '#0099B8' : '#2060E0'} />
                                    ) : perm.userType === 'guest' ? (
                                      <Globe size={14} color="#0099B8" />
                                    ) : perm.userType === 'group' || perm.userType === 'siteGroup' ? (
                                      <Users size={14} color="#2060E0" />
                                    ) : (
                                      <Key size={14} color="#6B7280" />
                                    )}
                                  </div>
                                  <div>
                                    <div style={{ fontWeight: 600, color: 'var(--text-heading, #0F2A66)' }}>
                                      {perm.displayName}
                                    </div>
                                    {perm.email && (
                                      <div style={{ fontSize: '0.71875rem', color: 'var(--color-slate, #6B7280)' }}>
                                        {perm.email}
                                      </div>
                                    )}
                                  </div>
                                </div>
                              </td>
                              <td>
                                <span className="badge" style={{ fontSize: '0.6875rem' }}>
                                  {perm.userType === 'siteGroup'
                                    ? 'Sitegroep'
                                    : perm.userType === 'group'
                                    ? 'M365 Groep'
                                    : perm.userType === 'guest'
                                    ? 'Externe Gast'
                                    : perm.userType === 'sharingLink'
                                    ? 'Koppeling'
                                    : 'Gebruiker'}
                                </span>
                              </td>
                              <td>
                                <span
                                  className={`badge ${
                                    perm.role === 'owner'
                                      ? 'badge-primary'
                                      : perm.role === 'editor'
                                      ? 'badge-cyan'
                                      : ''
                                  }`}
                                  style={{
                                    fontSize: '0.6875rem',
                                    backgroundColor:
                                      perm.role === 'reader' ? 'var(--bg-primary, #F1F5FB)' : undefined,
                                    color:
                                      perm.role === 'reader' ? 'var(--text-main, #1A2440)' : undefined,
                                  }}
                                >
                                  {perm.role === 'owner'
                                    ? '👑 Volledig Beheer'
                                    : perm.role === 'editor'
                                    ? '✏️ Bewerker'
                                    : '👁️ Lezer'}
                                </span>
                              </td>
                              <td>
                                {perm.isExternal ? (
                                  <div style={{ display: 'flex', flexDirection: 'column', gap: '3px' }}>
                                    <span
                                      className="badge badge-cyan"
                                      style={{
                                        fontSize: '0.6875rem',
                                        display: 'inline-flex',
                                        alignItems: 'center',
                                        gap: '4px',
                                        width: 'fit-content',
                                      }}
                                    >
                                      <Globe size={11} /> Externe Toegang
                                    </span>
                                    {perm.sharingLink && (
                                      <span style={{ fontSize: '0.6875rem', color: '#007A93' }}>
                                        Scope: {perm.sharingLink.scope}
                                      </span>
                                    )}
                                  </div>
                                ) : (
                                  <span
                                    style={{
                                      fontSize: '0.75rem',
                                      color: 'var(--color-slate, #6B7280)',
                                      display: 'flex',
                                      alignItems: 'center',
                                      gap: '4px',
                                    }}
                                  >
                                    <CheckCircle2 size={12} color="#10B981" /> Intern Tenant
                                  </span>
                                )}
                              </td>
                              <td>
                                {perm.inherited ? (
                                  <span style={{ fontSize: '0.71875rem', color: 'var(--color-slate, #6B7280)', display: 'flex', alignItems: 'center', gap: '4px' }}>
                                    <Lock size={12} /> {perm.inheritedFrom || 'Geërfd'}
                                  </span>
                                ) : (
                                  <span style={{ fontSize: '0.71875rem', color: 'var(--color-royal, #2060E0)', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '4px' }}>
                                    <Key size={12} /> Direct op Site
                                  </span>
                                )}
                              </td>
                              <td>
                                {perm.sharingLink?.webUrl ? (
                                  <a
                                    href={perm.sharingLink.webUrl}
                                    target="_blank"
                                    rel="noreferrer"
                                    style={{
                                      fontSize: '0.75rem',
                                      color: 'var(--color-royal, #2060E0)',
                                      textDecoration: 'none',
                                      display: 'inline-flex',
                                      alignItems: 'center',
                                      gap: '3px',
                                    }}
                                  >
                                    Koppeling openen <ExternalLink size={10} />
                                  </a>
                                ) : (
                                  <span style={{ fontSize: '0.71875rem', color: 'var(--color-slate, #6B7280)' }}>
                                    ID: {perm.id.substring(0, 10)}...
                                  </span>
                                )}
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* TAB 2: BIBLIOTHEKEN & LIJSTEN */}
              {activeTab === 'containers' && (
                <div>
                  <div style={{ marginBottom: '14px' }}>
                    <h3 style={{ margin: 0, fontSize: '0.9375rem', fontWeight: 700, color: 'var(--text-heading, #0F2A66)' }}>
                      Documentenbibliotheken & SharePoint Lijsten ({filteredContainers.length})
                    </h3>
                    <p style={{ margin: '2px 0 0 0', fontSize: '0.75rem', color: 'var(--color-slate, #6B7280)' }}>
                      Controleer per bibliotheek of lijst of de overerving actief is of dat er unieke rechten en externe deellinks zijn geconfigureerd.
                    </p>
                  </div>

                  <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                    {filteredContainers.length === 0 ? (
                      <div style={{ textAlign: 'center', padding: '36px', color: 'var(--color-slate, #6B7280)' }}>
                        Geen bibliotheken of lijsten gevonden voor het geselecteerde filter.
                      </div>
                    ) : (
                      filteredContainers.map((container) => {
                        const isExpanded = expandedContainerId === container.id;
                        return (
                          <div
                            key={container.id}
                            style={{
                              backgroundColor: 'var(--bg-card, #FFFFFF)',
                              borderRadius: '10px',
                              border: container.hasUniquePermissions
                                ? '1.5px solid rgba(245, 158, 11, 0.4)'
                                : '1px solid var(--border, #E4E9F2)',
                              overflow: 'hidden',
                              boxShadow: '0 2px 6px rgba(15, 42, 102, 0.04)',
                            }}
                          >
                            <div
                              style={{
                                padding: '14px 18px',
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'space-between',
                                background: container.hasUniquePermissions
                                  ? 'linear-gradient(90deg, rgba(245, 158, 11, 0.05) 0%, transparent 100%)'
                                  : undefined,
                              }}
                            >
                              <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                                <div
                                  style={{
                                    width: '36px',
                                    height: '36px',
                                    borderRadius: '8px',
                                    backgroundColor: container.type === 'documentLibrary' ? 'rgba(32, 96, 224, 0.1)' : 'rgba(0, 208, 240, 0.1)',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                  }}
                                >
                                  {container.type === 'documentLibrary' ? (
                                    <Folder size={18} color="var(--color-royal, #2060E0)" />
                                  ) : (
                                    <FileText size={18} color="#0099B8" />
                                  )}
                                </div>
                                <div>
                                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                    <span style={{ fontWeight: 600, color: 'var(--text-heading, #0F2A66)', fontSize: '0.875rem' }}>
                                      {container.name}
                                    </span>
                                    <span className="badge" style={{ fontSize: '0.625rem' }}>
                                      {container.type === 'documentLibrary' ? 'Documentenbibliotheek' : 'Lijst'}
                                    </span>
                                    {container.webUrl && (
                                      <a
                                        href={container.webUrl}
                                        target="_blank"
                                        rel="noreferrer"
                                        style={{ color: 'var(--color-royal, #2060E0)', display: 'inline-flex', alignItems: 'center' }}
                                        title="Open in SharePoint"
                                      >
                                        <ExternalLink size={12} />
                                      </a>
                                    )}
                                  </div>
                                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginTop: '3px', fontSize: '0.75rem' }}>
                                    {container.hasUniquePermissions ? (
                                      <span style={{ color: '#D97706', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '4px' }}>
                                        <Unlock size={12} /> Unieke rechten (Overerving verbroken)
                                      </span>
                                    ) : (
                                      <span style={{ color: '#10B981', display: 'flex', alignItems: 'center', gap: '4px' }}>
                                        <Lock size={12} /> Overgenomen van site
                                      </span>
                                    )}

                                    {(container.externalPermissionsCount || 0) > 0 && (
                                      <span className="badge badge-cyan" style={{ fontSize: '0.625rem', padding: '1px 6px' }}>
                                        🌐 {container.externalPermissionsCount} Externe(n) / Deellink(s)
                                      </span>
                                    )}
                                  </div>
                                </div>
                              </div>

                              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                {container.type === 'documentLibrary' && (
                                  <button
                                    type="button"
                                    className="btn btn-secondary"
                                    onClick={() => {
                                      setSelectedDriveId(container.id);
                                      setActiveTab('items');
                                    }}
                                    style={{ padding: '6px 10px', fontSize: '0.75rem', display: 'flex', alignItems: 'center', gap: '5px' }}
                                  >
                                    <FileText size={13} /> Inspecteer Bestanden
                                  </button>
                                )}

                                {container.permissions && container.permissions.length > 0 && (
                                  <button
                                    type="button"
                                    className="btn btn-secondary"
                                    onClick={() => setExpandedContainerId(isExpanded ? null : container.id)}
                                    style={{ padding: '6px 10px', fontSize: '0.75rem', display: 'flex', alignItems: 'center', gap: '4px' }}
                                  >
                                    <Eye size={13} />
                                    {isExpanded ? 'Verberg Permissies' : `Rechten (${container.permissions.length})`}
                                  </button>
                                )}
                              </div>
                            </div>

                            {/* Detail permissies container */}
                            {isExpanded && container.permissions && (
                              <div style={{ borderTop: '1px solid var(--border, #E4E9F2)', padding: '12px 18px', backgroundColor: 'var(--bg-primary, #F1F5FB)' }}>
                                <div style={{ fontSize: '0.75rem', fontWeight: 600, color: 'var(--text-heading, #0F2A66)', marginBottom: '8px' }}>
                                  Machtigingen voor "{container.name}":
                                </div>
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                                  {container.permissions.map((p) => (
                                    <div
                                      key={p.id}
                                      style={{
                                        display: 'flex',
                                        alignItems: 'center',
                                        justifyContent: 'space-between',
                                        padding: '8px 12px',
                                        borderRadius: '6px',
                                        backgroundColor: 'var(--bg-card, #FFFFFF)',
                                        border: '1px solid var(--border, #E4E9F2)',
                                        fontSize: '0.75rem',
                                      }}
                                    >
                                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                        {p.isExternal ? <Globe size={14} color="#0099B8" /> : <Users size={14} color="#2060E0" />}
                                        <span style={{ fontWeight: 600, color: 'var(--text-heading, #0F2A66)' }}>{p.displayName}</span>
                                        {p.email && <span style={{ color: 'var(--color-slate, #6B7280)' }}>({p.email})</span>}
                                      </div>
                                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                        <span className={`badge ${p.role === 'owner' ? 'badge-primary' : p.role === 'editor' ? 'badge-cyan' : ''}`} style={{ fontSize: '0.625rem' }}>
                                          {p.role}
                                        </span>
                                        {p.isExternal && (
                                          <span className="badge badge-cyan" style={{ fontSize: '0.625rem' }}>
                                            Extern
                                          </span>
                                        )}
                                      </div>
                                    </div>
                                  ))}
                                </div>
                              </div>
                            )}
                          </div>
                        );
                      })
                    )}
                  </div>
                </div>
              )}

              {/* TAB 3: BESTANDEN & MAPPEN */}
              {activeTab === 'items' && (
                <div>
                  <div style={{ marginBottom: '14px', display: 'flex', flexWrap: 'wrap', alignItems: 'center', justifyContent: 'space-between', gap: '12px' }}>
                    <div>
                      <h3 style={{ margin: 0, fontSize: '0.9375rem', fontWeight: 700, color: 'var(--text-heading, #0F2A66)' }}>
                        Bestand & Map Rechten & Deellinks
                      </h3>
                      <p style={{ margin: '2px 0 0 0', fontSize: '0.75rem', color: 'var(--color-slate, #6B7280)' }}>
                        Selecteer een documentenbibliotheek om individuele bestanden met unieke toegangsrechten en deellinks te inspecteren.
                      </p>
                    </div>

                    {auditData?.containers && auditData.containers.length > 0 && (
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <span style={{ fontSize: '0.75rem', fontWeight: 600, color: 'var(--color-slate, #6B7280)' }}>
                          Bibliotheek:
                        </span>
                        <select
                          className="portal-select"
                          value={selectedDriveId}
                          onChange={(e) => setSelectedDriveId(e.target.value)}
                          style={{ fontSize: '0.8125rem', padding: '6px 12px', minWidth: '220px' }}
                        >
                          {auditData.containers
                            .filter((c) => c.type === 'documentLibrary')
                            .map((c) => (
                              <option key={c.id} value={c.id}>
                                📁 {c.name} {c.hasUniquePermissions ? '⚠️ (Uniek)' : ''}
                              </option>
                            ))}
                        </select>
                      </div>
                    )}
                  </div>

                  {isLoadingItems ? (
                    <div style={{ padding: '40px 0', textAlign: 'center', color: 'var(--color-slate, #6B7280)' }}>
                      <RefreshCw size={24} className="spin" style={{ margin: '0 auto 10px auto', color: 'var(--color-royal, #2060E0)' }} />
                      <div style={{ fontSize: '0.8125rem' }}>Bestanden en deellinks ophalen van bibliotheek...</div>
                    </div>
                  ) : itemsError ? (
                    <div style={{ padding: '14px', borderRadius: '8px', backgroundColor: 'rgba(239, 68, 68, 0.1)', color: '#DC2626', fontSize: '0.8125rem' }}>
                      {itemsError}
                    </div>
                  ) : (
                    <div className="table-container" style={{ borderRadius: '10px', border: '1px solid var(--border, #E4E9F2)' }}>
                      <table className="portal-table" style={{ width: '100%', fontSize: '0.8125rem' }}>
                        <thead>
                          <tr>
                            <th>Bestand / Map</th>
                            <th>Overerving</th>
                            <th>Externe Deling & Links</th>
                            <th>Geautoriseerde Entiteiten & Permissies</th>
                            <th>Actie</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filteredItems.length === 0 ? (
                            <tr>
                              <td colSpan={5} style={{ textAlign: 'center', padding: '36px', color: 'var(--color-slate, #6B7280)' }}>
                                Geen bestanden of mappen gevonden in deze bibliotheek die voldoen aan het filter.
                              </td>
                            </tr>
                          ) : (
                            filteredItems.map((item) => (
                              <tr
                                key={item.id}
                                style={{
                                  backgroundColor: item.hasUniquePermissions
                                    ? 'rgba(245, 158, 11, 0.03)'
                                    : undefined,
                                }}
                              >
                                <td>
                                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                    <div
                                      style={{
                                        width: '28px',
                                        height: '28px',
                                        borderRadius: '6px',
                                        backgroundColor: item.isFolder
                                          ? 'rgba(32, 96, 224, 0.1)'
                                          : 'rgba(0, 208, 240, 0.1)',
                                        display: 'flex',
                                        alignItems: 'center',
                                        justifyContent: 'center',
                                        flexShrink: 0,
                                      }}
                                    >
                                      {item.isFolder ? (
                                        <Folder size={14} color="var(--color-royal, #2060E0)" />
                                      ) : (
                                        <FileText size={14} color="#0099B8" />
                                      )}
                                    </div>
                                    <div>
                                      <div style={{ fontWeight: 600, color: 'var(--text-heading, #0F2A66)' }}>
                                        {item.name}
                                      </div>
                                      <div style={{ fontSize: '0.6875rem', color: 'var(--color-slate, #6B7280)' }}>
                                        {item.isFolder ? 'Map' : 'Bestand'}
                                      </div>
                                    </div>
                                  </div>
                                </td>
                                <td>
                                  {item.hasUniquePermissions ? (
                                    <span
                                      style={{
                                        color: '#D97706',
                                        fontWeight: 600,
                                        fontSize: '0.75rem',
                                        display: 'inline-flex',
                                        alignItems: 'center',
                                        gap: '4px',
                                        backgroundColor: 'rgba(245, 158, 11, 0.1)',
                                        padding: '2px 8px',
                                        borderRadius: '4px',
                                      }}
                                    >
                                      <Unlock size={11} /> Unieke Item Rechten
                                    </span>
                                  ) : (
                                    <span
                                      style={{
                                        color: '#10B981',
                                        fontSize: '0.75rem',
                                        display: 'inline-flex',
                                        alignItems: 'center',
                                        gap: '4px',
                                      }}
                                    >
                                      <Lock size={11} /> Overgenomen
                                    </span>
                                  )}
                                </td>
                                <td>
                                  {item.externalSharesCount > 0 ? (
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '3px' }}>
                                      <span
                                        className="badge badge-cyan"
                                        style={{
                                          fontSize: '0.6875rem',
                                          display: 'inline-flex',
                                          alignItems: 'center',
                                          gap: '4px',
                                          width: 'fit-content',
                                        }}
                                      >
                                        <Globe size={11} /> {item.externalSharesCount} Extern gedeeld
                                      </span>
                                    </div>
                                  ) : (
                                    <span style={{ fontSize: '0.75rem', color: 'var(--color-slate, #6B7280)' }}>
                                      Geen externen
                                    </span>
                                  )}
                                </td>
                                <td>
                                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                    {item.permissions.length === 0 ? (
                                      <span style={{ fontSize: '0.71875rem', color: 'var(--color-slate, #6B7280)' }}>
                                        Standaard bibliotheekmachtigingen
                                      </span>
                                    ) : (
                                      item.permissions.map((p) => (
                                        <div
                                          key={p.id}
                                          style={{
                                            fontSize: '0.71875rem',
                                            display: 'flex',
                                            alignItems: 'center',
                                            gap: '6px',
                                          }}
                                        >
                                          {p.isExternal ? (
                                            <Globe size={11} color="#0099B8" />
                                          ) : p.userType === 'sharingLink' ? (
                                            <Link2 size={11} color="#2060E0" />
                                          ) : (
                                            <Users size={11} color="#6B7280" />
                                          )}
                                          <span style={{ fontWeight: 500, color: 'var(--text-heading, #0F2A66)' }}>
                                            {p.displayName}
                                          </span>
                                          <span className="badge" style={{ fontSize: '0.59375rem', padding: '1px 5px' }}>
                                            {p.role}
                                          </span>
                                          {p.sharingLink && (
                                            <span style={{ fontSize: '0.625rem', color: 'var(--color-royal, #2060E0)' }}>
                                              ({p.sharingLink.scope})
                                            </span>
                                          )}
                                        </div>
                                      ))
                                    )}
                                  </div>
                                </td>
                                <td>
                                  {item.webUrl && (
                                    <a
                                      href={item.webUrl}
                                      target="_blank"
                                      rel="noreferrer"
                                      className="btn btn-secondary"
                                      style={{
                                        padding: '4px 8px',
                                        fontSize: '0.71875rem',
                                        display: 'inline-flex',
                                        alignItems: 'center',
                                        gap: '4px',
                                        textDecoration: 'none',
                                      }}
                                    >
                                      Openen <ExternalLink size={10} />
                                    </a>
                                  )}
                                </td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              )}
            </>
          )}
        </div>

        {/* Footer */}
        <div
          style={{
            padding: '14px 24px',
            borderTop: '1px solid var(--border, #E4E9F2)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            backgroundColor: 'var(--bg-primary, #F1F5FB)',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '16px', fontSize: '0.75rem', color: 'var(--color-slate, #6B7280)' }}>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: '5px' }}>
              <Lock size={12} color="#10B981" /> Overgenomen = Rechten geërfd van bovenliggende site
            </span>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: '5px' }}>
              <Unlock size={12} color="#D97706" /> Unieke rechten = Overerving verbroken op container/item
            </span>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: '5px' }}>
              <Globe size={12} color="#0099B8" /> Extern = Gastaccount (#EXT#) of deellink
            </span>
          </div>

          <button type="button" className="btn btn-secondary" onClick={onClose}>
            Sluiten
          </button>
        </div>
      </div>
    </div>
  );
}
