'use client';

import React, { useState, useEffect, useMemo } from 'react';
import {
  X,
  Search,
  Globe,
  ExternalLink,
  Shield,
  RefreshCw,
  Folder,
  Users,
  Filter,
  Building,
  CheckCircle2,
  AlertCircle,
  ChevronRight,
  Mail,
} from 'lucide-react';
import {
  TenantExternalAuditResult,
  ExternalGuestAccessRecord,
  GraphSite,
} from '@/lib/graph/client';

interface TenantExternalAuditModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSelectSiteToInspect?: (site: GraphSite) => void;
}

import { useMsal } from '@azure/msal-react';
import { fetchTenantExternalAudit } from '@/lib/graph/clientService';

export function TenantExternalAuditModal({
  isOpen,
  onClose,
  onSelectSiteToInspect,
}: TenantExternalAuditModalProps) {
  const { instance, accounts } = useMsal();
  
  const [auditData, setAuditData] = useState<TenantExternalAuditResult | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isLiveApi, setIsLiveApi] = useState(false);

  const [searchQuery, setSearchQuery] = useState('');
  const [onlyWithAccess, setOnlyWithAccess] = useState(false);
  const [selectedGuest, setSelectedGuest] = useState<ExternalGuestAccessRecord | null>(null);

  useEffect(() => {
    if (isOpen) {
      loadExternalAudit();
      setSearchQuery('');
      setOnlyWithAccess(false);
      setSelectedGuest(null);
    }
  }, [isOpen]);

  const loadExternalAudit = async () => {
    setIsLoading(true);
    setError(null);
    try {
      if (accounts.length > 0) {
        setIsLiveApi(true);
      }
      const data = await fetchTenantExternalAudit(instance, accounts);
      setAuditData(data as any);
    } catch (err: any) {
      console.error(err);
      setError(err.message || 'Fout bij het ophalen van externe gasten.');
    } finally {
      setIsLoading(false);
    }
  };

  const filteredGuests = useMemo(() => {
    if (!auditData?.guests) return [];
    return auditData.guests.filter((g) => {
      if (onlyWithAccess && g.sharedSitesCount === 0) return false;
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const matchName = g.displayName.toLowerCase().includes(q);
        const matchEmail = g.email.toLowerCase().includes(q);
        const matchDomain = g.companyOrDomain.toLowerCase().includes(q);
        const matchSite = g.sharedSites.some(
          (s) =>
            s.siteTitle.toLowerCase().includes(q) ||
            s.siteUrl.toLowerCase().includes(q) ||
            (s.groupName && s.groupName.toLowerCase().includes(q))
        );
        return matchName || matchEmail || matchDomain || matchSite;
      }
      return true;
    });
  }, [auditData?.guests, onlyWithAccess, searchQuery]);

  if (!isOpen) return null;

  return (
    <div
      style={{
        position: 'fixed',
        inset: 0,
        backgroundColor: 'rgba(15, 42, 102, 0.6)',
        backdropFilter: 'blur(8px)',
        zIndex: 1150,
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
          maxWidth: '1180px',
          maxHeight: '92vh',
          display: 'flex',
          flexDirection: 'column',
          backgroundColor: 'var(--bg-card, #FFFFFF)',
          borderRadius: '16px',
          boxShadow: '0 25px 50px -12px rgba(15, 42, 102, 0.3)',
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
                backgroundColor: 'rgba(0, 208, 240, 0.12)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                border: '1px solid rgba(0, 208, 240, 0.3)',
              }}
            >
              <Globe size={24} color="#0099B8" />
            </div>
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <h2
                  style={{
                    margin: 0,
                    fontSize: '1.25rem',
                    fontWeight: 700,
                    color: 'var(--text-heading, #0F2A66)',
                    letterSpacing: '-0.02em',
                  }}
                >
                  Externe Gasten & Gedeelde Sites Explorer
                </h2>
                <span
                  className={`badge ${isLiveApi ? 'badge-cyan' : 'badge-primary'}`}
                  style={{ fontSize: '0.6875rem', padding: '2px 8px' }}
                >
                  {isLiveApi ? '🟢 Live Microsoft Graph' : '⚡ Demo Data'}
                </span>
              </div>
              <div style={{ fontSize: '0.8125rem', color: 'var(--color-slate, #6B7280)', marginTop: '2px' }}>
                Zoek en vind alle externe personen en gastaccounts en ontdek exact tot welke SharePoint- en Teams-locaties zij toegang hebben.
              </div>
            </div>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <button
              type="button"
              className="btn btn-secondary"
              onClick={loadExternalAudit}
              disabled={isLoading}
              style={{ padding: '6px 12px', fontSize: '0.75rem', display: 'flex', alignItems: 'center', gap: '6px' }}
            >
              <RefreshCw size={13} className={isLoading ? 'spin' : ''} />
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
              }}
              aria-label="Sluiten"
            >
              <X size={20} />
            </button>
          </div>
        </div>

        {/* Stats Ribbon */}
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
              padding: '12px 16px',
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
                backgroundColor: 'rgba(0, 208, 240, 0.12)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
              }}
            >
              <Users size={18} color="#0099B8" />
            </div>
            <div>
              <div style={{ fontSize: '0.6875rem', fontWeight: 600, color: 'var(--color-slate, #6B7280)', textTransform: 'uppercase' }}>
                Externe Gasten
              </div>
              <div style={{ fontSize: '1.25rem', fontWeight: 700, color: 'var(--text-heading, #0F2A66)', fontVariantNumeric: 'tabular-nums' }}>
                {auditData?.totalGuests ?? '-'}
              </div>
            </div>
          </div>

          <div
            style={{
              backgroundColor: 'var(--bg-card, #FFFFFF)',
              padding: '12px 16px',
              borderRadius: '10px',
              border: '1.5px solid rgba(0, 208, 240, 0.4)',
              display: 'flex',
              alignItems: 'center',
              gap: '12px',
              position: 'relative',
              overflow: 'hidden',
            }}
          >
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
            <div
              style={{
                width: '36px',
                height: '36px',
                borderRadius: '8px',
                backgroundColor: 'rgba(0, 208, 240, 0.15)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
              }}
            >
              <CheckCircle2 size={18} color="#0099B8" />
            </div>
            <div>
              <div style={{ fontSize: '0.6875rem', fontWeight: 600, color: 'var(--color-slate, #6B7280)', textTransform: 'uppercase' }}>
                Actief op Sites
              </div>
              <div style={{ fontSize: '1.25rem', fontWeight: 700, color: '#007A93', fontVariantNumeric: 'tabular-nums' }}>
                {auditData?.guestsWithSiteAccess ?? '-'}
              </div>
            </div>
          </div>

          <div
            style={{
              backgroundColor: 'var(--bg-card, #FFFFFF)',
              padding: '12px 16px',
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
              <Folder size={18} color="var(--color-royal, #2060E0)" />
            </div>
            <div>
              <div style={{ fontSize: '0.6875rem', fontWeight: 600, color: 'var(--color-slate, #6B7280)', textTransform: 'uppercase' }}>
                Gedeelde Sites
              </div>
              <div style={{ fontSize: '1.25rem', fontWeight: 700, color: 'var(--text-heading, #0F2A66)', fontVariantNumeric: 'tabular-nums' }}>
                {auditData?.totalSharedSites ?? '-'}
              </div>
            </div>
          </div>

          <div
            style={{
              backgroundColor: 'var(--bg-card, #FFFFFF)',
              padding: '12px 16px',
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
                backgroundColor: 'rgba(107, 114, 128, 0.1)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
              }}
            >
              <Building size={18} color="#4B5563" />
            </div>
            <div>
              <div style={{ fontSize: '0.6875rem', fontWeight: 600, color: 'var(--color-slate, #6B7280)', textTransform: 'uppercase' }}>
                Externe Domeinen
              </div>
              <div style={{ fontSize: '1.25rem', fontWeight: 700, color: 'var(--text-heading, #0F2A66)', fontVariantNumeric: 'tabular-nums' }}>
                {auditData?.uniqueExternalDomains ?? '-'}
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
            alignItems: 'center',
            justifyContent: 'space-between',
            gap: '12px',
          }}
        >
          <div style={{ position: 'relative', flex: 1, maxWidth: '480px' }}>
            <Search
              size={15}
              style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)', color: '#6B7280' }}
            />
            <input
              type="text"
              className="portal-input"
              placeholder="Zoek op gastnaam, e-mail, domein of sitenaam..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              style={{ paddingLeft: '36px', fontSize: '0.8125rem', height: '36px', width: '100%' }}
            />
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <button
              type="button"
              onClick={() => setOnlyWithAccess(!onlyWithAccess)}
              style={{
                height: '36px',
                padding: '0 14px',
                borderRadius: '8px',
                fontSize: '0.75rem',
                fontWeight: 600,
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                gap: '6px',
                border: onlyWithAccess ? '1.5px solid #00D0F0' : '1px solid var(--border, #E4E9F2)',
                backgroundColor: onlyWithAccess ? 'rgba(0, 208, 240, 0.12)' : 'var(--bg-primary, #F1F5FB)',
                color: onlyWithAccess ? '#007A93' : 'var(--text-main, #1A2440)',
                transition: 'all 0.15s ease',
              }}
            >
              <Filter size={13} color={onlyWithAccess ? '#0099B8' : '#6B7280'} />
              <span>Alleen met Actieve Sitetoegang</span>
              {onlyWithAccess && (
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
                  {auditData?.guestsWithSiteAccess || 0}
                </span>
              )}
            </button>
          </div>
        </div>

        {/* Content */}
        <div style={{ flex: 1, overflowY: 'auto', padding: '20px 24px' }}>
          {isLoading ? (
            <div style={{ padding: '60px 0', textAlign: 'center', color: 'var(--color-slate, #6B7280)' }}>
              <RefreshCw size={28} className="spin" style={{ margin: '0 auto 12px auto', color: 'var(--color-royal, #2060E0)' }} />
              <div style={{ fontWeight: 600, color: 'var(--text-heading, #0F2A66)' }}>
                Tenant-brede scan van externe gasten en gedeelde sites uitvoeren...
              </div>
              <div style={{ fontSize: '0.8125rem' }}>Analyseren van Entra ID gasten, lidmaatschappen en SharePoint sites...</div>
            </div>
          ) : error ? (
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
              <AlertCircle size={18} />
              <span>{error}</span>
            </div>
          ) : (
            <div>
              <div style={{ marginBottom: '12px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span style={{ fontSize: '0.8125rem', color: 'var(--color-slate, #6B7280)' }}>
                  Gevonden externe gasten: <strong style={{ color: 'var(--text-heading, #0F2A66)' }}>{filteredGuests.length}</strong>
                </span>
                {searchQuery && (
                  <button
                    type="button"
                    onClick={() => setSearchQuery('')}
                    style={{ fontSize: '0.75rem', color: 'var(--color-royal, #2060E0)', background: 'none', border: 'none', cursor: 'pointer' }}
                  >
                    Zoekfilter wissen
                  </button>
                )}
              </div>

              {filteredGuests.length === 0 ? (
                <div style={{ textAlign: 'center', padding: '50px 0', color: 'var(--color-slate, #6B7280)' }}>
                  <Globe size={32} style={{ margin: '0 auto 10px auto', opacity: 0.5 }} />
                  <div style={{ fontWeight: 600 }}>Geen externe gasten gevonden die voldoen aan de zoekcriteria.</div>
                  <div style={{ fontSize: '0.8125rem', marginTop: '4px' }}>
                    Controleer de zoekterm of schakel het filter "Alleen met Actieve Sitetoegang" uit.
                  </div>
                </div>
              ) : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                  {filteredGuests.map((guest) => {
                    const isSelected = selectedGuest?.id === guest.id;
                    const hasAccess = guest.sharedSitesCount > 0;

                    return (
                      <div
                        key={guest.id}
                        style={{
                          backgroundColor: 'var(--bg-card, #FFFFFF)',
                          borderRadius: '12px',
                          border: hasAccess ? '1px solid rgba(0, 208, 240, 0.4)' : '1px solid var(--border, #E4E9F2)',
                          boxShadow: '0 2px 6px rgba(15, 42, 102, 0.04)',
                          overflow: 'hidden',
                          transition: 'all 0.15s ease',
                        }}
                      >
                        <div
                          style={{
                            padding: '14px 18px',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'space-between',
                            background: hasAccess
                              ? 'linear-gradient(90deg, rgba(0, 208, 240, 0.04) 0%, transparent 100%)'
                              : undefined,
                          }}
                        >
                          <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
                            <div
                              style={{
                                width: '38px',
                                height: '38px',
                                borderRadius: '10px',
                                backgroundColor: hasAccess ? 'rgba(0, 208, 240, 0.15)' : 'rgba(107, 114, 128, 0.1)',
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                border: hasAccess ? '1px solid rgba(0, 208, 240, 0.3)' : '1px solid var(--border, #E4E9F2)',
                              }}
                            >
                              <Globe size={18} color={hasAccess ? '#0099B8' : '#6B7280'} />
                            </div>

                            <div>
                              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <span style={{ fontWeight: 700, fontSize: '0.9375rem', color: 'var(--text-heading, #0F2A66)' }}>
                                  {guest.displayName}
                                </span>
                                <span
                                  className="badge"
                                  style={{
                                    fontSize: '0.625rem',
                                    backgroundColor: 'var(--bg-primary, #F1F5FB)',
                                    color: 'var(--text-main, #1A2440)',
                                  }}
                                >
                                  {guest.companyOrDomain}
                                </span>
                                {hasAccess ? (
                                  <span className="badge badge-cyan" style={{ fontSize: '0.625rem', padding: '1px 6px' }}>
                                    {guest.sharedSitesCount} Gedeelde Site{guest.sharedSitesCount > 1 ? 's' : ''}
                                  </span>
                                ) : (
                                  <span style={{ fontSize: '0.6875rem', color: 'var(--color-slate, #6B7280)' }}>
                                    (Geen actieve sitetoegang)
                                  </span>
                                )}
                              </div>

                              <div style={{ fontSize: '0.75rem', color: 'var(--color-slate, #6B7280)', display: 'flex', alignItems: 'center', gap: '6px', marginTop: '2px' }}>
                                <Mail size={12} />
                                <span>{guest.email}</span>
                                {guest.userPrincipalName && guest.userPrincipalName !== guest.email && (
                                  <>
                                    <span>•</span>
                                    <span style={{ fontSize: '0.6875rem', opacity: 0.8 }}>UPN: {guest.userPrincipalName}</span>
                                  </>
                                )}
                              </div>
                            </div>
                          </div>

                          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            {hasAccess && (
                              <button
                                type="button"
                                className="btn btn-secondary"
                                onClick={() => setSelectedGuest(isSelected ? null : guest)}
                                style={{ padding: '6px 12px', fontSize: '0.75rem', display: 'flex', alignItems: 'center', gap: '6px' }}
                              >
                                <span>{isSelected ? 'Verberg Sites' : `Bekijk Sites (${guest.sharedSitesCount})`}</span>
                                <ChevronRight size={12} style={{ transform: isSelected ? 'rotate(90deg)' : 'none', transition: 'transform 0.15s ease' }} />
                              </button>
                            )}
                          </div>
                        </div>

                        {/* DETAIL: GEDEELDE SITES */}
                        {(isSelected || (hasAccess && searchQuery.trim())) && (
                          <div
                            style={{
                              borderTop: '1px solid var(--border, #E4E9F2)',
                              padding: '12px 18px',
                              backgroundColor: 'var(--bg-primary, #F1F5FB)',
                            }}
                          >
                            <div style={{ fontSize: '0.75rem', fontWeight: 600, color: 'var(--text-heading, #0F2A66)', marginBottom: '8px' }}>
                              SharePoint & Teams locaties gedeeld met {guest.displayName}:
                            </div>

                            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))', gap: '8px' }}>
                              {guest.sharedSites.map((siteRef, idx) => (
                                <div
                                  key={idx}
                                  style={{
                                    backgroundColor: 'var(--bg-card, #FFFFFF)',
                                    borderRadius: '8px',
                                    border: '1px solid var(--border, #E4E9F2)',
                                    padding: '10px 14px',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'space-between',
                                    gap: '10px',
                                  }}
                                >
                                  <div>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                      <Folder size={14} color="var(--color-royal, #2060E0)" />
                                      <span style={{ fontWeight: 600, fontSize: '0.8125rem', color: 'var(--text-heading, #0F2A66)' }}>
                                        {siteRef.siteTitle}
                                      </span>
                                    </div>
                                    <div style={{ fontSize: '0.6875rem', color: 'var(--color-slate, #6B7280)', marginTop: '2px' }}>
                                      {siteRef.groupName ? `Via Team: ${siteRef.groupName}` : siteRef.role}
                                    </div>
                                  </div>

                                  <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                    {siteRef.siteUrl && (
                                      <a
                                        href={siteRef.siteUrl}
                                        target="_blank"
                                        rel="noreferrer"
                                        className="btn btn-secondary"
                                        style={{ padding: '4px 8px', fontSize: '0.6875rem', display: 'inline-flex', alignItems: 'center', gap: '4px', textDecoration: 'none' }}
                                        title="Open SharePoint site"
                                      >
                                        Openen <ExternalLink size={10} />
                                      </a>
                                    )}

                                    {onSelectSiteToInspect && (
                                      <button
                                        type="button"
                                        className="btn btn-primary"
                                        onClick={() => {
                                          onClose();
                                          onSelectSiteToInspect({
                                            id: siteRef.siteId,
                                            title: siteRef.siteTitle,
                                            url: siteRef.siteUrl,
                                            template: 'Team',
                                            storageQuotaGb: 512,
                                            primaryAdmin: 'Beheerder',
                                            createdAt: new Date().toISOString(),
                                          });
                                        }}
                                        style={{ padding: '4px 8px', fontSize: '0.6875rem', display: 'inline-flex', alignItems: 'center', gap: '4px' }}
                                        title="Inspecteer volledige site-, bibliotheek- en itemrechten"
                                      >
                                        <Shield size={10} /> Inspecteren
                                      </button>
                                    )}
                                  </div>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
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
          <div style={{ fontSize: '0.75rem', color: 'var(--color-slate, #6B7280)', display: 'flex', alignItems: 'center', gap: '16px' }}>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: '5px' }}>
              <Globe size={12} color="#0099B8" /> Externe Gast = Entra ID account met #EXT# of gastlicentie
            </span>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: '5px' }}>
              <Shield size={12} color="var(--color-royal, #2060E0)" /> Klik op 'Inspecteren' voor 4-niveau permissie audit
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
