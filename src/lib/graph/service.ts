import { TenantConnectionConfig, getTenantConnection } from '@/lib/db/store';
import { getGraphAccessToken } from '@/lib/graph/auth';
import {
  GraphSite,
  SharePointPermissionLevel,
  SharePointAccessEntity,
  SharePointContainer,
  SharePointItemPermissionSummary,
  SiteAccessAuditSummary,
  ExternalGuestAccessRecord,
  TenantExternalAuditResult,
} from '@/lib/graph/client';

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

export async function callGraph(
  token: string,
  endpoint: string,
  options: RequestInit = {}
): Promise<any> {
  const url = endpoint.startsWith('http') ? endpoint : `${GRAPH_BASE_URL}${endpoint}`;
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      Accept: 'application/json',
      ...options.headers,
    },
  });

  if (!res.ok) {
    let errorDetail = res.statusText;
    try {
      const errJson = await res.json();
      errorDetail = errJson.error?.message || errJson.message || res.statusText;
    } catch {
      // ignore
    }
    throw new Error(`Graph API Fout [HTTP ${res.status}]: ${errorDetail}`);
  }

  if (res.status === 204) return null;
  return res.json();
}

/**
 * Haalt alle actieve SharePoint & Teamsites op via Graph API
 */
export async function fetchGraphSites(config: TenantConnectionConfig): Promise<GraphSite[]> {
  const token = await getGraphAccessToken(config);
  const data = await callGraph(token, '/sites?search=*&$select=id,name,displayName,webUrl,createdDateTime,siteCollection');
  const rawSites = data.value || [];

  return rawSites.map((s: any) => {
    const isTeam =
      s.webUrl?.toLowerCase().includes('/teams/') ||
      s.webUrl?.toLowerCase().includes('/sites/group') ||
      s.name?.startsWith('group-');

    return {
      id: s.id,
      title: s.displayName || s.name,
      name: s.name,
      displayName: s.displayName || s.name,
      url: s.webUrl || '',
      webUrl: s.webUrl || '',
      template: isTeam ? 'Team' : 'Communication',
      createdDateTime: s.createdDateTime,
      isTeam,
      externalSharingCapability: s.siteCollection?.externalSharingCapability || 'Disabled',
    };
  });
}

/**
 * Los een group-id of site-naam op naar een canonical SharePoint site-id
 */
export async function resolveActualSiteId(token: string, siteOrGroupId: string): Promise<string> {
  if (siteOrGroupId.includes(',')) {
    return siteOrGroupId;
  }

  try {
    const siteData = await callGraph(token, `/sites/${siteOrGroupId}`);
    if (siteData && siteData.id) return siteData.id;
  } catch {
    // Probeer via groups endpoint
  }

  try {
    const groupSiteData = await callGraph(token, `/groups/${siteOrGroupId}/sites/root`);
    if (groupSiteData && groupSiteData.id) return groupSiteData.id;
  } catch {
    // ignore
  }

  return siteOrGroupId;
}

/**
 * Converteer een raw Graph permission entry naar een gestandaardiseerde SharePointAccessEntity
 */
export function normalizeGraphPermission(
  p: any,
  containerType: 'site' | 'library' | 'list' | 'item' = 'site'
): SharePointAccessEntity {
  let userType: SharePointAccessEntity['userType'] = 'user';
  let displayName = 'Onbekend';
  let email = '';
  let principalName = '';
  let isExternal = false;
  let permissionLevel: SharePointPermissionLevel = 'Read';

  // Rollen bepalen
  const roles: string[] = p.roles || [];
  if (roles.includes('owner') || roles.includes('write') && roles.includes('manage')) {
    permissionLevel = 'Full Control';
  } else if (roles.includes('write')) {
    permissionLevel = 'Contribute';
  } else if (roles.includes('read')) {
    permissionLevel = 'Read';
  }

  // Identiteit analyseren
  if (p.grantedToIdentities && p.grantedToIdentities.length > 0) {
    const idObj = p.grantedToIdentities[0];
    if (idObj.user) {
      userType = 'user';
      displayName = idObj.user.displayName || idObj.user.email || 'Gebruiker';
      email = idObj.user.email || '';
      principalName = idObj.user.userPrincipalName || email;
    } else if (idObj.application) {
      userType = 'application';
      displayName = idObj.application.displayName || 'Applicatie';
    } else if (idObj.siteGroup) {
      userType = 'sharePointGroup';
      displayName = idObj.siteGroup.displayName || 'SharePoint Groep';
    }
  } else if (p.grantedTo) {
    if (p.grantedTo.user) {
      userType = 'user';
      displayName = p.grantedTo.user.displayName || p.grantedTo.user.email || 'Gebruiker';
      email = p.grantedTo.user.email || '';
      principalName = p.grantedTo.user.userPrincipalName || email;
    } else if (p.grantedTo.application) {
      userType = 'application';
      displayName = p.grantedTo.application.displayName || 'Applicatie';
    } else if (p.grantedTo.siteGroup) {
      userType = 'sharePointGroup';
      displayName = p.grantedTo.siteGroup.displayName || 'SharePoint Groep';
    }
  } else if (p.link) {
    userType = 'sharingLink';
    displayName = p.link.type === 'anonymous' ? 'Anonieme Gastkoppeling (Iedereen)' : 'Organisatie Deellink';
    principalName = p.link.webUrl || '';
    if (p.link.type === 'anonymous') {
      isExternal = true;
    }
  } else if (p.invitation) {
    userType = 'externalGuest';
    displayName = p.invitation.email || 'Uitgenodigde Gast';
    email = p.invitation.email || '';
    isExternal = true;
  }

  // Detecteer extern gastaccount
  if (
    principalName.includes('#EXT#') ||
    email.includes('#EXT#') ||
    p.userType === 'Guest' ||
    userType === 'externalGuest'
  ) {
    isExternal = true;
    userType = 'externalGuest';
  }

  return {
    id: p.id || Math.random().toString(),
    displayName,
    email,
    principalName,
    role: roles.join(', ') || permissionLevel,
    permissionLevel,
    isExternal,
    userType,
    inherited: Boolean(p.inheritedFrom),
    inheritedFrom: p.inheritedFrom?.displayName || p.inheritedFrom?.webUrl,
    expirationDateTime: p.expirationDateTime,
    hasPassword: Boolean(p.hasPassword),
  };
}

/**
 * Haalt site-niveau permissies op via /sites/{id}/permissions
 */
export async function fetchGraphSitePermissions(
  config: TenantConnectionConfig,
  siteId: string
): Promise<SharePointAccessEntity[]> {
  const token = await getGraphAccessToken(config);
  const resolvedId = await resolveActualSiteId(token, siteId);
  const data = await callGraph(token, `/sites/${resolvedId}/permissions`);
  const rawList = data.value || [];

  const entities: SharePointAccessEntity[] = rawList.map((p: any) =>
    normalizeGraphPermission(p, 'site')
  );

  // Indien het een M365 Unified Groep/Team site betreft, haal ook de groepsleden op
  try {
    const siteObj = await callGraph(token, `/sites/${resolvedId}`);
    if (siteObj && siteObj.webUrl) {
      const urlParts = siteObj.webUrl.split('/');
      const sitePath = urlParts[urlParts.length - 1];

      const grpData = await callGraph(
        token,
        `/groups?$filter=mailNickname eq '${sitePath}'&$select=id,displayName`
      );
      if (grpData.value && grpData.value.length > 0) {
        const groupId = grpData.value[0].id;
        const membersData = await callGraph(
          token,
          `/groups/${groupId}/members?$select=id,displayName,mail,userPrincipalName,userType`
        );
        const members = membersData.value || [];

        for (const m of members) {
          const isGuest = m.userType === 'Guest' || (m.userPrincipalName && m.userPrincipalName.includes('#EXT#'));
          entities.push({
            id: m.id,
            displayName: m.displayName || m.mail || 'Teamlid',
            email: m.mail || m.userPrincipalName || '',
            principalName: m.userPrincipalName || '',
            role: isGuest ? 'Gastlid (Team)' : 'Teamlid (M365)',
            permissionLevel: 'Contribute',
            isExternal: isGuest,
            userType: isGuest ? 'externalGuest' : 'user',
            inherited: true,
            inheritedFrom: 'Microsoft Teams / M365 Group',
          });
        }
      }
    }
  } catch (grpErr) {
    console.warn('[fetchGraphSitePermissions] Kon Teams groepsleden niet ophalen:', grpErr);
  }

  return entities;
}

/**
 * Haalt containers (Documentbibliotheken en Lijsten) van een site op
 */
export async function fetchGraphSiteContainers(
  config: TenantConnectionConfig,
  siteId: string
): Promise<SharePointContainer[]> {
  const token = await getGraphAccessToken(config);
  const resolvedId = await resolveActualSiteId(token, siteId);

  // 1. Documentbibliotheken (Drives)
  const drivesData = await callGraph(token, `/sites/${resolvedId}/drives`);
  const drives = drivesData.value || [];

  // 2. Lijsten (Lists)
  const listsData = await callGraph(token, `/sites/${resolvedId}/lists?$filter=hidden eq false`);
  const lists = listsData.value || [];

  const containers: SharePointContainer[] = [];

  for (const d of drives) {
    let drivePerms: SharePointAccessEntity[] = [];
    let hasUnique = false;
    let externalCount = 0;

    try {
      const pData = await callGraph(token, `/drives/${d.id}/root/permissions`);
      drivePerms = (pData.value || []).map((p: any) => normalizeGraphPermission(p, 'library'));
      hasUnique = drivePerms.some((p) => !p.inherited || p.userType === 'sharingLink');
      externalCount = drivePerms.filter((p) => p.isExternal).length;
    } catch {
      // ignore
    }

    containers.push({
      id: d.id,
      name: d.name,
      displayName: d.name,
      type: 'documentLibrary',
      webUrl: d.webUrl,
      itemCount: d.quota?.used || 0,
      hasUniquePermissions: hasUnique,
      externalPermissionsCount: externalCount,
      permissions: drivePerms,
    });
  }

  for (const l of lists) {
    if (containers.some((c) => c.name === l.name)) continue;

    containers.push({
      id: l.id,
      name: l.name,
      displayName: l.displayName || l.name,
      type: 'list',
      webUrl: l.webUrl,
      itemCount: l.list?.itemCount || 0,
      hasUniquePermissions: false,
      externalPermissionsCount: 0,
    });
  }

  return containers;
}

/**
 * Haalt items van een bibliotheek op inclusief unieke permissies en deellinks
 */
export async function fetchGraphContainerItemPermissions(
  config: TenantConnectionConfig,
  siteId: string,
  driveId: string,
  folderId?: string
): Promise<SharePointItemPermissionSummary[]> {
  const token = await getGraphAccessToken(config);
  const endpoint = folderId
    ? `/drives/${driveId}/items/${folderId}/children?$select=id,name,file,folder,webUrl&$top=50`
    : `/drives/${driveId}/root/children?$select=id,name,file,folder,webUrl&$top=50`;

  const childrenData = await callGraph(token, endpoint);
  const items = childrenData.value || [];
  const results: SharePointItemPermissionSummary[] = [];

  for (const item of items) {
    let perms: SharePointAccessEntity[] = [];
    let hasUnique = false;
    let externalCount = 0;

    try {
      const pData = await callGraph(token, `/drives/${driveId}/items/${item.id}/permissions`);
      perms = (pData.value || []).map((p: any) => normalizeGraphPermission(p, 'item'));
      hasUnique = perms.some((p) => !p.inherited || p.userType === 'sharingLink');
      externalCount = perms.filter((p) => p.isExternal).length;
    } catch (itemErr) {
      console.warn(`[fetchGraphContainerItemPermissions] Permissies ophalen voor item ${item.name} mislukt:`, itemErr);
    }

    results.push({
      id: item.id,
      name: item.name,
      isFolder: !!item.folder,
      webUrl: item.webUrl,
      hasUniquePermissions: hasUnique,
      externalSharesCount: externalCount,
      permissions: perms,
    });
  }

  return results;
}

/**
 * Genereert een complete audit-samenvatting van site- en containerrechten
 */
export async function fetchCompleteSiteAccessAudit(
  config: TenantConnectionConfig,
  siteId: string
): Promise<SiteAccessAuditSummary> {
  const [sitePerms, containers] = await Promise.all([
    fetchGraphSitePermissions(config, siteId),
    fetchGraphSiteContainers(config, siteId),
  ]);

  let totalPermissions = sitePerms.length;
  let externalCount = sitePerms.filter((p) => p.isExternal).length;
  let sharingLinksCount = sitePerms.filter((p) => p.userType === 'sharingLink').length;
  let uniquePermissionItemsCount = containers.filter((c) => c.hasUniquePermissions).length;

  for (const c of containers) {
    if (c.permissions) {
      totalPermissions += c.permissions.length;
      externalCount += c.externalPermissionsCount || 0;
      sharingLinksCount += c.permissions.filter((p) => p.userType === 'sharingLink').length;
    }
  }

  return {
    siteId,
    siteTitle: 'Site',
    totalPermissions,
    externalCount,
    sharingLinksCount,
    uniquePermissionItemsCount,
    sitePermissions: sitePerms,
    containers,
  };
}

/**
 * Voert een tenant-brede audit uit op alle externe gasten en gedeelde sites/teams
 */
export async function fetchTenantExternalAccessAudit(
  config?: TenantConnectionConfig | null,
  query?: string
): Promise<TenantExternalAuditResult> {
  const activeConfig = config || getTenantConnection();
  if (!activeConfig || !activeConfig.tenantId || !activeConfig.clientId || !activeConfig.clientSecret) {
    return getDemoTenantExternalAudit(query);
  }

  try {
    const token = await getGraphAccessToken(activeConfig);

    // 1. Haal alle Unified groepen op en map hun SharePoint sites
    const grpRes = await callGraph(
      token,
      "/groups?$filter=groupTypes/any(c:c eq 'Unified')&$select=id,displayName,mailNickname&$top=100"
    );
    const groups = grpRes.value || [];
    const groupSiteMap = new Map<string, { groupId: string; groupName: string; siteId?: string }>();

    for (const g of groups) {
      groupSiteMap.set(g.id, {
        groupId: g.id,
        groupName: g.displayName,
      });
    }

    // 2. Haal alle Gastgebruikers op (userType eq 'Guest')
    const guestsRes = await callGraph(
      token,
      "/users?$filter=userType eq 'Guest'&$select=id,displayName,mail,userPrincipalName,companyName,createdDateTime&$top=200"
    );
    const guests = guestsRes.value || [];

    // 3. Voor elke gast: controleer lidmaatschap van groepen/teams
    const guestRecords: ExternalGuestAccessRecord[] = await Promise.all(
      guests.map(async (g: any) => {
        const email = g.mail || g.userPrincipalName || '';
        let domain = '';
        if (email.includes('@')) {
          domain = email.split('@')[1];
        }

        const sharedSites: ExternalGuestAccessRecord['sharedSites'] = [];

        try {
          const memberships = await callGraph(
            token,
            `/users/${g.id}/transitiveMemberOf?$select=id,displayName`
          );
          const memberGroups = memberships.value || [];

          for (const mg of memberGroups) {
            if (groupSiteMap.has(mg.id)) {
              const info = groupSiteMap.get(mg.id)!;
              sharedSites.push({
                siteId: info.siteId || mg.id,
                siteTitle: info.groupName,
                siteUrl: '',
                groupName: info.groupName,
                accessType: 'groupMember',
                role: 'Gast / Deelnemer',
              });
            }
          }
        } catch {
          // Geen lidmaatschappen kunnen lezen
        }

        return {
          id: g.id,
          displayName: g.displayName || email,
          email,
          userPrincipalName: g.userPrincipalName || '',
          companyOrDomain: g.companyName || domain,
          sharedSitesCount: sharedSites.length,
          sharedSites,
        };
      })
    );

    // 4. Filteren op zoekterm indien opgegeven
    let filtered = guestRecords;
    if (query && query.trim()) {
      const q = query.toLowerCase().trim();
      filtered = guestRecords.filter(
        (g) =>
          g.displayName.toLowerCase().includes(q) ||
          g.email.toLowerCase().includes(q) ||
          g.companyOrDomain.toLowerCase().includes(q) ||
          g.sharedSites.some(
            (s) =>
              s.siteTitle.toLowerCase().includes(q) ||
              s.siteUrl.toLowerCase().includes(q) ||
              (s.groupName && s.groupName.toLowerCase().includes(q))
          )
      );
    }

    filtered.sort((a, b) => b.sharedSitesCount - a.sharedSitesCount);

    const guestsWithAccess = guestRecords.filter((g) => g.sharedSitesCount > 0);
    const allSharedSiteIds = new Set<string>();
    const allDomains = new Set<string>();

    for (const g of guestRecords) {
      if (g.companyOrDomain) allDomains.add(g.companyOrDomain.toLowerCase());
      for (const s of g.sharedSites) {
        allSharedSiteIds.add(s.siteId);
      }
    }

    return {
      totalGuests: guestRecords.length,
      guestsWithSiteAccess: guestsWithAccess.length,
      totalSharedSites: allSharedSiteIds.size,
      uniqueExternalDomains: allDomains.size,
      guests: filtered,
    };
  } catch (err) {
    console.warn('[fetchTenantExternalAccessAudit] Fout bij ophalen live audit, terugval op demo data:', err);
    return getDemoTenantExternalAudit(query);
  }
}

// ============================================================
// HIGH LEVEL APP WRAPPERS (With Demo Mode fallback)
// ============================================================

export async function fetchSharePointSites(
  search?: string
): Promise<{ sites: GraphSite[]; isDemo: boolean; error?: string }> {
  const config = getTenantConnection();
  if (!config || !config.tenantId || !config.clientId || !config.clientSecret) {
    const sites = filterDemoSites(search);
    return { sites, isDemo: true };
  }

  try {
    const sites = await fetchGraphSites(config);
    if (search && search.trim()) {
      const q = search.toLowerCase();
      return {
        sites: sites.filter((s) => {
          const title = s.displayName || s.title || s.name || '';
          const name = s.name || s.title || '';
          const url = s.webUrl || s.url || '';
          return title.toLowerCase().includes(q) || name.toLowerCase().includes(q) || url.toLowerCase().includes(q);
        }),
        isDemo: false,
      };
    }
    return { sites, isDemo: false };
  } catch (err: any) {
    console.warn('[fetchSharePointSites] Live ophalen mislukt, terugval naar demo:', err.message);
    const sites = filterDemoSites(search);
    return { sites, isDemo: true, error: err.message };
  }
}

export async function fetchSiteAccessAudit(
  siteId: string
): Promise<{ audit: SiteAccessAuditSummary; isLive: boolean }> {
  const config = getTenantConnection();
  if (!config || !config.tenantId || !config.clientId || !config.clientSecret) {
    return { audit: getDemoSiteAudit(siteId), isLive: false };
  }

  try {
    const audit = await fetchCompleteSiteAccessAudit(config, siteId);
    return { audit, isLive: true };
  } catch (err: any) {
    console.warn(`[fetchSiteAccessAudit] Fout bij live site audit ${siteId}:`, err.message);
    return { audit: getDemoSiteAudit(siteId), isLive: false };
  }
}

export async function fetchContainerItems(
  siteId: string,
  driveId: string
): Promise<{ items: SharePointItemPermissionSummary[]; isLive: boolean }> {
  const config = getTenantConnection();
  if (!config || !config.tenantId || !config.clientId || !config.clientSecret) {
    return { items: getDemoDriveItems(driveId), isLive: false };
  }

  try {
    const items = await fetchGraphContainerItemPermissions(config, siteId, driveId);
    return { items, isLive: true };
  } catch (err: any) {
    console.warn(`[fetchContainerItems] Fout bij ophalen van items voor drive ${driveId}:`, err.message);
    return { items: getDemoDriveItems(driveId), isLive: false };
  }
}

// ============================================================
// REALISTIC DEMO DATA SETS
// ============================================================

const DEMO_SITES: GraphSite[] = [
  {
    id: 'demo-site-1',
    name: 'directie-management',
    displayName: 'Directie & Management',
    title: 'Directie & Management',
    url: 'https://noppa.sharepoint.com/sites/directie-management',
    webUrl: 'https://noppa.sharepoint.com/sites/directie-management',
    template: 'Team',
    isTeam: true,
    createdDateTime: '2025-01-10T09:00:00Z',
    lastModifiedDateTime: '2026-09-28T14:30:00Z',
    externalSharingCapability: 'Disabled',
  },
  {
    id: 'demo-site-2',
    name: 'cloud-migratie-project',
    displayName: 'Project Noppa Cloud Migratie',
    title: 'Project Noppa Cloud Migratie',
    url: 'https://noppa.sharepoint.com/sites/cloud-migratie-project',
    webUrl: 'https://noppa.sharepoint.com/sites/cloud-migratie-project',
    template: 'Team',
    isTeam: true,
    createdDateTime: '2025-04-15T11:20:00Z',
    lastModifiedDateTime: '2026-09-30T10:15:00Z',
    externalSharingCapability: 'ExternalUserSharingOnly',
  },
  {
    id: 'demo-site-3',
    name: 'kennisbank-docs',
    displayName: 'Kennisbank & Organisatie Documenten',
    title: 'Kennisbank & Organisatie Documenten',
    url: 'https://noppa.sharepoint.com/sites/kennisbank-docs',
    webUrl: 'https://noppa.sharepoint.com/sites/kennisbank-docs',
    template: 'Communication',
    isTeam: false,
    createdDateTime: '2024-11-01T08:00:00Z',
    lastModifiedDateTime: '2026-09-29T16:45:00Z',
    externalSharingCapability: 'Disabled',
  },
  {
    id: 'demo-site-4',
    name: 'klantportaal-samenwerking',
    displayName: 'Klantportaal & Partner Samenwerking',
    title: 'Klantportaal & Partner Samenwerking',
    url: 'https://noppa.sharepoint.com/sites/klantportaal-samenwerking',
    webUrl: 'https://noppa.sharepoint.com/sites/klantportaal-samenwerking',
    template: 'Communication',
    isTeam: false,
    createdDateTime: '2025-06-20T13:00:00Z',
    lastModifiedDateTime: '2026-09-30T16:00:00Z',
    externalSharingCapability: 'ExternalUserAndGuestSharing',
  },
  {
    id: 'demo-site-5',
    name: 'hr-personeelszaken',
    displayName: 'HR & Personeelszaken',
    title: 'HR & Personeelszaken',
    url: 'https://noppa.sharepoint.com/sites/hr-personeelszaken',
    webUrl: 'https://noppa.sharepoint.com/sites/hr-personeelszaken',
    template: 'Team',
    isTeam: true,
    createdDateTime: '2025-02-14T10:00:00Z',
    lastModifiedDateTime: '2026-09-25T11:30:00Z',
    externalSharingCapability: 'Disabled',
  },
  {
    id: 'demo-site-6',
    name: 'marketing-communicatie',
    displayName: 'Marketing & Externe Communicatie',
    title: 'Marketing & Externe Communicatie',
    url: 'https://noppa.sharepoint.com/sites/marketing-communicatie',
    webUrl: 'https://noppa.sharepoint.com/sites/marketing-communicatie',
    template: 'Communication',
    isTeam: false,
    createdDateTime: '2025-05-01T09:30:00Z',
    lastModifiedDateTime: '2026-09-29T18:20:00Z',
    externalSharingCapability: 'ExternalUserAndGuestSharing',
  },
];

function filterDemoSites(search?: string): GraphSite[] {
  if (!search || !search.trim()) return DEMO_SITES;
  const q = search.toLowerCase().trim();
  return DEMO_SITES.filter((s) => {
    const title = s.displayName || s.title || s.name || '';
    const name = s.name || s.title || '';
    const url = s.webUrl || s.url || '';
    return title.toLowerCase().includes(q) || name.toLowerCase().includes(q) || url.toLowerCase().includes(q);
  });
}

function getDemoSiteAudit(siteId: string): SiteAccessAuditSummary {
  const site: GraphSite = DEMO_SITES.find((s) => s.id === siteId) || {
    id: siteId,
    title: 'Geselecteerde SharePoint Site',
    url: '',
    name: 'Geselecteerde SharePoint Site',
    displayName: 'Geselecteerde SharePoint Site',
  };

  const sitePerms: SharePointAccessEntity[] = [
    {
      id: 'perm-1',
      displayName: 'Beheerdersgroep Noppa',
      principalName: 'beheerders@noppa.nl',
      email: 'beheerders@noppa.nl',
      role: 'Site Owners (Volledig beheer)',
      permissionLevel: 'Full Control',
      isExternal: false,
      userType: 'sharePointGroup',
      inherited: false,
    },
    {
      id: 'perm-2',
      displayName: 'Rik Dobbelsteen',
      principalName: 'rik@noppa.nl',
      email: 'rik@noppa.nl',
      role: 'Site Owner',
      permissionLevel: 'Full Control',
      isExternal: false,
      userType: 'user',
      inherited: false,
    },
    {
      id: 'perm-3',
      displayName: 'Peter de Vries (Externe Partner BV)',
      principalName: 'peter.de.vries_externe-partner.nl#EXT#@noppa.onmicrosoft.com',
      email: 'peter.de.vries@externe-partner.nl',
      role: 'Gastlid / Medewerker',
      permissionLevel: 'Contribute',
      isExternal: true,
      userType: 'externalGuest',
      inherited: false,
    },
    {
      id: 'perm-4',
      displayName: 'Lisa Jansen (Klant BV)',
      principalName: 'lisa.jansen_klant-bv.nl#EXT#@noppa.onmicrosoft.com',
      email: 'lisa.jansen@klant-bv.nl',
      role: 'Leestoegang (Gast)',
      permissionLevel: 'Read',
      isExternal: true,
      userType: 'externalGuest',
      inherited: false,
    },
    {
      id: 'perm-5',
      displayName: 'Alle interne medewerkers',
      principalName: 'Everyone except external users',
      email: '',
      role: 'Lezers',
      permissionLevel: 'Read',
      isExternal: false,
      userType: 'sharePointGroup',
      inherited: true,
      inheritedFrom: 'Tenant Root Site',
    },
  ];

  const containers: SharePointContainer[] = [
    {
      id: 'drive-demo-docs',
      name: 'Documenten (Gedeelde Documenten)',
      displayName: 'Documenten',
      type: 'documentLibrary',
      webUrl: 'https://noppa.sharepoint.com/sites/shared-docs',
      itemCount: 142,
      hasUniquePermissions: true,
      externalPermissionsCount: 2,
      permissions: [
        {
          id: 'cp-1',
          displayName: 'Peter de Vries (Externe Partner BV)',
          principalName: 'peter.de.vries@externe-partner.nl',
          email: 'peter.de.vries@externe-partner.nl',
          role: 'Bewerken (Contribute)',
          permissionLevel: 'Contribute',
          isExternal: true,
          userType: 'externalGuest',
          inherited: false,
        },
        {
          id: 'cp-2',
          displayName: 'Anonieme Deellink (Alleen Lezen)',
          principalName: 'https://noppa.sharepoint.com/:f:/s/shared-docs/guestlink',
          email: '',
          role: 'Anonieme Gastkoppeling',
          permissionLevel: 'Read',
          isExternal: true,
          userType: 'sharingLink',
          inherited: false,
        },
      ],
    },
    {
      id: 'drive-demo-contracts',
      name: 'Contracten & Overeenkomsten',
      displayName: 'Contracten & Overeenkomsten',
      type: 'documentLibrary',
      webUrl: 'https://noppa.sharepoint.com/sites/contracts',
      itemCount: 38,
      hasUniquePermissions: true,
      externalPermissionsCount: 1,
      permissions: [
        {
          id: 'cp-3',
          displayName: 'Lisa Jansen (Klant BV)',
          principalName: 'lisa.jansen@klant-bv.nl',
          email: 'lisa.jansen@klant-bv.nl',
          role: 'Lezen & Downloaden',
          permissionLevel: 'Read',
          isExternal: true,
          userType: 'externalGuest',
          inherited: false,
        },
      ],
    },
    {
      id: 'list-demo-tasks',
      name: 'Projecttaken & Actiepunten',
      displayName: 'Projecttaken',
      type: 'list',
      webUrl: 'https://noppa.sharepoint.com/sites/lists/tasks',
      itemCount: 65,
      hasUniquePermissions: false,
      externalPermissionsCount: 0,
    },
  ];

  return {
    siteId,
    siteTitle: site.displayName || site.title || site.name || 'Site',
    totalPermissions: 8,
    externalCount: 3,
    sharingLinksCount: 1,
    uniquePermissionItemsCount: 2,
    sitePermissions: sitePerms,
    containers,
  };
}

function getDemoDriveItems(driveId: string): SharePointItemPermissionSummary[] {
  return [
    {
      id: 'item-demo-1',
      name: '2026-Project-Strategie-Noppa.docx',
      isFolder: false,
      webUrl: 'https://noppa.sharepoint.com/item/project-strategie',
      hasUniquePermissions: false,
      externalSharesCount: 0,
      permissions: [
        {
          id: 'ip-1',
          displayName: 'Overgeërfd van Bibliotheek',
          principalName: '',
          email: '',
          role: 'Overgeërfd',
          permissionLevel: 'Read',
          isExternal: false,
          userType: 'user',
          inherited: true,
        },
      ],
    },
    {
      id: 'item-demo-2',
      name: 'Samenwerkingsovereenkomst-Klant-BV.pdf',
      isFolder: false,
      webUrl: 'https://noppa.sharepoint.com/item/samenwerking-klant',
      hasUniquePermissions: true,
      externalSharesCount: 2,
      permissions: [
        {
          id: 'ip-2',
          displayName: 'Lisa Jansen (Klant BV)',
          principalName: 'lisa.jansen@klant-bv.nl',
          email: 'lisa.jansen@klant-bv.nl',
          role: 'Directe Toegang (Bewerken)',
          permissionLevel: 'Contribute',
          isExternal: true,
          userType: 'externalGuest',
          inherited: false,
        },
        {
          id: 'ip-3',
          displayName: 'Gedeelde Link (Klant BV Organisatie)',
          principalName: 'https://noppa.sharepoint.com/:b:/s/shared/link-klant-bv',
          email: '',
          role: 'Bedrijfsbrede Deellink',
          permissionLevel: 'Read',
          isExternal: true,
          userType: 'sharingLink',
          inherited: false,
        },
      ],
    },
    {
      id: 'item-demo-3',
      name: 'Financieel-Overzicht-Q3.xlsx',
      isFolder: false,
      webUrl: 'https://noppa.sharepoint.com/item/financieel-q3',
      hasUniquePermissions: true,
      externalSharesCount: 1,
      permissions: [
        {
          id: 'ip-4',
          displayName: 'Peter de Vries (Externe Partner BV)',
          principalName: 'peter.de.vries@externe-partner.nl',
          email: 'peter.de.vries@externe-partner.nl',
          role: 'Directe Leestoegang',
          permissionLevel: 'Read',
          isExternal: true,
          userType: 'externalGuest',
          inherited: false,
        },
      ],
    },
    {
      id: 'item-demo-4',
      name: 'Presentatiemateriaal-en-Assets',
      isFolder: true,
      webUrl: 'https://noppa.sharepoint.com/folder/assets',
      hasUniquePermissions: false,
      externalSharesCount: 0,
      permissions: [],
    },
  ];
}

function getDemoTenantExternalAudit(query?: string): TenantExternalAuditResult {
  const guests: ExternalGuestAccessRecord[] = [
    {
      id: 'guest-1',
      displayName: 'Peter de Vries',
      email: 'peter.de.vries@externe-partner.nl',
      userPrincipalName: 'peter.de.vries_externe-partner.nl#EXT#@noppa.onmicrosoft.com',
      companyOrDomain: 'externe-partner.nl',
      sharedSitesCount: 2,
      sharedSites: [
        {
          siteId: 'demo-site-2',
          siteTitle: 'Project Noppa Cloud Migratie',
          siteUrl: 'https://noppa.sharepoint.com/sites/cloud-migratie-project',
          groupName: 'Project Noppa Cloud Migratie Team',
          accessType: 'groupMember',
          role: 'Gast / Teamlid',
        },
        {
          siteId: 'demo-site-4',
          siteTitle: 'Klantportaal & Partner Samenwerking',
          siteUrl: 'https://noppa.sharepoint.com/sites/klantportaal-samenwerking',
          accessType: 'directShare',
          role: 'Bewerken (Contribute)',
        },
      ],
    },
    {
      id: 'guest-2',
      displayName: 'Lisa Jansen',
      email: 'lisa.jansen@klant-bv.nl',
      userPrincipalName: 'lisa.jansen_klant-bv.nl#EXT#@noppa.onmicrosoft.com',
      companyOrDomain: 'klant-bv.nl',
      sharedSitesCount: 2,
      sharedSites: [
        {
          siteId: 'demo-site-4',
          siteTitle: 'Klantportaal & Partner Samenwerking',
          siteUrl: 'https://noppa.sharepoint.com/sites/klantportaal-samenwerking',
          accessType: 'directShare',
          role: 'Lezen & Downloaden',
        },
        {
          siteId: 'demo-site-6',
          siteTitle: 'Marketing & Externe Communicatie',
          siteUrl: 'https://noppa.sharepoint.com/sites/marketing-communicatie',
          accessType: 'directShare',
          role: 'Lezen',
        },
      ],
    },
    {
      id: 'guest-3',
      displayName: 'Mark Bakker',
      email: 'm.bakker@leverancier-ict.com',
      userPrincipalName: 'm.bakker_leverancier-ict.com#EXT#@noppa.onmicrosoft.com',
      companyOrDomain: 'leverancier-ict.com',
      sharedSitesCount: 1,
      sharedSites: [
        {
          siteId: 'demo-site-2',
          siteTitle: 'Project Noppa Cloud Migratie',
          siteUrl: 'https://noppa.sharepoint.com/sites/cloud-migratie-project',
          groupName: 'Project Noppa Cloud Migratie Team',
          accessType: 'groupMember',
          role: 'Gastlid',
        },
      ],
    },
    {
      id: 'guest-4',
      displayName: 'Sarah Smith',
      email: 'sarah.smith@global-consulting.org',
      userPrincipalName: 'sarah.smith_global-consulting.org#EXT#@noppa.onmicrosoft.com',
      companyOrDomain: 'global-consulting.org',
      sharedSitesCount: 1,
      sharedSites: [
        {
          siteId: 'demo-site-4',
          siteTitle: 'Klantportaal & Partner Samenwerking',
          siteUrl: 'https://noppa.sharepoint.com/sites/klantportaal-samenwerking',
          accessType: 'sharingLink',
          role: 'Deellinktoegang',
        },
      ],
    },
  ];

  let filtered = guests;
  if (query && query.trim()) {
    const q = query.toLowerCase().trim();
    filtered = guests.filter(
      (g) =>
        g.displayName.toLowerCase().includes(q) ||
        g.email.toLowerCase().includes(q) ||
        g.companyOrDomain.toLowerCase().includes(q) ||
        g.sharedSites.some(
          (s) =>
            s.siteTitle.toLowerCase().includes(q) ||
            s.siteUrl.toLowerCase().includes(q) ||
            (s.groupName && s.groupName.toLowerCase().includes(q))
        )
    );
  }

  return {
    totalGuests: guests.length,
    guestsWithSiteAccess: guests.length,
    totalSharedSites: 3,
    uniqueExternalDomains: 4,
    guests: filtered,
  };
}
