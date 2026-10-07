// Typings voor SharePoint Access Governance & Permissions

export type SharingCapability =
  | 'Disabled'
  | 'ExistingExternalUserSharingOnly'
  | 'ExternalUserSharingOnly'
  | 'ExternalUserAndGuestSharing';

export interface GraphSite {
  id: string;
  title: string;
  url: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
  description?: string;
  template?: 'Communication' | 'Team' | 'Standard' | string;
  isTeam?: boolean;
  storageQuotaGb?: number;
  primaryAdmin?: string;
  admins?: string[];
  members?: string[];
  visitors?: string[];
  sharingCapability?: SharingCapability;
  externalSharingCapability?: string;
  sensitivityLabel?: string;
  lcid?: number;
  timeZone?: string;
  createdAt?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
}

export type SharePointScopeLevel = 'site' | 'library' | 'list' | 'item';
export type SharePointPermissionLevel =
  | 'Full Control'
  | 'Contribute'
  | 'Read'
  | 'Limited Access'
  | 'site'
  | 'library'
  | 'list'
  | 'item'
  | string;

export interface SharePointAccessEntity {
  id: string;
  scopeLevel?: SharePointPermissionLevel;
  roles?: ('owner' | 'write' | 'read' | string)[];
  role: 'owner' | 'editor' | 'reader' | string;
  permissionLevel?: 'Full Control' | 'Contribute' | 'Read' | 'Limited Access' | string;
  principalName: string;
  displayName: string;
  email?: string;
  userType: 'member' | 'guest' | 'user' | 'externalGuest' | 'group' | 'siteGroup' | 'sharingLink' | 'application' | string;
  isExternal: boolean;
  linkDetails?: {
    type?: string;
    scope?: 'anonymous' | 'organization' | 'users' | string;
    webUrl?: string;
    preventsDownload?: boolean;
  };
  sharingLink?: {
    scope?: 'anonymous' | 'organization' | 'users' | string;
    type?: string;
    webUrl?: string;
  };
  inherited: boolean;
  inheritedFrom?: string;
  expirationDateTime?: string;
  hasPassword?: boolean;
}

export interface SharePointContainer {
  id: string;
  name: string;
  displayName?: string;
  type: 'documentLibrary' | 'genericList' | 'list' | string;
  webUrl?: string;
  itemCount?: number;
  hasUniquePermissions: boolean;
  externalPermissionsCount?: number;
  permissions?: SharePointAccessEntity[];
}

export interface SharePointItemPermissionSummary {
  id: string;
  name: string;
  isFolder: boolean;
  webUrl: string;
  hasUniquePermissions: boolean;
  externalSharesCount: number;
  permissions: SharePointAccessEntity[];
}

export interface SiteAccessAuditSummary {
  siteId: string;
  siteTitle: string;
  totalPermissions: number;
  externalCount: number;
  sharingLinksCount: number;
  uniquePermissionItemsCount: number;
  sitePermissions: SharePointAccessEntity[];
  containers: SharePointContainer[];
}

export interface SharedSiteReference {
  siteId: string;
  siteTitle: string;
  siteUrl: string;
  groupName?: string;
  source?: 'm365_group' | 'direct_permission' | 'sharing_link' | string;
  accessType?: 'groupMember' | 'directShare' | 'sharingLink' | string;
  role: string;
}

export interface ExternalGuestAccessRecord {
  id: string;
  displayName: string;
  email: string;
  userPrincipalName: string;
  companyOrDomain: string;
  sharedSitesCount: number;
  sharedSites: SharedSiteReference[];
}

export interface TenantExternalAuditResult {
  totalGuests: number;
  guestsWithSiteAccess: number;
  totalSharedSites: number;
  uniqueExternalDomains: number;
  guests: ExternalGuestAccessRecord[];
}
