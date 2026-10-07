import { TenantConnectionConfig } from '@/lib/db/store';

export function hasValidGraphCredentials(config: TenantConnectionConfig | null): boolean {
  if (!config) return false;
  return !!(
    config.tenantId &&
    config.clientId &&
    config.clientSecret &&
    config.tenantId !== '00000000-0000-0000-0000-000000000000' &&
    config.clientSecret.length > 5
  );
}

export async function getGraphAccessToken(config: TenantConnectionConfig): Promise<string> {
  const tokenEndpoint = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;

  const params = new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    grant_type: 'client_credentials',
    scope: 'https://graph.microsoft.com/.default',
  });

  const response = await fetch(tokenEndpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString(),
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    const desc = errorData.error_description || response.statusText;
    throw new Error(`Graph Authenticatie Mislukt: ${desc}`);
  }

  const data = await response.json();
  return data.access_token;
}

export async function testTenantConnection(
  config: TenantConnectionConfig
): Promise<{ success: boolean; tenantName?: string; error?: string }> {
  try {
    const token = await getGraphAccessToken(config);
    const orgRes = await fetch('https://graph.microsoft.com/v1.0/organization', {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!orgRes.ok) {
      throw new Error(`Organisatiegegevens ophalen mislukt: HTTP ${orgRes.status}`);
    }

    const orgData = await orgRes.json();
    const tenantName = orgData.value?.[0]?.displayName || 'Microsoft 365 Tenant';

    return {
      success: true,
      tenantName,
    };
  } catch (err: any) {
    return {
      success: false,
      error: err.message || 'Onbekende authenticatiefout',
    };
  }
}
