import fs from 'fs';
import path from 'path';
import os from 'os';

export interface TenantConnectionConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  tenantName?: string;
  configuredAt?: string;
}

interface AppStore {
  connection: TenantConnectionConfig | null;
}

const DEFAULT_STORE: AppStore = {
  connection: null,
};

function getStoragePath(): string {
  const homeDir = os.homedir();
  const dirPath = path.join(homeDir, '.sharepoint-access-governance');
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
  return path.join(dirPath, 'config.json');
}

export function readStore(): AppStore {
  try {
    const filePath = getStoragePath();
    if (!fs.existsSync(filePath)) {
      writeStore(DEFAULT_STORE);
      return DEFAULT_STORE;
    }
    const data = fs.readFileSync(filePath, 'utf-8');
    return JSON.parse(data) as AppStore;
  } catch (err) {
    console.error('[Store] Fout bij uitlezen store:', err);
    return DEFAULT_STORE;
  }
}

export function writeStore(store: AppStore): void {
  try {
    const filePath = getStoragePath();
    fs.writeFileSync(filePath, JSON.stringify(store, null, 2), 'utf-8');
  } catch (err) {
    console.error('[Store] Fout bij wegschrijven store:', err);
  }
}

export function getTenantConnection(): TenantConnectionConfig | null {
  return readStore().connection;
}

export function saveTenantConnection(config: TenantConnectionConfig): void {
  const store = readStore();
  store.connection = {
    ...config,
    configuredAt: new Date().toISOString(),
  };
  writeStore(store);
}

export function clearTenantConnection(): void {
  const store = readStore();
  store.connection = null;
  writeStore(store);
}

// Aliases for compatibility
export const getTenantCredentials = getTenantConnection;
export const saveTenantCredentials = saveTenantConnection;
export const clearTenantCredentials = clearTenantConnection;

