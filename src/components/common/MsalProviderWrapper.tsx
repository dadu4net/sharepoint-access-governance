'use client';

import React, { ReactNode, useEffect, useState } from 'react';
import { PublicClientApplication, Configuration, LogLevel } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';

interface MsalProviderWrapperProps {
  children: ReactNode;
}

export function MsalProviderWrapper({ children }: MsalProviderWrapperProps) {
  const [msalInstance, setMsalInstance] = useState<PublicClientApplication | null>(null);

  useEffect(() => {
    // Haal de opgeslagen configuratie op uit localStorage
    const savedConfig = localStorage.getItem('noppa_graph_config');
    let clientId = process.env.NEXT_PUBLIC_CLIENT_ID || '';
    let authority = process.env.NEXT_PUBLIC_TENANT_ID 
      ? `https://login.microsoftonline.com/${process.env.NEXT_PUBLIC_TENANT_ID}`
      : 'https://login.microsoftonline.com/common';

    if (savedConfig) {
      try {
        const parsed = JSON.parse(savedConfig);
        if (parsed.clientId) clientId = parsed.clientId;
        if (parsed.tenantId) authority = `https://login.microsoftonline.com/${parsed.tenantId}`;
      } catch (e) {
        console.error('Kon opgeslagen Graph config niet parsen', e);
      }
    }

    if (clientId) {
      const msalConfig: Configuration = {
        auth: {
          clientId: clientId,
          authority: authority,
          redirectUri: typeof window !== 'undefined' ? window.location.origin : '/',
          postLogoutRedirectUri: typeof window !== 'undefined' ? window.location.origin : '/',
        },
        cache: {
          cacheLocation: 'sessionStorage',
        },
        system: {
          loggerOptions: {
            loggerCallback: (level, message, containsPii) => {
              if (containsPii) return;
              switch (level) {
                case LogLevel.Error:
                  console.error(message);
                  return;
                case LogLevel.Warning:
                  console.warn(message);
                  return;
              }
            },
            piiLoggingEnabled: false,
          },
        },
      };

      const pca = new PublicClientApplication(msalConfig);
      pca.initialize().then(() => {
        setMsalInstance(pca);
      });
    }
  }, []);

  // Als we nog geen MSAL instance hebben (omdat ClientID niet is ingesteld, of het nog laadt)
  // tonen we toch de children, zodat de configuratie modal wel geopend kan worden!
  if (!msalInstance) {
    return <>{children}</>;
  }

  return (
    <MsalProvider instance={msalInstance}>
      {children}
    </MsalProvider>
  );
}
