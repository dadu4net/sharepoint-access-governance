'use client';

import React from 'react';
import {
  Shield,
  Lock,
  RefreshCw,
  Check,
  AlertTriangle,
  X,
  Users,
  Settings,
  Folder,
  ExternalLink,
  Layers,
  FileText,
  Search,
  Info,
  Share2,
} from 'lucide-react';

export type NoppaOfficialIconName =
  | 'brain'
  | 'calendar'
  | 'chat'
  | 'checklist'
  | 'clock'
  | 'cloud'
  | 'dashboard'
  | 'document'
  | 'email'
  | 'gear'
  | 'growth'
  | 'handshake'
  | 'idea'
  | 'lightning'
  | 'link'
  | 'rocket'
  | 'search'
  | 'shield'
  | 'sparkles'
  | 'star'
  | 'sync'
  | 'target'
  | 'team'
  | 'workflow';

export type NoppaIconName =
  | NoppaOfficialIconName
  | 'sharepoint'
  | 'teams'
  | 'security-shield'
  | 'security-lock'
  | 'refresh'
  | 'close'
  | 'check'
  | 'alert-triangle'
  | 'users'
  | 'settings'
  | 'folder'
  | 'external-link'
  | 'info'
  | string;

interface NoppaIconProps {
  name: NoppaIconName;
  size?: number;
  className?: string;
  alt?: string;
  style?: React.CSSProperties;
}

const OFFICIAL_ICONS: Set<string> = new Set([
  'brain',
  'calendar',
  'chat',
  'checklist',
  'clock',
  'cloud',
  'dashboard',
  'document',
  'email',
  'gear',
  'growth',
  'handshake',
  'idea',
  'lightning',
  'link',
  'rocket',
  'search',
  'shield',
  'sparkles',
  'star',
  'sync',
  'target',
  'team',
  'workflow',
]);

export function NoppaIcon({ name, size = 24, className = '', alt, style }: NoppaIconProps) {
  // 1. If it's one of the official 24 Noppa SVG gradient vector icons
  if (OFFICIAL_ICONS.has(name)) {
    return (
      <img
        src={`/assets/icons/${name}.svg`}
        alt={alt || `Noppa · ${name}`}
        width={size}
        height={size}
        className={`noppa-icon ${className}`}
        style={{
          width: `${size}px`,
          height: `${size}px`,
          display: 'inline-block',
          verticalAlign: 'middle',
          flexShrink: 0,
          ...style,
        }}
        onError={(e) => {
          e.currentTarget.style.visibility = 'hidden';
        }}
      />
    );
  }

  // 2. Lucide UI icon mappings for system controls
  const lucideProps = {
    size,
    className,
    style: { flexShrink: 0, ...style },
  };

  switch (name) {
    case 'sharepoint':
      return <Layers {...lucideProps} />;
    case 'teams':
      return <Users {...lucideProps} />;
    case 'security-shield':
      return <Shield {...lucideProps} />;
    case 'security-lock':
      return <Lock {...lucideProps} />;
    case 'refresh':
      return <RefreshCw {...lucideProps} />;
    case 'close':
      return <X {...lucideProps} />;
    case 'check':
      return <Check {...lucideProps} />;
    case 'alert-triangle':
      return <AlertTriangle {...lucideProps} />;
    case 'users':
      return <Users {...lucideProps} />;
    case 'settings':
      return <Settings {...lucideProps} />;
    case 'folder':
      return <Folder {...lucideProps} />;
    case 'external-link':
      return <ExternalLink {...lucideProps} />;
    case 'info':
      return <Info {...lucideProps} />;
    case 'share':
      return <Share2 {...lucideProps} />;
    case 'document':
      return <FileText {...lucideProps} />;
    default:
      return <Shield {...lucideProps} />;
  }
}
