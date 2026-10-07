'use client';

import React, { useState, useEffect } from 'react';
import { Sun, Moon, Monitor } from 'lucide-react';

type Theme = 'light' | 'dark' | 'system';

export function ThemeSelector() {
  const [theme, setTheme] = useState<Theme>('system');

  useEffect(() => {
    const saved = localStorage.getItem('noppa-theme') as Theme | null;
    if (saved) {
      setTheme(saved);
      applyTheme(saved);
    } else {
      applyTheme('system');
    }
  }, []);

  const applyTheme = (t: Theme) => {
    const root = document.documentElement;
    if (t === 'system') {
      const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
      root.setAttribute('data-theme', prefersDark ? 'dark' : 'light');
    } else {
      root.setAttribute('data-theme', t);
    }
  };

  const handleSelect = (t: Theme) => {
    setTheme(t);
    localStorage.setItem('noppa-theme', t);
    applyTheme(t);
  };

  return (
    <div
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        padding: '2px',
        backgroundColor: 'var(--bg-primary, #F1F5FB)',
        borderRadius: '8px',
        border: '1px solid var(--border, #E4E9F2)',
      }}
    >
      <button
        type="button"
        onClick={() => handleSelect('light')}
        style={{
          padding: '5px 8px',
          borderRadius: '6px',
          border: 'none',
          cursor: 'pointer',
          display: 'flex',
          alignItems: 'center',
          backgroundColor: theme === 'light' ? 'var(--bg-card, #FFFFFF)' : 'transparent',
          color: theme === 'light' ? 'var(--color-royal, #2060E0)' : 'var(--color-slate, #6B7280)',
          boxShadow: theme === 'light' ? '0 1px 3px rgba(0,0,0,0.08)' : 'none',
        }}
        title="Licht thema"
      >
        <Sun size={14} />
      </button>

      <button
        type="button"
        onClick={() => handleSelect('dark')}
        style={{
          padding: '5px 8px',
          borderRadius: '6px',
          border: 'none',
          cursor: 'pointer',
          display: 'flex',
          alignItems: 'center',
          backgroundColor: theme === 'dark' ? 'var(--bg-card, #FFFFFF)' : 'transparent',
          color: theme === 'dark' ? '#00D0F0' : 'var(--color-slate, #6B7280)',
          boxShadow: theme === 'dark' ? '0 1px 3px rgba(0,0,0,0.2)' : 'none',
        }}
        title="Donker thema"
      >
        <Moon size={14} />
      </button>

      <button
        type="button"
        onClick={() => handleSelect('system')}
        style={{
          padding: '5px 8px',
          borderRadius: '6px',
          border: 'none',
          cursor: 'pointer',
          display: 'flex',
          alignItems: 'center',
          backgroundColor: theme === 'system' ? 'var(--bg-card, #FFFFFF)' : 'transparent',
          color: theme === 'system' ? 'var(--text-heading, #0F2A66)' : 'var(--color-slate, #6B7280)',
          boxShadow: theme === 'system' ? '0 1px 3px rgba(0,0,0,0.08)' : 'none',
        }}
        title="Systeemvoorkeur volgen"
      >
        <Monitor size={14} />
      </button>
    </div>
  );
}
