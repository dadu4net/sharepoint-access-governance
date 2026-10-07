/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    './src/pages/**/*.{js,ts,jsx,tsx,mdx}',
    './src/components/**/*.{js,ts,jsx,tsx,mdx}',
    './src/app/**/*.{js,ts,jsx,tsx,mdx}',
  ],
  darkMode: ['class', '[data-theme="dark"]'],
  theme: {
    extend: {
      colors: {
        primary: 'var(--color-primary-bg, #F1F5FB)',
        card: {
          DEFAULT: 'var(--color-card-bg, #FFFFFF)',
          header: 'var(--color-card-header, #F8FAFD)',
        },
        heading: 'var(--text-heading, #0F2A66)',
        muted: 'var(--text-muted, #6B7280)',
        hover: 'var(--color-hover-bg, #F8FAFD)',
        input: 'var(--color-input-border, #E4E9F2)',
        cyan: {
          DEFAULT: '#00D0F0',
          50: '#E6FAFD',
          100: '#CCF5FC',
          500: '#00D0F0',
          600: '#00B0D0',
        },
        royal: {
          DEFAULT: '#2060E0',
          50: '#EEF4FD',
          100: '#DDE9FC',
          500: '#2060E0',
          600: '#1A4FB8',
        },
        navy: {
          DEFAULT: '#0F2A66',
          900: '#091738',
        },
      },
      fontFamily: {
        kentledge: ['Kentledge', '-apple-system', 'BlinkMacSystemFont', 'sans-serif'],
      },
      boxShadow: {
        'apple-sm': '0 1px 3px rgba(15, 42, 102, 0.05), 0 1px 2px rgba(15, 42, 102, 0.03)',
        'apple-md': '0 4px 16px rgba(15, 42, 102, 0.08), 0 1px 4px rgba(15, 42, 102, 0.04)',
        'apple-lg': '0 14px 36px rgba(15, 42, 102, 0.12), 0 2px 8px rgba(15, 42, 102, 0.04)',
      },
      animation: {
        'fade-in': 'fadeIn 0.2s cubic-bezier(0.16, 1, 0.3, 1) forwards',
        'scale-up': 'scaleUp 0.2s cubic-bezier(0.16, 1, 0.3, 1) forwards',
      },
      keyframes: {
        fadeIn: {
          from: { opacity: '0' },
          to: { opacity: '1' },
        },
        scaleUp: {
          from: { opacity: '0', transform: 'scale(0.96) translateY(8px)' },
          to: { opacity: '1', transform: 'scale(1) translateY(0)' },
        },
      },
    },
  },
  plugins: [],
};
