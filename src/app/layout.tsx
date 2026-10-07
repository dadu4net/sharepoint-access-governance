import type { Metadata } from 'next';
import './globals.css';
import { MsalProviderWrapper } from '@/components/common/MsalProviderWrapper';
import { Header } from '@/components/common/Header';

export const metadata: Metadata = {
  title: 'Noppa | SharePoint & Teams Access Governance',
  description: 'Gespecialiseerde 4-niveau autorisatie audit en externe toegangsverkenner voor SharePoint & Microsoft Teams.',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="nl" suppressHydrationWarning>
      <body className="min-h-screen flex flex-col font-kentledge selection:bg-cyan/20 selection:text-cyan">
        <MsalProviderWrapper>
        {/* Navigation / Header */}
        <Header />

        {/* Main Content Body */}
        <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8">
          {children}
        </main>
        </MsalProviderWrapper>

        {/* Footer */}
        <footer className="border-t border-card bg-card/40 py-6 text-center text-xs text-muted">
          <div className="max-w-7xl mx-auto px-4 flex flex-col sm:flex-row items-center justify-between gap-3">
            <p>
              &copy; {new Date().getFullYear()} Noppa Solutions &amp; Consultants. Alle rechten voorbehouden.
            </p>
            <p className="flex items-center gap-2">
              <span className="inline-block w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
              <span>SharePoint Permissions &amp; External Access Auditor</span>
            </p>
          </div>
        </footer>
      </body>
    </html>
  );
}
