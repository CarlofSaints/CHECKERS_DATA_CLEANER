import type { Metadata } from "next";
import "./globals.css";
import Image from "next/image";

export const metadata: Metadata = {
  title: "Checkers Data Cleaner",
  description: "Converts raw Checkers B2B vnd-art-sales files into the clean SEPARATE VIEW format",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body className="min-h-screen flex flex-col" style={{ backgroundColor: '#0f172a', color: '#f1f5f9' }}>

        {/* Header */}
        <header style={{ backgroundColor: '#111827', borderBottom: '1px solid #1f2937' }}>
          <div className="max-w-4xl mx-auto px-6 h-14 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <Image
                src="/oj-logo.png"
                alt="OuterJoin"
                height={28}
                width={130}
                style={{ objectFit: 'contain', maxHeight: '28px', width: 'auto' }}
              />
              <span style={{ color: '#475569', fontSize: '0.85rem' }}>|</span>
              <span style={{ color: '#f1f5f9', fontWeight: 600, fontSize: '0.95rem', letterSpacing: '0.01em' }}>
                Checkers Data Cleaner
              </span>
            </div>
            <span style={{ color: '#64748b', fontSize: '0.75rem' }}>Internal Tool</span>
          </div>
        </header>

        {/* Main content */}
        <main className="flex-1 flex flex-col">
          {children}
        </main>

      </body>
    </html>
  );
}
