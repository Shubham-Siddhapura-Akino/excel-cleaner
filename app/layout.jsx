import './globals.css';

export const metadata = {
  title: 'Excel Cleaner Studio | Next.js + Fastify',
  description: 'Clean, normalize, and batch process Excel workbooks while preserving 100% of formatting, formulas, and sheet structures. Deployed on Vercel.',
};

export default function RootLayout({ children }) {
  return (
    <html lang="en" className="dark">
      <head>
        <link rel="preconnect" href="https://fonts.googleapis.com" />
        <link rel="preconnect" href="https://fonts.gstatic.com" crossOrigin="anonymous" />
        <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet" />
      </head>
      <body className="bg-slate-950 text-slate-100 min-h-screen antialiased flex flex-col font-sans selection:bg-emerald-500 selection:text-white">
        {children}
      </body>
    </html>
  );
}
