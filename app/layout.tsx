import type { Metadata } from "next";
import { getMessages } from "next-intl/server";
import { I18nProvider } from "@/components/i18n-provider";
import { ProgressProvider } from "@/components/progress-provider";
import { PWAInstallPrompt } from "@/components/pwa-install-prompt";
import "./globals.css";

export const metadata: Metadata = {
  metadataBase: new URL("https://editor.app.br"),
  title: {
    default: "EDITOR GRATUITO — Editor Online (Word, Excel, PowerPoint, PDF)",
    template: "%s | EDITOR GRATUITO",
  },  description:
    "Suíte de escritório online moderna, 100% gratuita e de alta fidelidade. Abra e edite documentos DOCX, XLSX, PPTX e PDF diretamente no navegador, sem upload para a nuvem. 100% privado e local-first.",
  keywords: [
    "editor online",
    "editor word online",
    "editor excel online",
    "editor powerpoint online",
    "editor pdf online",
    "abrir docx online",
    "office gratis",
    "editor gratuito brasil",
    "editor.app.br",
    "local first office",
  ],
  manifest: "/manifest.json",
  icons: {
    icon: [
      { url: "/favicon.ico", sizes: "any" },
      { url: "/logo.svg", type: "image/svg+xml" },
      { url: "/favicon-16x16.png", sizes: "16x16", type: "image/png" },
      { url: "/favicon-32x32.png", sizes: "32x32", type: "image/png" },
      { url: "/icon-192x192.png", sizes: "192x192", type: "image/png" },
      { url: "/icon-512x512.png", sizes: "512x512", type: "image/png" },
    ],
    shortcut: "/favicon.ico",
    apple: [
      { url: "/apple-touch-icon.png", sizes: "180x180", type: "image/png" },
      { url: "/apple-touch-icon-precomposed.png", sizes: "180x180", type: "image/png" },
    ],
  },
  appleWebApp: {
    capable: true,
    statusBarStyle: "default",
    title: "editor.app.br",
  },
  openGraph: {
    siteName: "editor.app.br",
    title: "EDITOR GRATUITO — Editor Online de Documentos",
    description:
      "Suíte de escritório online moderna e gratuita para toda a comunidade brasileira. Edite Word, Excel, PowerPoint e PDF sem envio de dados para servidores.",
    type: "website",
    locale: "pt_BR",
    images: ["/og.png"],
  },
  twitter: {
    card: "summary_large_image",
    title: "EDITOR GRATUITO — Editor Online",
    description: "Edite Word, Excel, PowerPoint e PDF 100% no seu navegador.",
    images: ["/og.png"],
  },
  robots: {
    index: true,
    follow: true,
  },
  other: {
    "mobile-web-app-capable": "yes",
    "apple-mobile-web-app-capable": "yes",
    "apple-mobile-web-app-title": "editor.app.br",
    "theme-color": "#059669",
  },
};

export default async function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  const messages = await getMessages();

  const preload = () => {
    const theme = document.cookie.match(/theme=([^;]+)/)?.[1] || "";
    const dark = window.matchMedia("(prefers-color-scheme: dark)").matches;
    const isDark = theme == "dark" || (dark && theme != "light");
    document.documentElement.classList.toggle("dark", isDark);
  };

  const jsonLd = {
    "@context": "https://schema.org",
    "@type": "WebApplication",
    name: "editor.app.br",
    url: "https://editor.app.br",
    description:
      "Suíte de escritório online gratuita para visualização e edição de Word, Excel, PowerPoint e PDF diretamente no navegador via WebAssembly.",
    applicationCategory: "BusinessApplication",
    operatingSystem: "Any",
    offers: {
      "@type": "Offer",
      price: "0",
      priceCurrency: "BRL",
    },
    featureList: [
      "Edição completa de DOCX, XLSX, PPTX e PDF",
      "Processamento 100% local no dispositivo (Zero-Knowledge)",
      "Totalmente gratuito e de código aberto sob AGPL v3",
      "Compatibilidade total com Microsoft Office",
      "Funciona Offline via PWA",
    ],
  };

  return (
    <html suppressHydrationWarning lang="pt-BR">
      <head>
        <link rel="manifest" href="/manifest.json" />
        <link rel="apple-touch-icon" sizes="180x180" href="/apple-touch-icon.png" />
        <script>{`(${preload.toString()})()`}</script>
        <script
          dangerouslySetInnerHTML={{
            __html: `if("serviceWorker"in navigator){navigator.serviceWorker.getRegistrations().then(function(r){r.forEach(function(x){x.unregister()})});if(window.caches)caches.keys().then(function(k){k.forEach(function(n){caches.delete(n)})})}`,
          }}
        />
        <script
          type="application/ld+json"
          dangerouslySetInnerHTML={{ __html: JSON.stringify(jsonLd) }}
        />
      </head>
      <body>
        <ProgressProvider>
          <I18nProvider initialMessages={messages}>{children}</I18nProvider>
        </ProgressProvider>
        <PWAInstallPrompt />
      </body>
    </html>
  );
}

