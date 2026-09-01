"use client";

import { useEffect, useState } from "react";
import { Download, X, Sparkles } from "lucide-react";

interface BeforeInstallPromptEvent extends Event {
  prompt: () => Promise<void>;
  userChoice: Promise<{ outcome: "accepted" | "dismissed" }>;
}

export function PWAInstallPrompt() {
  const [deferredPrompt, setDeferredPrompt] = useState<BeforeInstallPromptEvent | null>(null);
  const [showPrompt, setShowPrompt] = useState(false);
  const [isInstalled, setIsInstalled] = useState(false);

  useEffect(() => {
    // Do not register a SW. Cached OnlyOffice/WASM + COEP crashed Chrome tabs.
    if (typeof window !== "undefined" && "serviceWorker" in navigator) {
      navigator.serviceWorker.getRegistrations().then((regs) => {
        regs.forEach((reg) => reg.unregister());
      });
    }

    // Check if already in standalone mode
    if (
      typeof window !== "undefined" &&
      (window.matchMedia("(display-mode: standalone)").matches ||
        (window.navigator as unknown as { standalone?: boolean }).standalone === true)
    ) {
      setIsInstalled(true);
      return;
    }

    const handleBeforeInstallPrompt = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e as BeforeInstallPromptEvent);

      const dismissedUntil = localStorage.getItem("pwa_install_dismissed_until");
      if (dismissedUntil && Date.now() < parseInt(dismissedUntil, 10)) {
        return;
      }
      setShowPrompt(true);
    };

    window.addEventListener("beforeinstallprompt", handleBeforeInstallPrompt);

    window.addEventListener("appinstalled", () => {
      setIsInstalled(true);
      setShowPrompt(false);
      setDeferredPrompt(null);
    });

    return () => {
      window.removeEventListener("beforeinstallprompt", handleBeforeInstallPrompt);
    };
  }, []);

  const handleInstallClick = async () => {
    if (!deferredPrompt) return;
    await deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    if (outcome === "accepted") {
      setIsInstalled(true);
      setShowPrompt(false);
    }
    setDeferredPrompt(null);
  };

  const handleDismiss = () => {
    setShowPrompt(false);
    localStorage.setItem(
      "pwa_install_dismissed_until",
      (Date.now() + 7 * 24 * 60 * 60 * 1000).toString(),
    );
  };

  if (!showPrompt || isInstalled) return null;

  return (
    <div className="fixed bottom-6 right-6 z-50 max-w-sm w-[calc(100vw-3rem)] animate-in fade-in slide-in-from-bottom-5 duration-300">
      <div className="relative p-4 rounded-2xl bg-background/95 dark:bg-slate-900/95 backdrop-blur-md border border-border dark:border-white/15 shadow-2xl flex flex-col gap-3">
        <button
          onClick={handleDismiss}
          className="absolute top-3 right-3 text-text-secondary hover:text-foreground p-1 rounded-full hover:bg-white/10 transition-colors"
          aria-label="Fechar"
        >
          <X className="w-4 h-4" />
        </button>

        <div className="flex items-start gap-3">
          <div className="relative w-12 h-12 rounded-xl overflow-hidden bg-primary/10 border border-primary/20 shrink-0 p-1 flex items-center justify-center">
            <img
              src="/icon-192x192.png"
              alt="editor.app.br"
              className="w-full h-full object-contain"
            />
          </div>
          <div className="pr-6">
            <div className="flex items-center gap-1.5 font-bold text-sm text-foreground">
              <span>Instalar editor.app.br</span>
              <Sparkles className="w-3.5 h-3.5 text-amber-500" />
            </div>
            <p className="text-xs text-text-secondary dark:text-slate-400 mt-0.5 leading-relaxed">
              Adicione à sua área de trabalho para acesso rápido e edição 100% offline.
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 pt-1">
          <button
            onClick={handleInstallClick}
            className="flex-1 inline-flex items-center justify-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white rounded-xl text-xs font-semibold shadow-md shadow-emerald-600/20 transition-all cursor-pointer"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Instalar Aplicativo</span>
          </button>
          <button
            onClick={handleDismiss}
            className="px-3 py-2 text-xs font-medium text-text-secondary hover:text-foreground hover:bg-white/5 rounded-xl transition-colors cursor-pointer"
          >
            Agora não
          </button>
        </div>
      </div>
    </div>
  );
}
