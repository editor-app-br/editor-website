"use client";

import { Github, ShieldCheck, GitBranch } from "lucide-react";
import Image from "next/image";
import Link from "next/link";
import { usePageTitle } from "@/hooks/use-page-title";
import { PUBLIC_WEBSITE_REPO, UPSTREAM_COMPONENTS } from "@/utils/attribution";

export function AboutView() {
  usePageTitle("Sobre");

  return (
    <div className="max-w-4xl mx-auto space-y-10 animate-in fade-in slide-in-from-bottom-4 duration-500 py-6 px-4">
      {/* Hero Header */}
      <div className="text-center space-y-4">
        <div className="inline-flex flex-col items-center justify-center gap-3 mb-2">
          <Image
            width={80}
            height={80}
            src="/editor_cube_transparent.png"
            className="w-20 h-20 rounded-3xl object-contain"
            alt="EDITOR GRATUITO"
          />
          <span className="font-bold text-lg tracking-wide text-foreground">
            EDITOR GRATUITO
          </span>
        </div>
        <p className="text-lg text-text-secondary max-w-2xl mx-auto">
          Suíte de escritório online moderna, 100% gratuita e de alta fidelidade. Edite Word, Excel, PowerPoint e PDF diretamente no navegador, sem envio de dados para a nuvem.
        </p>
      </div>

      {/* Corporate Sponsor in Text */}
      <div className="p-8 bg-white/5 dark:bg-white/5 border border-border dark:border-white/10 rounded-3xl space-y-4">
        <div className="text-xs uppercase tracking-wider font-semibold text-primary">
          Desenvolvimento e Patrocínio Institucional
        </div>
        <h2 className="text-xl font-bold text-foreground">
          PRATA CONSULTORIA EM TECNOLOGIA DA INFORMACAO LTDA
        </h2>
        <p className="text-sm text-text-secondary dark:text-slate-400 leading-relaxed">
          Este serviço público e de código aberto é patrocinado e mantido no Brasil pela <strong>Prata Consultoria</strong>, promovendo a democratização do software livre, a soberania digital e o acesso a ferramentas de escritório seguras para todos os cidadãos e empresas.
        </p>
        <div className="pt-2 flex flex-wrap items-center gap-4 text-xs text-text-secondary">
          <span className="px-3 py-1 bg-white/10 rounded-full font-mono border border-border">
            CNPJ: 48.889.573/0001-44
          </span>
          <a
            href="https://prata.dev.br"
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center gap-1 text-primary font-semibold hover:underline"
          >
            prata.dev.br
          </a>
        </div>
      </div>

      {/* Feature Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
        <div className="p-8 bg-white/5 dark:bg-white/5 border border-border dark:border-white/10 rounded-3xl hover:border-primary/50 transition-all">
          <div className="w-12 h-12 bg-primary/10 dark:bg-primary/20 rounded-2xl flex items-center justify-center mb-6">
            <Github className="w-6 h-6 text-primary" />
          </div>
          <h3 className="text-xl font-bold mb-3">Código Aberto (GNU AGPL v3)</h3>
          <p className="text-text-secondary dark:text-slate-400 text-sm leading-relaxed mb-6">
            Distribuído sob a licença GNU AGPL v3. Em conformidade com a Seção 13, todo usuário tem acesso ao código fonte correspondente.
          </p>
          <a
            href={PUBLIC_WEBSITE_REPO}
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center text-sm font-semibold text-primary hover:underline"
          >
            Ver repositório no GitHub
          </a>
        </div>

        <div className="p-8 bg-white/5 dark:bg-white/5 border border-border dark:border-white/10 rounded-3xl hover:border-primary/50 transition-all">
          <div className="w-12 h-12 bg-primary/10 dark:bg-primary/20 rounded-2xl flex items-center justify-center mb-6">
            <ShieldCheck className="w-6 h-6 text-primary" />
          </div>
          <h3 className="text-xl font-bold mb-3">Motor ONLYOFFICE™</h3>
          <p className="text-text-secondary dark:text-slate-400 text-sm leading-relaxed">
            Construído sobre o núcleo ONLYOFFICE DocumentServer Community Edition (© Ascensio System SIA) sob licença AGPL v3, respeitando integralmente as diretrizes comunitárias.
          </p>
        </div>

        <div className="p-8 bg-white/5 dark:bg-white/5 border border-border dark:border-white/10 rounded-3xl hover:border-primary/50 transition-all md:col-span-2">
          <div className="w-12 h-12 bg-primary/10 dark:bg-primary/20 rounded-2xl flex items-center justify-center mb-6">
            <GitBranch className="w-6 h-6 text-primary" />
          </div>
          <h3 className="text-xl font-bold mb-3">Origem do código</h3>
          <p className="text-text-secondary dark:text-slate-400 text-sm leading-relaxed mb-4">
            Este projeto é um fork comunitário brasileiro. A arquitetura do
            cliente local deriva do projeto ZIZIYI; o motor de edição é
            ONLYOFFICE Community Edition; a conversão WASM usa CryptPad x2t.
          </p>
          <ul className="space-y-3 text-sm mb-6">
            <li className="text-text-secondary">
              <span className="font-semibold text-foreground">
                editor.app.br (este projeto)
              </span>
              {" — "}
              © PRATA CONSULTORIA / comunidade
              {" · "}
              <a
                href={PUBLIC_WEBSITE_REPO}
                target="_blank"
                rel="noopener noreferrer"
                className="text-primary hover:underline"
              >
                repositório
              </a>
            </li>
            {UPSTREAM_COMPONENTS.map((item) => (
              <li key={item.name} className="text-text-secondary">
                <span className="font-semibold text-foreground">{item.name}</span>
                {" — "}
                {item.author}
                {" · "}
                <a
                  href={item.url}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="text-primary hover:underline"
                >
                  repositório
                </a>
              </li>
            ))}
          </ul>
          <div className="flex flex-wrap gap-4 text-sm font-semibold">
            <Link href="/integrate" className="text-primary hover:underline">
              Integrar via iframe
            </Link>
            <Link href="/source" className="text-primary hover:underline">
              Código-fonte (AGPL §13)
            </Link>
            <Link href="/legal" className="text-primary hover:underline">
              Aviso legal
            </Link>
          </div>
        </div>
      </div>
    </div>
  );
}
