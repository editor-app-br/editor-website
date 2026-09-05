"use client";

import Link from "next/link";
import { ExternalLink, FileCode2, GitBranch } from "lucide-react";
import { usePageTitle } from "@/hooks/use-page-title";
import {
  LEGAL_CONTACT,
  PROJECT_REPOS,
  UPSTREAM_COMPONENTS,
} from "@/utils/attribution";

export function SourceView() {
  usePageTitle("Código-fonte");

  return (
    <div className="max-w-3xl mx-auto space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500 py-6 px-4">
      <div className="space-y-3">
        <h1 className="text-2xl font-bold tracking-tight">Código-fonte</h1>
        <p className="text-sm text-text-secondary leading-relaxed">
          O EDITOR GRATUITO (editor.app.br) é software livre sob{" "}
          <strong className="text-foreground">GNU AGPL v3</strong>. Pela{" "}
          <strong className="text-foreground">Seção 13</strong> da licença,
          qualquer pessoa que usa este serviço pela rede pode obter o código
          correspondente à versão em execução.
        </p>
      </div>

      <section className="p-6 rounded-2xl border border-border bg-muted/30 space-y-4">
        <div className="flex items-center gap-2 text-primary">
          <GitBranch className="w-5 h-5" />
          <h2 className="text-lg font-semibold">Repositórios deste projeto</h2>
        </div>
        <ul className="space-y-3 text-sm">
          {Object.values(PROJECT_REPOS).map((repo) => (
            <li key={repo.web} className="flex flex-col gap-1">
              <span className="font-medium text-foreground">{repo.label}</span>
              <a
                href={repo.web}
                target="_blank"
                rel="noopener noreferrer"
                className="inline-flex items-center gap-1 text-primary hover:underline break-all"
              >
                {repo.web}
                <ExternalLink className="w-3.5 h-3.5 shrink-0" />
              </a>
            </li>
          ))}
        </ul>
        <p className="text-xs text-text-secondary leading-relaxed">
          Site e extensão são públicos no GitHub. O repositório de deploy no
          GitLab interno pode exigir login; peça uma cópia do tarball
          correspondente em{" "}
          <a
            href={`mailto:${LEGAL_CONTACT}`}
            className="text-primary hover:underline"
          >
            {LEGAL_CONTACT}
          </a>
          .
        </p>
      </section>

      <section className="p-6 rounded-2xl border border-border bg-muted/30 space-y-4">
        <div className="flex items-center gap-2 text-primary">
          <FileCode2 className="w-5 h-5" />
          <h2 className="text-lg font-semibold">Componentes upstream</h2>
        </div>
        <ul className="space-y-4 text-sm">
          {UPSTREAM_COMPONENTS.map((item) => (
            <li key={item.name} className="space-y-1">
              <div className="font-medium text-foreground">{item.name}</div>
              <div className="text-text-secondary">{item.author}</div>
              <div className="text-text-secondary">{item.license}</div>
              <p className="text-text-secondary leading-relaxed">
                {item.contribution}
              </p>
              <a
                href={item.url}
                target="_blank"
                rel="noopener noreferrer"
                className="inline-flex items-center gap-1 text-primary hover:underline"
              >
                {item.url.replace(/^https:\/\//, "")}
                <ExternalLink className="w-3.5 h-3.5" />
              </a>
            </li>
          ))}
        </ul>
      </section>

      <p className="text-sm text-text-secondary">
        Integração iframe + plugin Agente:{" "}
        <Link href="/integrate" className="text-primary font-medium hover:underline">
          /integrate
        </Link>
        {" · "}
        Aviso legal completo:{" "}
        <Link href="/legal" className="text-primary font-medium hover:underline">
          /legal
        </Link>
        {" · "}
        <a
          href="/LEGAL_NOTICE.md"
          className="text-primary font-medium hover:underline"
        >
          LEGAL_NOTICE.md
        </a>
        {" · "}
        <a
          href="/LICENSE.txt"
          className="text-primary font-medium hover:underline"
        >
          LICENSE.txt
        </a>
      </p>
    </div>
  );
}
