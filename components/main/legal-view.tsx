"use client";

import Link from "next/link";
import { usePageTitle } from "@/hooks/use-page-title";
import { LEGAL_CONTACT, PUBLIC_WEBSITE_REPO, UPSTREAM_COMPONENTS } from "@/utils/attribution";

export function LegalView() {
  usePageTitle("Aviso legal");

  return (
    <div className="max-w-3xl mx-auto space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500 py-6 px-4 prose prose-sm dark:prose-invert max-w-none">
      <div className="not-prose space-y-3">
        <h1 className="text-2xl font-bold tracking-tight">
          Aviso legal e licenciamento
        </h1>
        <p className="text-sm text-text-secondary">
          EDITOR GRATUITO · editor.app.br · PRATA CONSULTORIA · CNPJ
          48.889.573/0001-44
        </p>
      </div>

      <section>
        <h2>1. Software livre (GNU AGPL v3, Seção 13)</h2>
        <p>
          Este software é distribuído sob a{" "}
          <strong>GNU Affero General Public License v3.0</strong>. Usuários que
          interagem com o serviço pela rede têm direito ao código-fonte
          correspondente.
        </p>
        <ul>
          <li>
            Página de código-fonte:{" "}
            <Link href="/source">editor.app.br/source</Link>
          </li>
          <li>
            Repositório público:{" "}
            <a
              href={PUBLIC_WEBSITE_REPO}
              target="_blank"
              rel="noopener noreferrer"
            >
              github.com/editor-app-br/editor-website
            </a>
          </li>
          <li>
            Contato:{" "}
            <a href={`mailto:${LEGAL_CONTACT}`}>{LEGAL_CONTACT}</a>
          </li>
        </ul>
      </section>

      <section>
        <h2>2. Atribuições upstream</h2>
        {UPSTREAM_COMPONENTS.map((item) => (
          <div key={item.name} className="mb-6">
            <h3>{item.name}</h3>
            <p>
              <strong>Autor:</strong> {item.author}
              <br />
              <strong>Licença:</strong> {item.license}
              <br />
              <strong>Contribuição:</strong> {item.contribution}
              <br />
              <a href={item.url} target="_blank" rel="noopener noreferrer">
                {item.url}
              </a>
            </p>
          </div>
        ))}
        <p>
          <strong>ONLYOFFICE™</strong> é marca registrada da Ascensio System
          SIA. Este projeto usa o motor ONLYOFFICE Community Edition conforme
          as regras comunitárias de créditos e marca.
        </p>
      </section>

      <section>
        <h2>3. Integração Thin-Embed</h2>
        <p>
          Integrações externas usam <code>&lt;iframe&gt;</code> +{" "}
          <code>postMessage</code>. O sistema host troca apenas bytes de
          documentos; documentos são processados localmente no navegador.
        </p>
      </section>

      <p className="not-prose text-sm text-text-secondary">
        Texto integral:{" "}
        <a href="/LEGAL_NOTICE.md" className="text-primary hover:underline">
          LEGAL_NOTICE.md
        </a>
      </p>
    </div>
  );
}
