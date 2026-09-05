"use client";

import { useState } from "react";
import {
  LEGAL_CONTACT,
  PUBLIC_ABOUT_URL,
  PUBLIC_CUBE_LOGO,
  PUBLIC_EDITOR_ALIAS_URL,
  PUBLIC_EDITOR_URL,
  PUBLIC_INTEGRATE_URL,
  PUBLIC_LEGAL_URL,
  PUBLIC_SOURCE_URL,
  PUBLIC_WEBSITE_REPO,
  UPSTREAM_COMPONENTS,
} from "@/utils/attribution";

type EmbedLegalProps = {
  visible: boolean;
};

export function EmbedLegal({ visible }: EmbedLegalProps) {
  const [open, setOpen] = useState(false);
  if (!visible) return null;

  return (
    <>
      <button
        type="button"
        className="absolute bottom-2 left-2 z-[60] flex items-center gap-2 rounded-md border border-border/80 bg-background/95 px-2 py-1 text-left shadow-sm backdrop-blur-sm hover:bg-background"
        onClick={() => setOpen(true)}
        aria-haspopup="dialog"
        aria-expanded={open}
      >
        <img
          src={PUBLIC_CUBE_LOGO}
          alt=""
          width={20}
          height={20}
          className="h-5 w-5 rounded-sm object-contain"
        />
        <span className="flex flex-col leading-tight">
          <span className="text-[10px] font-semibold tracking-wide">EDITOR GRATUITO</span>
          <span className="text-[10px] text-text-secondary">Sobre o editor · AGPL</span>
        </span>
      </button>
      {open ? (
        <div
          className="absolute inset-0 z-[70] flex items-center justify-center bg-black/45 p-4"
          role="presentation"
          onClick={() => setOpen(false)}
        >
          <div
            role="dialog"
            aria-modal="true"
            aria-labelledby="embed-legal-title"
            className="max-h-[88vh] w-full max-w-lg overflow-y-auto rounded-xl border border-border bg-background p-5 text-sm shadow-xl"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="mb-4 flex items-center gap-3">
              <img
                src={PUBLIC_CUBE_LOGO}
                alt="EDITOR GRATUITO"
                width={40}
                height={40}
                className="h-10 w-10 rounded-lg object-contain"
              />
              <div>
                <h2 id="embed-legal-title" className="text-base font-semibold">
                  Sobre o editor
                </h2>
                <p className="text-xs text-text-secondary">
                  EDITOR GRATUITO · GNU AGPL v3
                </p>
              </div>
            </div>
            <div className="space-y-3 text-text-secondary leading-relaxed">
              <p>
                Este documento é editado no projeto livre{" "}
                <a className="text-primary underline" href={PUBLIC_EDITOR_URL} target="_blank" rel="noreferrer">
                  editor.app.br
                </a>{" "}
                (também{" "}
                <a className="text-primary underline" href={PUBLIC_EDITOR_ALIAS_URL} target="_blank" rel="noreferrer">
                  editor.com.br
                </a>
                ), incorporado aqui por <strong className="text-foreground">iframe</strong> +{" "}
                <code className="text-foreground">postMessage</code>. O sistema anfitrião
                não embute OnlyOffice, sdkjs nem x2t; só troca os bytes do seu arquivo.
              </p>
              <p>
                Software livre sob <strong className="text-foreground">GNU AGPL v3</strong>,
                Seção 13: quem usa o editor pela rede tem direito ao código-fonte
                correspondente.
              </p>
              <ul className="list-disc space-y-1 pl-4">
                <li>
                  Fonte:{" "}
                  <a className="text-primary underline" href={PUBLIC_SOURCE_URL} target="_blank" rel="noreferrer">
                    {PUBLIC_SOURCE_URL}
                  </a>
                </li>
                <li>
                  Aviso legal:{" "}
                  <a className="text-primary underline" href={PUBLIC_LEGAL_URL} target="_blank" rel="noreferrer">
                    {PUBLIC_LEGAL_URL}
                  </a>
                </li>
                <li>
                  Sobre:{" "}
                  <a className="text-primary underline" href={PUBLIC_ABOUT_URL} target="_blank" rel="noreferrer">
                    {PUBLIC_ABOUT_URL}
                  </a>
                </li>
                <li>
                  Integração iframe:{" "}
                  <a className="text-primary underline" href={PUBLIC_INTEGRATE_URL} target="_blank" rel="noreferrer">
                    {PUBLIC_INTEGRATE_URL}
                  </a>
                </li>
                <li>
                  Repositório:{" "}
                  <a className="text-primary underline" href={PUBLIC_WEBSITE_REPO} target="_blank" rel="noreferrer">
                    github.com/editor-app-br/editor-website
                  </a>
                </li>
              </ul>
              <p>
                Motor ONLYOFFICE Document Server Community Edition © Ascensio System
                SIA (AGPL v3 e cláusulas de marca). O item <strong className="text-foreground">Arquivo → Sobre</strong> do
                próprio editor também mostra esses créditos — não o oculte.
              </p>
              <ul className="list-disc space-y-1 pl-4">
                {UPSTREAM_COMPONENTS.map((item) => (
                  <li key={item.name}>
                    {item.name} — {item.author} · {item.license} ·{" "}
                    <a className="text-primary underline" href={item.url} target="_blank" rel="noreferrer">
                      fonte
                    </a>
                  </li>
                ))}
              </ul>
              <p>
                PRATA CONSULTORIA EM TECNOLOGIA DA INFORMACAO LTDA · CNPJ
                48.889.573/0001-44 ·{" "}
                <a className="text-primary underline" href={`mailto:${LEGAL_CONTACT}`}>
                  {LEGAL_CONTACT}
                </a>
              </p>
            </div>
            <div className="mt-5 flex justify-end">
              <button
                type="button"
                className="rounded-md bg-emerald-700 px-3 py-1.5 text-xs font-medium text-white hover:bg-emerald-800"
                onClick={() => setOpen(false)}
              >
                Fechar
              </button>
            </div>
          </div>
        </div>
      ) : null}
    </>
  );
}
