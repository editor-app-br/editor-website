"use client";

import Link from "next/link";
import { Code2, Puzzle, Shield } from "lucide-react";
import { usePageTitle } from "@/hooks/use-page-title";
import { LEGAL_CONTACT, PUBLIC_WEBSITE_REPO } from "@/utils/attribution";

const EMBED_URL = "https://editor.app.br/embed";
const AGENT_GUID = "asc.{7E4A1C90-2B6D-4F11-9A33-8C0E5D71B2A4}";

const HOST_MESSAGES = [
  { type: "open", detail: "Abre o documento (ArrayBuffer + fileName, fileType, editing, theme, lang). Opcional: variant, hideChrome, plugins." },
  { type: "save", detail: "Pede exportação. A resposta chega como saved com os bytes." },
  { type: "command", detail: "Encaminha um comando ao plugin Agente (name + payload)." },
  { type: "setTheme", detail: "theme: dark | light." },
  { type: "print", detail: "O embed devolve print; o host decide como imprimir." },
  { type: "destroy", detail: "Derruba a instância do editor." },
];

const EDITOR_MESSAGES = [
  { type: "ready", detail: "Embed pronto para receber open (repete até o host abrir)." },
  { type: "documentReady", detail: "Documento carregado no motor." },
  { type: "dirty", detail: "value: true quando o usuário edita." },
  { type: "saved", detail: "Bytes exportados (fileName, fileType, mime, bytes)." },
  { type: "commandResult", detail: "Resultado ou error de um command." },
  { type: "completionRequest", detail: "Pedido de autocomplete (prefix / suffix) do ghost." },
  { type: "ghost", detail: "shown / text / accepted da sugestão inline." },
  { type: "escape / print / error / warmed", detail: "Tecla Esc, pedido de impressão, falha, ou cache de warmup." },
];

const AGENT_COMMANDS = [
  { name: "get_text", detail: "Lê o texto (Word, tabelas, slides, planilha ativa)." },
  { name: "get_outline", detail: "Lista títulos (Heading / Título)." },
  { name: "get_selection", detail: "Texto selecionado." },
  { name: "insert_text / type", detail: "Insere ou digita no cursor (fonte, tamanho, negrito opcionais)." },
  { name: "replace_selection", detail: "Substitui a seleção." },
  { name: "format", detail: "Negrito, itálico, alinhamento, estilo, lista." },
  { name: "table", detail: "Insere tabela (rows, cols, data, header)." },
  { name: "layout", detail: "Página (A4/letter, orientação, margens, colunas)." },
  { name: "goto / select / delete", detail: "Navega, seleciona (all ou heading) ou apaga." },
  { name: "find_replace", detail: "Localizar e substituir (uma ocorrência ou all)." },
  { name: "get_range / set_range / goto_cell", detail: "Células de planilha." },
  { name: "suggest / context", detail: "Ghost inline; o embed pede completionRequest ao host." },
  { name: "undo / redo / ping", detail: "Histórico e health-check." },
];

export function IntegrateView() {
  usePageTitle("Integração iframe");

  return (
    <div className="max-w-3xl mx-auto space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500 py-6 px-4 prose prose-sm dark:prose-invert max-w-none">
      <div className="not-prose space-y-3">
        <h1 className="text-2xl font-bold tracking-tight">
          Integrar via iframe
        </h1>
        <p className="text-sm text-text-secondary leading-relaxed">
          Qualquer aplicação web pode incorporar o EDITOR GRATUITO em{" "}
          <code className="text-foreground">https://editor.app.br/embed</code>{" "}
          com <code className="text-foreground">&lt;iframe&gt;</code> e{" "}
          <code className="text-foreground">postMessage</code>. O motor, o
          protocolo e o plugin Agente são software livre sob{" "}
          <strong className="text-foreground">GNU AGPL v3</strong>.
        </p>
      </div>

      <section>
        <div className="flex items-center gap-2 mb-3">
          <Shield className="w-5 h-5 text-primary" />
          <h2 className="text-lg font-semibold m-0">Licença (GNU AGPL v3)</h2>
        </div>
        <p>
          Usar o iframe público não copia o motor OnlyOffice para o seu
          domínio: o host só troca bytes de documentos. Mesmo assim, se você{" "}
          <strong>modificar, redistribuir ou hospedar</strong> este editor
          (incluindo um fork do plugin Agente), a AGPL v3 Seção 13 exige
          disponibilizar o código-fonte correspondente aos usuários que
          interagem pelo rede.
        </p>
        <ul>
          <li>
            Código deste site:{" "}
            <a
              href={PUBLIC_WEBSITE_REPO}
              target="_blank"
              rel="noopener noreferrer"
            >
              github.com/editor-app-br/editor-website
            </a>
          </li>
          <li>
            Aviso legal: <Link href="/legal">/legal</Link>
          </li>
          <li>
            Fonte (AGPL §13): <Link href="/source">/source</Link>
          </li>
          <li>
            Contato:{" "}
            <a href={`mailto:${LEGAL_CONTACT}`}>{LEGAL_CONTACT}</a>
          </li>
        </ul>
        <p>
          Não embuta <code>/v*</code>, sdkjs ou x2t WASM no binário do host.
          O About do editor e as páginas legais permanecem a atribuição
          OnlyOffice / CryptPad / ZIZIYI.
        </p>
      </section>

      <section>
        <div className="flex items-center gap-2 mb-3">
          <Code2 className="w-5 h-5 text-primary" />
          <h2 className="text-lg font-semibold m-0">Thin embed</h2>
        </div>
        <ol>
          <li>
            Coloque um iframe em{" "}
            <a href={EMBED_URL}>{EMBED_URL}</a> (ou{" "}
            <code>/embed?warmup=1</code> só para aquecer o cache).
          </li>
          <li>
            Espere a mensagem <code>ready</code> (origem{" "}
            <code>https://editor.app.br</code>).
          </li>
          <li>
            Envie <code>open</code> com o <code>ArrayBuffer</code> do
            arquivo. O documento fica na memória do navegador; o servidor
            do editor não guarda cópia.
          </li>
        </ol>
        <p>
          Origens permitidas: este site, localhost e apps desktop (Tauri).
          Outros hosts de produção precisam entrar na allowlist (
          <code>/embed-partner.json</code> no deploy, ou{" "}
          <code>NEXT_PUBLIC_EMBED_PARTNER_HOSTS</code> em desenvolvimento).
          Peça inclusão em {LEGAL_CONTACT}.
        </p>
      </section>

      <section className="not-prose space-y-3">
        <h2 className="text-lg font-semibold">Host → editor</h2>
        <ul className="space-y-2 text-sm">
          {HOST_MESSAGES.map((row) => (
            <li key={row.type} className="leading-relaxed text-text-secondary">
              <code className="text-foreground">{row.type}</code>
              {" — "}
              {row.detail}
            </li>
          ))}
        </ul>
        <p className="text-xs text-text-secondary">
          Protocolo 1.1.0. Tipos em{" "}
          <code className="text-foreground">utils/embed-protocol.ts</code>.
        </p>
      </section>

      <section className="not-prose space-y-3">
        <h2 className="text-lg font-semibold">Editor → host</h2>
        <ul className="space-y-2 text-sm">
          {EDITOR_MESSAGES.map((row) => (
            <li key={row.type} className="leading-relaxed text-text-secondary">
              <code className="text-foreground">{row.type}</code>
              {" — "}
              {row.detail}
            </li>
          ))}
        </ul>
      </section>

      <section>
        <h2>Exemplo mínimo</h2>
        <pre className="text-xs overflow-x-auto rounded-xl border border-border bg-muted/40 p-4">
          <code>{`const iframe = document.querySelector("iframe");
const origin = "https://editor.app.br";

window.addEventListener("message", (event) => {
  if (event.origin !== origin) return;
  if (event.data?.type === "ready") {
    iframe.contentWindow.postMessage(
      {
        type: "open",
        requestId: crypto.randomUUID(),
        fileName: "contrato.docx",
        fileType: "docx",
        editing: true,
        lang: "pt-BR",
        theme: "dark",
        plugins: "agent",
        bytes: fileBytes,
      },
      origin,
      [fileBytes],
    );
  }
});`}</code>
        </pre>
      </section>

      <section>
        <div className="flex items-center gap-2 mb-3">
          <Puzzle className="w-5 h-5 text-primary" />
          <h2 className="text-lg font-semibold m-0">Plugin Agente</h2>
        </div>
        <p>
          Com <code>plugins: &quot;agent&quot;</code> no <code>open</code>, o
          embed carrega o plugin de sistema (GUID{" "}
          <code>{AGENT_GUID}</code>,{" "}
          <code>/office-plugins/agent/</code>). O host{" "}
          <strong>não</strong> acessa{" "}
          <code>iframe.contentDocument</code>: envia{" "}
          <code>{`{ type: "command", requestId, name, payload }`}</code> e
          recebe <code>commandResult</code>.
        </p>
        <ul className="not-prose space-y-2 text-sm mt-3">
          {AGENT_COMMANDS.map((row) => (
            <li key={row.name} className="leading-relaxed text-text-secondary">
              <code className="text-foreground">{row.name}</code>
              {" — "}
              {row.detail}
            </li>
          ))}
        </ul>
        <p>
          Autocomplete (ghost): o plugin emite{" "}
          <code>completionRequest</code>; o host devolve o texto via{" "}
          <code>command</code> <code>suggest</code> /{" "}
          <code>suggest_update</code>. Preview sem agente:{" "}
          <code>plugins: &quot;none&quot;</code> ou{" "}
          <code>variant: &quot;preview&quot;</code>.
        </p>
      </section>

      <section>
        <h2>Warmup / cache</h2>
        <p>
          Um iframe oculto em <code>/embed?warmup=1</code> faz o navegador
          persistir JS/WASM no origin <code>editor.app.br</code> (Cache
          Storage). Isso não é redistribuição: os arquivos não entram no
          seu app. O HTML de <code>/</code>, <code>/editor</code> e{" "}
          <code>/embed</code> não é cacheado de forma imutável.
        </p>
      </section>

      <p className="not-prose text-sm text-text-secondary">
        Chrome da integração: o embed esconde o avatar dummy e o wordmark
        EDITOR GRATUITO. O editor público em{" "}
        <Link href="/editor" className="text-primary hover:underline">
          /editor
        </Link>{" "}
        mantém a marca. Atribuições:{" "}
        <Link href="/legal" className="text-primary hover:underline">
          /legal
        </Link>
        .
      </p>
    </div>
  );
}
