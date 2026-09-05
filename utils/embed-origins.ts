/**
 * Origins allowed to drive /embed via postMessage.
 *
 * First-party: this site + local/desktop hosts.
 * Partners: runtime /embed-partner.json (Helm) and optional NEXT_PUBLIC_* for local .env.
 */

type PartnerAllowlist = {
  exactOrigins: Set<string>;
  hosts: Set<string>;
  suffixes: string[];
};

const FIRST_PARTY_ORIGINS = new Set([
  "https://editor.app.br",
  "https://www.editor.app.br",
  "https://editor.com.br",
  "https://www.editor.com.br",
  "https://workspace.jusintegra.com.br",
  "https://www.workspace.jusintegra.com.br",
  "https://local.ji.app.br",
  "http://localhost:5173",
  "http://127.0.0.1:5173",
  "http://localhost:3000",
  "http://127.0.0.1:3000",
  "https://tauri.localhost",
  "http://tauri.localhost",
  "tauri://localhost",
  "https://asset.localhost",
  "asset://localhost",
]);

const FIRST_PARTY_SUFFIXES = [".jusintegra.com.br", ".ji.app.br"];

const EMBED_PARTNER_JSON = "/embed-partner.json";

function splitCsv(raw: string | undefined): string[] {
  return (raw || "")
    .split(/[,;\s]+/)
    .map((item) => item.trim())
    .filter(Boolean);
}

export function parsePartnerAllowlist(
  hostsRaw?: string,
  partnerUrl?: string,
): PartnerAllowlist {
  const exactOrigins = new Set<string>();
  const hosts = new Set<string>();
  const suffixes: string[] = [];

  const tokens = splitCsv(hostsRaw);
  if (partnerUrl) tokens.push(partnerUrl);

  for (const token of tokens) {
    if (token.startsWith("*.")) {
      suffixes.push(token.slice(1).toLowerCase());
      continue;
    }
    try {
      const url = token.includes("://") ? new URL(token) : new URL(`https://${token}`);
      exactOrigins.add(url.origin);
      hosts.add(url.hostname.toLowerCase());
    } catch {
      hosts.add(token.toLowerCase());
    }
  }

  return { exactOrigins, hosts, suffixes };
}

function mergeAllowlists(left: PartnerAllowlist, right: PartnerAllowlist): PartnerAllowlist {
  return {
    exactOrigins: new Set([...left.exactOrigins, ...right.exactOrigins]),
    hosts: new Set([...left.hosts, ...right.hosts]),
    suffixes: [...new Set([...left.suffixes, ...right.suffixes])],
  };
}

function envAllowlist(): PartnerAllowlist {
  return parsePartnerAllowlist(
    process.env.NEXT_PUBLIC_EMBED_PARTNER_HOSTS,
    process.env.NEXT_PUBLIC_EMBED_PARTNER_URL,
  );
}

let cached: PartnerAllowlist = envAllowlist();
let partnerUrl: string | null = (process.env.NEXT_PUBLIC_EMBED_PARTNER_URL || "").trim() || null;
let loadPromise: Promise<void> | null = null;

export function embedPartnerUrl(): string | null {
  return partnerUrl;
}

/** Load Helm/runtime partner file. Safe to call more than once. */
export function loadEmbedPartnerConfig(): Promise<void> {
  if (loadPromise) return loadPromise;
  loadPromise = (async () => {
    try {
      const response = await fetch(EMBED_PARTNER_JSON, { cache: "no-store" });
      if (!response.ok) return;
      const data = (await response.json()) as { hosts?: unknown; url?: unknown };
      const hosts = typeof data.hosts === "string" ? data.hosts : "";
      const url = typeof data.url === "string" ? data.url.trim() : "";
      cached = mergeAllowlists(envAllowlist(), parsePartnerAllowlist(hosts, url));
      if (url) partnerUrl = url;
    } catch {
      /* local Next without the file, or offline — keep env allowlist */
    }
  })();
  return loadPromise;
}

function hostMatchesSuffix(host: string, suffix: string): boolean {
  return host === suffix.replace(/^\./, "") || host.endsWith(suffix);
}

export function isAllowedEmbedHostOrigin(origin: string): boolean {
  if (!origin || origin === "null") return false;
  if (FIRST_PARTY_ORIGINS.has(origin)) return true;

  try {
    const url = new URL(origin);
    if (url.hostname === "localhost" || url.hostname === "127.0.0.1") {
      return url.protocol === "http:" || url.protocol === "https:";
    }
    const host = url.hostname.toLowerCase();
    if (FIRST_PARTY_SUFFIXES.some((suffix) => hostMatchesSuffix(host, suffix))) return true;
    if (cached.exactOrigins.has(url.origin) || cached.hosts.has(host)) return true;
    return cached.suffixes.some((suffix) => hostMatchesSuffix(host, suffix));
  } catch {
    return origin.startsWith("tauri://") || origin.startsWith("asset://");
  }
}
