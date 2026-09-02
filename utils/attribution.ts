/** Upstream credits and source URLs (AGPL §13 / legal notice). */

export const PUBLIC_WEBSITE_REPO = "https://github.com/editor-app-br/editor-website";

export const PROJECT_REPOS = {
  website: {
    label: "editor-website (site + editor shell)",
    web: PUBLIC_WEBSITE_REPO,
    git: `${PUBLIC_WEBSITE_REPO}.git`,
  },
  extension: {
    label: "editor-google-app (Chrome extension)",
    web: "https://github.com/editor-app-br/editor-google-app",
    git: "https://github.com/editor-app-br/editor-google-app.git",
  },
  deploy: {
    label: "deploy (Docker / Helm)",
    web: "https://gitlab.bra.prata.dev.br/editor/deploy",
    git: "https://gitlab.bra.prata.dev.br/editor/deploy.git",
  },
  plugins: {
    label: "editor-plugins (ONLYOFFICE plugins + Marketplace)",
    web: "https://github.com/editor-app-br/editor-plugins",
    git: "https://github.com/editor-app-br/editor-plugins.git",
  },
} as const;

export const UPSTREAM_COMPONENTS = [
  {
    name: "ZIZIYI office-website",
    author: "© baotlake (ZIZIYI)",
    license: "GNU AGPL v3",
    contribution:
      "Local client architecture, in-browser socket emulation, and WebAssembly orchestration.",
    url: "https://github.com/baotlake/office-website",
  },
  {
    name: "ONLYOFFICE DocumentServer",
    author: "© Ascensio System SIA",
    license: "GNU AGPL v3 (+ brand/UI additional terms)",
    contribution: "Office editor engine (Community Edition web-apps / sdkjs).",
    url: "https://www.onlyoffice.com",
  },
  {
    name: "CryptPad onlyoffice-x2t-wasm",
    author: "© XWiki SAS / CryptPad team",
    license: "GNU AGPL v3",
    contribution: "x2t document conversion compiled to WebAssembly.",
    url: "https://github.com/cryptpad/onlyoffice-x2t-wasm",
  },
] as const;

export const LEGAL_CONTACT = "contato@prata.dev.br";
