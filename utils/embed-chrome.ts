const HIDE_PREVIEW_CHROME_CSS = `
  #app-title,
  #header,
  .header,
  [data-layout-name="header"],
  [data-layout-name="toolbar"],
  [data-layout-name="leftMenu"],
  [data-layout-name="statusBar"] {
    display: none !important;
  }
`;

/** Hide dummy Me avatar only. Keep EDITOR GRATUITO logo and File → About (AGPL). */
const HIDE_EMBED_BRAND_CSS = `
  .slot-btn-user-name,
  .btn-current-user,
  .color-user-name,
  .box-cousers,
  .btn-users {
    display: none !important;
  }
`;

function injectStyle(doc: Document, id: string, css: string) {
  if (doc.getElementById(id)) return;
  const style = doc.createElement("style");
  style.id = id;
  style.textContent = css;
  doc.head.appendChild(style);
}

export function injectPreviewChrome(doc: Document | null | undefined) {
  if (!doc) return;
  injectStyle(doc, "ji-embed-hide-chrome", HIDE_PREVIEW_CHROME_CSS);
}

export function injectEmbedChrome(doc: Document | null | undefined) {
  if (!doc) return;
  injectStyle(doc, "ji-embed-hide-brand", HIDE_EMBED_BRAND_CSS);
}
