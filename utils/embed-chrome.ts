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

export function injectPreviewChrome(doc: Document | null | undefined) {
  if (!doc) return;
  if (doc.getElementById("ji-embed-hide-chrome")) return;
  const style = doc.createElement("style");
  style.id = "ji-embed-hide-chrome";
  style.textContent = HIDE_PREVIEW_CHROME_CSS;
  doc.head.appendChild(style);
}
