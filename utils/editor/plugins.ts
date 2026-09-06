export const pluginsBase = "https://plugins.editor.app.br/sdkjs-plugins";

export const allPlugins = [
  "ai",
  "apertium",
  "autocomplete",
  "bergamot",
  "chess",
  "cvbuilder",
  "datepicker",
  "deepl",
  "doc2md",
  "drawio",
  "easybib",
  "glavred",
  "grammalecte",
  "highlightcode",
  "html",
  "icons",
  "idphoto",
  "insertQR",
  "jitsi",
  "languagetool",
  "marketplace",
  "mathpix",
  "mendeley",
  "news",
  "ocr",
  "onlydraw",
  "photoeditor",
  "pixabay",
  "pomodoro",
  "rainbow",
  "speech",
  "speechrecognition",
  "telegram",
  "termef",
  "textcleaner",
  "texthighlighter",
  "thesaurus",
  "translator",
  "typograf",
  "videoembedder",
  "wordpress",
  "wordscounter",
  "youtube",
  "zhipu",
  "zoom",
  "zotero",
];

export const featuredPlugins = [
  "marketplace",
  "ai",
  "youtube",
  "jitsi",
  "photoeditor",
  "typograf",
  "languagetool",
  "thesaurus",
  "deepl",
  "zhipu",
];

export function getPluginConfigUrl(name: string) {
  return `${pluginsBase}/${name}/config.json`;
}

export const AGENT_PLUGIN_GUID = "asc.{7E4A1C90-2B6D-4F11-9A33-8C0E5D71B2A4}";
export const AGENT_PLUGIN_PATH = "/office-plugins/agent/config.json";

/** Keep in sync with public/office-plugins/agent/config.json */
export const AGENT_PLUGIN_MANIFEST = {
  name: "Agent",
  nameLocale: { en: "Agent", "pt-BR": "Agente" },
  guid: AGENT_PLUGIN_GUID,
  version: "1.2.6",
  variations: [
    {
      description: "Document tools for the host Agent",
      url: "index.html",
      icons: ["icon.png", "icon.png"],
      isViewer: false,
      EditorsSupport: ["word", "cell", "slide", "pdf"],
      isVisual: false,
      isModal: false,
      isInsideMode: false,
      // OnlyOffice 9: isSystem → visible=false and the iframe never mounts.
      // type:background → Background Plugins menu, off until the user toggles it
      // (callCommand on a half-started frame crashes the editor).
      // type:invisible maps to PluginType.Invisible (sdk) and autostarts hidden.
      // Note: docs say "unvisible" but live getType() only accepts "invisible".
      isSystem: false,
      type: "invisible",
      initDataType: "none",
      initData: "",
      isUpdateOleOnResize: false,
      buttons: [] as unknown[],
    },
  ],
};

export function isPluginsJsonPath(pathname: string) {
  return pathname === "/plugins.json" || pathname.endsWith("/plugins.json");
}

export function isAgentPluginConfigPath(pathname: string) {
  return pathname === AGENT_PLUGIN_PATH || pathname.endsWith(AGENT_PLUGIN_PATH);
}

export function getPluginsData(list: string[]) {
  return {
    url: "",
    pluginsData: list.map(getPluginConfigUrl),
    autostart: [],
  };
}

export function getAgentPluginsData(origin = "") {
  const base = origin.replace(/\/$/, "");
  return {
    url: "",
    pluginsData: [`${base}${AGENT_PLUGIN_PATH}`],
    autostart: [AGENT_PLUGIN_GUID],
  };
}
