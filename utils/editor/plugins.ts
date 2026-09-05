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
