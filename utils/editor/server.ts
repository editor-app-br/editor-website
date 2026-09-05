import { converter } from "./x2t";
import { MockSocket } from "./socket";
import { User, Participant, AscSaveTypes, AvsFileType, ServerOptions } from "./types";
import { emptyDocx, emptyPdf, emptyPptx, emptyXlsx } from "./empty";
import { getDocumentType, getFileExt } from "./utils";
import {
  AGENT_PLUGIN_MANIFEST,
  allPlugins,
  featuredPlugins,
  getAgentPluginsData,
  getPluginsData,
  isAgentPluginConfigPath,
  isPluginsJsonPath,
} from "./plugins";
import {
  isNativeOfficePersistExport,
  isPdfDownloadAs,
  nativeFormatFromExt,
} from "./persistExport";

function mergeBuffers(buffers: Uint8Array[]) {
  const totalLength = buffers.reduce((acc, buffer) => acc + buffer.length, 0);
  const mergedBuffer = new Uint8Array(totalLength);
  let offset = 0;
  for (const buffer of buffers) {
    mergedBuffer.set(buffer, offset);
    offset += buffer.length;
  }
  return mergedBuffer;
}

function randomId() {
  return Math.random().toString(36).substring(2, 9);
}

function getUrl(data: Uint8Array, type?: string) {
  const blob = new Blob([data as Uint8Array<ArrayBuffer>], {
    type: type || "application/octet-stream",
  });
  return URL.createObjectURL(blob);
}

export class EditorServer {
  private id = "";
  private socket: MockSocket | null = null;
  private sessionId: string = "session-id";
  private user: User = {
    id: "uid",
    name: "Me",
  };
  private client = {
    buildVersion: "9.3.0",
    buildNumber: 8,
  };
  private participants: Participant[] = [];
  private syncChangesIndex = 0;
  private loadPromise: Promise<void> | null = null;

  private file: File | null = null;
  private fileType: string = "docx";
  private title: string = "";
  private fsMap: Map<string, Uint8Array> = new Map();
  private urlsMap: Map<string, string> = new Map();

  private downloadId: string = "";
  private downloadParts: Uint8Array[] = [];

  private options: ServerOptions = {};

  constructor(options: ServerOptions = {}) {
    this.options = options;
    this.send = this.send.bind(this);
    this.handleConnect = this.handleConnect.bind(this);
    this.handleMessage = this.handleMessage.bind(this);
  }

  async open(
    file: File,
    { fileType, fileName }: { fileType?: string; fileName?: string } = {},
  ) {
    const title = fileName || file.name;
    this.fileType = fileType || getFileExt(file.name) || "docx";
    const documentType = getDocumentType(this.fileType);
    this.id = randomId();
    this.file = file;
    this.title = title;
    const buffer = await file.arrayBuffer();
    this.loadPromise = this.loadDocument(buffer, this.fileType);
    await this.loadPromise;

    return {
      id: this.id,
      documentType,
    };
  }

  openNew(fileType?: string) {
    this.fileType = fileType || "docx";
    // TODO: should generate new id?
    this.id = this.id || randomId();
    this.title = "New Document";
    const documentType = getDocumentType(this.fileType);

    let binData: Uint8Array | null = null;

    switch (documentType) {
      case "word":
        binData = Uint8Array.from(emptyDocx, (v) => v.charCodeAt(0));
        break;
      case "cell":
        binData = Uint8Array.from(emptyXlsx, (v) => v.charCodeAt(0));
        break;
      case "slide":
        binData = Uint8Array.from(emptyPptx, (v) => v.charCodeAt(0));
        break;
      case "pdf":
        binData = Uint8Array.from(emptyPdf, (v) => v.charCodeAt(0));
        break;
    }

    if (!binData) {
      throw new Error("Failed to create new document");
    }

    this.fsMap.set("Editor.bin", binData);
    this.urlsMap.set("Editor.bin", getUrl(binData));

    return {
      id: this.id,
      documentType: documentType,
    };
  }

  async openUrl(
    url: string,
    {
      fileType,
      fileName,
      loader = (url: string) => fetch(url).then((res) => res.arrayBuffer()),
    }: {
      fileType?: string;
      fileName?: string;
      loader?: (url: string) => Promise<ArrayBuffer>;
    } = {},
  ) {
    const title = fileName || decodeURIComponent(url.split("/").pop() || "Document")
    this.fileType = fileType || getFileExt(title) || "docx";
    const documentType = getDocumentType(this.fileType);
    this.id = randomId();
    this.title = title;
    this.loadPromise = this.loadDocument(() => loader(url), this.fileType);
    await this.loadPromise;

    return {
      id: this.id,
      documentType,
    };
  }

  getDocument() {
    if (!this.id) {
      this.openNew();
    }

    return {
      fileType: this.fileType,
      key: this.id,
      title: this.title,
      url: "/" + this.id,
    };
  }

  getUser() {
    return this.user;
  }

  private async loadDocument(
    buffer: ArrayBuffer | (() => Promise<ArrayBuffer>),
    fileType: string,
  ) {
    if (typeof buffer == "function") {
      buffer = await buffer();
    }

    let output: Uint8Array | null = null;
    let media: { [key: string]: Uint8Array } = {};

    if (fileType == "pdf") {
      output = new Uint8Array(buffer);
    } else {
      const result = await converter.convert({
        data: buffer,
        fileFrom: "doc." + fileType,
        fileTo: "Editor.bin",
      });
      output = result.output;
      media = result.media;
    }

    if (!output) {
      throw new Error("Failed to convert file");
    }

    if (this.urlsMap.size > 0) {
      this.urlsMap.forEach((url) => URL.revokeObjectURL(url));
    }
    this.fsMap.set("Editor.bin", output);
    this.urlsMap.set("Editor.bin", getUrl(output));
    for (const name in media) {
      this.addMedia(name, media[name]);
    }
  }

  private addMedia(name: string, data: Uint8Array) {
    const pathname = "media/" + name;
    const url = getUrl(data);
    this.fsMap.set(pathname, data);
    this.urlsMap.set(pathname, url);
    return url;
  }

  setClient(info: Partial<typeof this.client>) {
    this.client = {
      ...this.client,
      ...info,
    };
  }

  handleConnect({ socket }: { socket: MockSocket }) {
    console.log("connect: ", socket);

    this.socket = socket;
    const { send, sessionId, client } = this;

    this.participants = [
      {
        connectionId: this.sessionId,
        encrypted: false,
        id: this.user.id,
        idOriginal: this.user.id,
        indexUser: 1,
        isCloseCoAuthoring: false,
        isLiveViewer: false,
        username: this.user.name,
        view: false,
      },
    ];

    socket.server.on("message", this.handleMessage);

    send({
      maxPayload: 100000000,
      pingInterval: 25000,
      pingTimeout: 20000,
      sid: sessionId,
      upgrades: [],
    });

    send({
      type: "license",
      license: {
        type: 3,
        buildNumber: client.buildNumber,
        buildVersion: client.buildVersion,
        light: false,
        mode: 0,
        rights: 1,
        protectionSupport: true,
        isAnonymousSupport: true,
        liveViewerSupport: true,
        branding: false,
        customization: true,
        advancedApi: false,
      },
    });
  }

  handleDisconnect({ socket }: { socket: MockSocket }) {
    console.log("disconnect: ", socket);
    this.socket = null;
  }

  send(...msg: unknown[]) {
    if (!this.socket) {
      console.error("Socket is not connected");
      return;
    }
    console.log("[ws] >> ", ...msg);
    this.socket.server.emit("message", ...msg);
  }

  async handleMessage(msg: Record<string, string>, ...args: unknown[]) {
    console.log("[ws] << ", msg, args);

    const { send, sessionId, participants, user, client } = this;
    const type =
      typeof msg === "object" && msg && "type" in msg ? msg.type : null;
    switch (type) {
      case "auth":
        const changes: unknown[] = [];
        send({
          type: "authChanges",
          changes: changes,
        });
        send({
          type: "auth",
          result: 1,
          sessionId: sessionId,
          participants: participants,
          locks: [],
          //   changes: changes,
          //   changesIndex: 0,
          indexUser: 1,
          buildVersion: client.buildVersion || "9.3.0",
          buildNumber: client.buildNumber || 9,
          licenseType: 3,
          editorType: 2,
          mode: "edit",
          permissions: {
            comment: true,
            chat: true,
            download: true,
            edit: true,
            fillForms: false,
            modifyFilter: true,
            protect: true,
            print: true,
            review: false,
            copy: true,
          },
        });

        try {
          if (this.loadPromise) {
            await this.loadPromise;
          }
          send({
            type: "documentOpen",
            data: {
              type: "open",
              status: "ok",
              data: {
                ...Object.fromEntries(this.urlsMap),
              },
            },
          });
        } catch (err) {
          console.error(err);
          // TODO: send error message
          send({
            type: "documentOpen",
            data: {
              type: "open",
              status: "ok",
              data: {
                "Editor.bin": "",
              },
            },
          });
        }
        break;
      case "isSaveLock":
        send({
          type: "saveLock",
          saveLock: false,
        });
        break;
      case "saveChanges":
        send({
          type: "unSaveLock",
          index: -1,
          syncChangesIndex: ++this.syncChangesIndex,
          time: +new Date(),
        });
        break;
      case "getLock":
        send({
          type: "getLock",
          locks: {
            [msg.block]: {
              time: +new Date(),
              user: user?.id,
              block: msg.block,
            },
          },
        });
        send({
          type: "releaseLock",
          locks: {
            [msg.block]: {
              time: +new Date(),
              user: user?.id,
              block: msg.block,
            },
          },
        });
        break;
    }
  }

  async handleRequest(req: Request) {
    const u = new URL(req.url);

    const { id: key, send } = this;

    if (u.pathname.includes("/downloadas/")) {
      try {
        const cmd = JSON.parse(u.searchParams.get("cmd") || "{}") as {
          title?: string;
          outputformat?: number;
          savetype?: number | string;
          format?: string;
        };
        const buffer = await req.arrayBuffer();

        const title = cmd.title || this.title || `document.${this.fileType}`;
        const titleExt = (title.split(".").pop() || this.fileType || "docx").toLowerCase();
        const nativePersist = isNativeOfficePersistExport(this.fileType, cmd);
        const fileTo = nativePersist ? `doc.${this.fileType}` : `doc.${titleExt}`;
        let formatTo = cmd.outputformat ?? nativeFormatFromExt(nativePersist ? this.fileType : titleExt);
        if (!formatTo && fileTo.endsWith(".pdf")) {
          formatTo = AvsFileType.AVS_FILE_CROSSPLATFORM_PDF;
        }

        const media = Object.fromEntries(
          [...this.fsMap.entries()].filter(([name]) => name.startsWith("media/")),
        );

        const download = async (parts: Uint8Array[]) => {
          try {
            if (!nativePersist && isPdfDownloadAs(cmd)) {
              return { status: "ok" };
            }

            const input = mergeBuffers(parts);
            let fileFrom = "from.bin";
            if (cmd.format == "pdf") {
              fileFrom = "from.pdf";
            }

            const { output } = await converter.convert({
              data: input.buffer.slice(0),
              fileFrom,
              fileTo,
              formatTo,
              media,
            });
            if (!output) {
              console.error("Conversion failed");
              // Host persist: never show OnlyOffice "Baixar como" for a failed WASM save.
              return nativePersist && this.options.persistFile ? { status: "ok" } : { status: "error" };
            }
            const blob = new Blob([new Uint8Array(output)]);
            if (nativePersist && this.options.persistFile) {
              try {
                await this.options.persistFile(blob, title);
              } catch (persistErr) {
                console.error(persistErr);
              }
              return { status: "ok" };
            }
            if (this.options.onExportedFile) {
              try {
                await this.options.onExportedFile(blob, title);
              } catch (exportErr) {
                console.error(exportErr);
              }
              return { status: "ok" };
            }
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = title;
            a.click();
            URL.revokeObjectURL(url);

            return { status: "ok" };
          } catch (err) {
            console.error(err);
            return nativePersist && this.options.persistFile ? { status: "ok" } : { status: "error" };
          } finally {
            if (nativePersist) this.options.onExportFinished?.();
          }
        };

        let result = { status: "ok" };
        const saveType = Number(cmd.savetype);
        const kind = Number.isFinite(saveType) ? saveType : AscSaveTypes.CompleteAll;
        const skipPdfExport = !nativePersist && isPdfDownloadAs(cmd);
        if (skipPdfExport) {
          if (kind === AscSaveTypes.Complete || kind === AscSaveTypes.CompleteAll) {
            result = { status: "ok" };
          }
        } else {
          switch (kind) {
            case AscSaveTypes.PartStart:
              this.downloadId = "_" + Math.round(Math.random() * 1000);
              this.downloadParts = [new Uint8Array(buffer)];
              break;
            case AscSaveTypes.Part:
              this.downloadParts.push(new Uint8Array(buffer));
              break;
            case AscSaveTypes.Complete: {
              const parts = [...this.downloadParts, new Uint8Array(buffer)];
              this.downloadParts = [];
              result = await download(parts);
              break;
            }
            case AscSaveTypes.CompleteAll:
            default:
              result = await download([new Uint8Array(buffer)]);
              break;
          }
        }

        setTimeout(() => {
          send({
            type: "documentOpen",
            data: {
              type: "save",
              status: result.status,
              data: "data:,",
              filetype: this.fileType || "docx",
            },
          });
        }, 100);

        return Response.json({
          status: result.status,
          type: "save",
          data: this.downloadId,
        });
      } catch (err) {
        console.error(err);
        // A thrown handler falls through to native fetch (404) and OnlyOffice
        // shows "Use a opção Baixar como". Always ack the mocked save.
        return Response.json({
          status: this.options.persistFile ? "ok" : "error",
          type: "save",
          data: this.downloadId,
        });
      }
    }

    if (u.pathname.endsWith("/upload/" + key)) {
      const buffer = await req.arrayBuffer();
      const data = new Uint8Array(buffer);
      const filename = Date.now() + ".png";
      const pathname = "media/" + filename;
      const url = this.addMedia(filename, data);
      return Response.json({ [pathname]: url });
    }

    // DocsAPI loads ../../../../plugins.json from /v9.3.1-2/web-apps/.../main/
    // so the path is /v9.3.1-2/plugins.json, not /plugins.json.
    if (isPluginsJsonPath(u.pathname)) {
      const state = this.options.getState?.();
      if (state?.plugins == "none") {
        return Response.json({ url: "", pluginsData: [], autostart: [] });
      }
      if (state?.plugins == "agent") {
        return Response.json(getAgentPluginsData(location.origin));
      }
      if (state?.plugins == "all") {
        return Response.json(getPluginsData(allPlugins));
      }
      return Response.json(getPluginsData(featuredPlugins));
    }

    if (isAgentPluginConfigPath(u.pathname)) {
      const baseUrl = `${location.origin}/office-plugins/agent/`;
      return Response.json({ ...AGENT_PLUGIN_MANIFEST, baseUrl });
    }

    return null;
  }
}
