/**
 * X2T converter. Prefer /x2t-worker.js; fall back to the main thread.
 *
 * Next static export copies `x2t.worker.ts` as media (linguist MIME + nosniff).
 * Chrome refuses that Worker. Production must load a real JS worker.
 */

import { converter as mainThreadConverter } from "./x2t.main";
import { X2tConvertParams, X2tConvertResult } from "./types";

const WORKER_URL = "/x2t-worker.js";
const WORKER_CONVERT_MS = 45_000;

interface PendingMessage {
  resolve: (value: X2tConvertResult) => void;
  reject: (error: Error) => void;
}

interface WorkerResponse {
  id?: number;
  type: string;
  payload?: X2tConvertResult;
  error?: string;
}

function fileExt(name: string) {
  const base = name.split("/").pop() || name;
  const dot = base.lastIndexOf(".");
  return dot >= 0 ? base.slice(dot + 1) : base;
}

async function convertOnMain(params: X2tConvertParams): Promise<X2tConvertResult> {
  const result = await mainThreadConverter.convert(
    params.data,
    fileExt(params.fileFrom),
    fileExt(params.fileTo),
  );
  const media: X2tConvertResult["media"] = {};
  for (const row of result.medias || []) {
    media[row.name] = row.data as Uint8Array<ArrayBuffer>;
  }
  let output: X2tConvertResult["output"] = null;
  if (result.output) {
    output =
      result.output instanceof Uint8Array
        ? (result.output as Uint8Array<ArrayBuffer>)
        : new Uint8Array(result.output);
  }
  return { output, media };
}

export class X2tConverter {
  private worker: Worker | null = null;
  private initPromise: Promise<void> | null = null;
  private useMain = false;
  private messageId = 0;
  private pendingMessages = new Map<number, PendingMessage>();

  constructor() {
    if (typeof globalThis.Worker === "function") {
      void this.init();
    } else {
      this.useMain = true;
    }
  }

  private getNextId(): number {
    this.messageId += 1;
    return this.messageId;
  }

  private handleWorkerMessage = (event: MessageEvent<WorkerResponse>) => {
    const { id, type, payload, error } = event.data;
    if (type === "ready") return;
    if (id == null) return;
    const pending = this.pendingMessages.get(id);
    if (!pending) return;
    this.pendingMessages.delete(id);
    if (type === "error") {
      pending.reject(new Error(error || "Unknown worker error"));
      return;
    }
    pending.resolve(payload as X2tConvertResult);
  };

  private failWorker(message: string) {
    this.useMain = true;
    for (const pending of this.pendingMessages.values()) {
      pending.reject(new Error(message));
    }
    this.pendingMessages.clear();
    try {
      this.worker?.terminate();
    } catch {
      /* already dead */
    }
    this.worker = null;
  }

  private handleWorkerError = (error: ErrorEvent) => {
    console.error("[X2tConverter] Worker error:", error);
    this.failWorker(`Worker error: ${error.message}`);
  };

  public init(): Promise<void> {
    if (this.initPromise) return this.initPromise;

    this.initPromise = new Promise<void>((resolve) => {
      try {
        this.worker = new Worker(WORKER_URL);
        this.worker.onmessage = this.handleWorkerMessage;
        this.worker.onerror = this.handleWorkerError;
        resolve();
      } catch (err) {
        console.warn("[X2tConverter] Worker create failed; using main-thread x2t", err);
        this.useMain = true;
        resolve();
      }
    });

    return this.initPromise;
  }

  private sendMessage(payload: X2tConvertParams): Promise<X2tConvertResult> {
    return new Promise((resolve, reject) => {
      if (!this.worker) {
        reject(new Error("Worker not initialized"));
        return;
      }
      const id = this.getNextId();
      const timer = window.setTimeout(() => {
        this.pendingMessages.delete(id);
        reject(new Error("x2t worker convert timed out"));
      }, WORKER_CONVERT_MS);
      this.pendingMessages.set(id, {
        resolve: (value) => {
          window.clearTimeout(timer);
          resolve(value);
        },
        reject: (error) => {
          window.clearTimeout(timer);
          reject(error);
        },
      });
      if (payload.data instanceof ArrayBuffer) {
        this.worker.postMessage({ id, type: "convert", payload }, [payload.data]);
      } else {
        this.worker.postMessage({ id, type: "convert", payload });
      }
    });
  }

  public async convert(params: X2tConvertParams): Promise<X2tConvertResult> {
    await this.init();
    if (this.useMain || !this.worker) {
      return convertOnMain(params);
    }
    try {
      const dataClone = params.data.slice(0);
      return await this.sendMessage({ ...params, data: dataClone });
    } catch (err) {
      console.warn("[X2tConverter] Worker convert failed; using main-thread x2t", err);
      this.useMain = true;
      return convertOnMain(params);
    }
  }

  public terminate(): void {
    this.failWorker("Worker terminated");
    this.initPromise = null;
  }

  public get isInitialized(): boolean {
    return this.useMain || this.worker !== null;
  }
}

export const converter = new X2tConverter();
