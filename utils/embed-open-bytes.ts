function base64ToArrayBuffer(b64: string): ArrayBuffer | null {
  try {
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i += 1) {
      out[i] = bin.charCodeAt(i);
    }
    return out.buffer;
  } catch {
    return null;
  }
}

/** Resolve document bytes from a host message (ArrayBuffer or base64 fallback for WKWebView). */
export function hostOpenBytes(data: { bytes?: unknown; bytesB64?: string }): ArrayBuffer | null {
  const row = data as { bytes?: unknown; bytesB64?: string };
  if (row.bytes instanceof ArrayBuffer && row.bytes.byteLength > 0) {
    return row.bytes;
  }
  if (ArrayBuffer.isView(row.bytes) && row.bytes.byteLength > 0) {
    const view = row.bytes as ArrayBufferView;
    const slice = new Uint8Array(view.byteLength);
    slice.set(new Uint8Array(view.buffer, view.byteOffset, view.byteLength));
    return slice.buffer;
  }
  if (typeof row.bytesB64 === "string" && row.bytesB64.length > 0) {
    return base64ToArrayBuffer(row.bytesB64);
  }
  return null;
}
