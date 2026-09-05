export interface XHRMiddleware {
  (request: Request): Response | null | Promise<Response | null>;
}

export type SyncXHRMiddleware = (
  url: string,
  method: string,
) => { status?: number; statusText?: string; body: string; contentType?: string } | null;

/**
 * Creates an XMLHttpRequest proxy class that supports middleware
 * @param BaseXHR The original XMLHttpRequest class
 * @returns The enhanced XMLHttpRequest class
 */
export function createXHRProxy(
  BaseXHR = globalThis.XMLHttpRequest,
  scope: Pick<typeof globalThis, "Request"> = globalThis,
) {
  return class ProxyXMLHttpRequest extends BaseXHR {
    private static _middlewares: XHRMiddleware[] = [];
    private static _syncMiddlewares: SyncXHRMiddleware[] = [];

    private _isMocked: boolean = false;
    private _requestMethod: string = "GET";
    private _requestUrl: string = "";
    private _requestHeaders: Headers = new Headers();
    private _requestBody: any = null;
    private _async: boolean = true;

    /**
     * Register global middleware
     */
    static use(middleware: XHRMiddleware) {
      this._middlewares.push(middleware);
    }

    /** Sync XHR (OnlyOffice getConfigJson) cannot await async middleware. */
    static useSync(middleware: SyncXHRMiddleware) {
      this._syncMiddlewares.push(middleware);
    }

    /**
     * Clear all middleware
     */
    static clearMiddlewares() {
      this._middlewares.length = 0;
      this._syncMiddlewares.length = 0;
    }

    open(
      method: string,
      url: string | URL,
      async: boolean = true,
      username?: string | null,
      password?: string | null,
    ): void {
      this._requestMethod = method;
      this._requestUrl = url.toString();
      this._requestHeaders = new Headers();
      this._isMocked = false;
      this._async = async !== false;

      // Call native open
      super.open(
        method,
        url,
        async,
        username ?? undefined,
        password ?? undefined,
      );
    }

    setRequestHeader(name: string, value: string): void {
      this._requestHeaders.append(name, value);

      // If it is not a mock request, also set it on the native XHR
      if (!this._isMocked) {
        super.setRequestHeader(name, value);
      }
    }

    send(body?: Document | XMLHttpRequestBodyInit | null): void {
      this._requestBody = body;

      // OnlyOffice plugin bootstrap uses sync XHR (getConfigJson).
      // Async middleware would return empty responseText immediately, so serve
      // known plugin paths synchronously instead of falling through to 404.
      if (!this._async) {
        for (const mw of ProxyXMLHttpRequest._syncMiddlewares) {
          try {
            const mocked = mw(this._requestUrl, this._requestMethod);
            if (mocked) {
              this._isMocked = true;
              this._applySyncMock(mocked);
              return;
            }
          } catch (err) {
            console.error("ProxyXMLHttpRequest sync middleware error:", err);
          }
        }
        super.send(body);
        return;
      }

      // Try to run middleware
      this._tryMiddlewares()
        .then((handled) => {
          if (!handled) {
            // No middleware handled it, use native send
            super.send(body);
          }
        })
        .catch((err) => {
          console.error("ProxyXMLHttpRequest middleware error:", err);
          // Fallback to native implementation on error
          super.send(body);
        });
    }

    private _applySyncMock(mocked: {
      status?: number;
      statusText?: string;
      body: string;
      contentType?: string;
    }) {
      const status = mocked.status ?? 200;
      const statusText = mocked.statusText ?? "OK";
      const body = mocked.body;
      let responseData: any = body;
      if (this.responseType === "json") {
        try {
          responseData = JSON.parse(body);
        } catch {
          responseData = null;
        }
      }

      this.dispatchEvent(new ProgressEvent("loadstart"));
      Object.defineProperty(this, "readyState", {
        value: 2,
        writable: false,
        configurable: true,
      });
      this.dispatchEvent(new Event("readystatechange"));
      Object.defineProperty(this, "readyState", {
        value: 3,
        writable: false,
        configurable: true,
      });
      this.dispatchEvent(new Event("readystatechange"));

      Object.defineProperty(this, "status", {
        value: status,
        writable: false,
        configurable: true,
      });
      Object.defineProperty(this, "statusText", {
        value: statusText,
        writable: false,
        configurable: true,
      });
      Object.defineProperty(this, "response", {
        value: responseData,
        writable: false,
        configurable: true,
      });
      Object.defineProperty(this, "responseText", {
        value: body,
        writable: false,
        configurable: true,
      });
      Object.defineProperty(this, "responseURL", {
        value: this._requestUrl,
        writable: false,
        configurable: true,
      });
      Object.defineProperty(this, "getResponseHeader", {
        value: (name: string) => {
          if (name.toLowerCase() === "content-type") {
            return mocked.contentType || "application/json";
          }
          return null;
        },
        writable: false,
        configurable: true,
      });

      this.dispatchEvent(
        new ProgressEvent("progress", {
          lengthComputable: true,
          loaded: body.length,
          total: body.length,
        }),
      );
      Object.defineProperty(this, "readyState", {
        value: 4,
        writable: false,
        configurable: true,
      });
      this.dispatchEvent(new Event("readystatechange"));
      this.dispatchEvent(new ProgressEvent("load"));
      this.dispatchEvent(new ProgressEvent("loadend"));
    }

    private async _tryMiddlewares(): Promise<boolean> {
      // Create Request object
      let request: Request;
      try {
        const reqInit: RequestInit = {
          method: this._requestMethod,
          headers: this._requestHeaders,
          body: this._requestBody as BodyInit,
          mode: "cors",
        };

        if (this.withCredentials) {
          reqInit.credentials = "include";
        }

        request = new scope.Request(this._requestUrl, reqInit);
      } catch (e) {
        // Unable to create Request, do not use middleware
        return false;
      }

      // Run middleware
      for (const mw of ProxyXMLHttpRequest._middlewares) {
        const response = await mw(request.clone());
        if (response) {
          this._isMocked = true;
          await this._handleMockResponse(response);
          return true;
        }
      }

      return false;
    }

    private async _handleMockResponse(response: Response) {
      // 1. Trigger loadstart
      this.dispatchEvent(new ProgressEvent("loadstart"));

      // 2. HEADERS_RECEIVED (readyState = 2)
      Object.defineProperty(this, "readyState", {
        value: 2,
        writable: false,
        configurable: true,
      });
      this.dispatchEvent(new Event("readystatechange"));

      // 3. LOADING (readyState = 3)
      Object.defineProperty(this, "readyState", {
        value: 3,
        writable: false,
        configurable: true,
      });
      this.dispatchEvent(new Event("readystatechange"));

      try {
        // Read response body
        let responseData: any;

        if (this.responseType === "json") {
          responseData = await response.json();
        } else if (this.responseType === "arraybuffer") {
          responseData = await response.arrayBuffer();
        } else if (this.responseType === "blob") {
          responseData = await response.blob();
        } else if (this.responseType === "document") {
          const text = await response.text();
          responseData = new DOMParser().parseFromString(text, "text/xml");
        } else {
          responseData = await response.text();
        }

        // Set response properties
        Object.defineProperty(this, "status", {
          value: response.status,
          writable: false,
          configurable: true,
        });

        Object.defineProperty(this, "statusText", {
          value: response.statusText,
          writable: false,
          configurable: true,
        });

        Object.defineProperty(this, "response", {
          value: responseData,
          writable: false,
          configurable: true,
        });

        Object.defineProperty(this, "responseText", {
          value:
            typeof responseData === "string"
              ? responseData
              : JSON.stringify(responseData),
          writable: false,
          configurable: true,
        });

        Object.defineProperty(this, "responseURL", {
          value: response.url,
          writable: false,
          configurable: true,
        });

        // 4. Trigger progress event
        this.dispatchEvent(
          new ProgressEvent("progress", {
            lengthComputable: true,
            loaded: 100,
            total: 100,
          }),
        );

        // 5. DONE (readyState = 4)
        Object.defineProperty(this, "readyState", {
          value: 4,
          writable: false,
          configurable: true,
        });
        this.dispatchEvent(new Event("readystatechange"));

        // 6. Trigger load event
        this.dispatchEvent(new ProgressEvent("load"));

        // 7. Trigger loadend event
        this.dispatchEvent(new ProgressEvent("loadend"));
      } catch (e) {
        console.error("ProxyXHR: error handling response", e);

        // Set readyState to DONE
        Object.defineProperty(this, "readyState", {
          value: 4,
          writable: false,
          configurable: true,
        });
        this.dispatchEvent(new Event("readystatechange"));

        // Trigger error event
        this.dispatchEvent(new ProgressEvent("error"));
        this.dispatchEvent(new ProgressEvent("loadend"));
      }
    }
  };
}
