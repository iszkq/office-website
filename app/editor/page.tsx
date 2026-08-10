"use client";

import { useLayoutEffect, useRef, useEffect, useState } from "react";
import { useAppStore, useResolvedLanguage, useHasHydrated } from "@/store";
import {
  API_JS,
  APP_ROOT,
  getDocumentType,
  PRELOAD_HTML,
} from "@/utils/editor/utils";
import io, { MockSocket } from "@/utils/editor/socket";
import { createFetchProxy } from "@/utils/editor/fetch";
import { createXHRProxy } from "@/utils/editor/xhr";
import { DocEditor } from "@/utils/editor/types";
import { createExtensionLoader } from "@/utils/extension";
import InstallExtensionDialog from "@/components/install-extension-dialog";

const BRIDGE_READY = "xinghuo-office-ready";
const BRIDGE_OPEN = "xinghuo-office-open";
const BRIDGE_OPENED = "xinghuo-office-opened";
const BRIDGE_DIRTY = "xinghuo-office-dirty";
const BRIDGE_SAVE = "xinghuo-office-save";
const BRIDGE_SAVING = "xinghuo-office-saving";
const BRIDGE_SAVED = "xinghuo-office-saved";
const BRIDGE_ERROR = "xinghuo-office-error";

const MIME_TYPES: Record<string, string> = {
  doc: "application/msword",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xls: "application/vnd.ms-excel",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ppt: "application/vnd.ms-powerpoint",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  pdf: "application/pdf",
};

export default function Page() {
  const server = useAppStore((state) => state.server);
  const language = useResolvedLanguage();
  const theme = useAppStore((state) => state.theme);
  const hasHydrated = useHasHydrated();
  const isDirty = useRef(false);
  const [showInstallHint, setShowInstallHint] = useState(false);
  const tryDirectRef = useRef<(() => Promise<void>) | null>(null);

  useEffect(() => {
    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      if (isDirty.current) {
        e.preventDefault();
        e.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", handleBeforeUnload);
    return () => {
      window.removeEventListener("beforeunload", handleBeforeUnload);
    };
  }, []);

  useLayoutEffect(() => {
    if (!hasHydrated) return;

    const apiUrl = APP_ROOT + API_JS;
    const searchParams = new URLSearchParams(window.location.search);

    const fileId = searchParams.get("fileId");
    const newDoc = searchParams.get("new");
    const fileUrl = searchParams.get("url");
    const paramEditing = searchParams.get("editing");
    const paramLang = searchParams.get("lang");
    const paramTheme = searchParams.get("theme");
    const embedded = searchParams.get("embed") === "1";
    const requestId = searchParams.get("requestId");
    const requestedParentOrigin = searchParams.get("parentOrigin");

    let parentOrigin: string | null = null;
    if (embedded && requestedParentOrigin && requestId) {
      try {
        const parsedOrigin = new URL(requestedParentOrigin).origin;
        if (parsedOrigin === requestedParentOrigin) parentOrigin = parsedOrigin;
      } catch {
        parentOrigin = null;
      }
    }
    const bridgeEnabled = Boolean(parentOrigin && requestId);

    const editing = paramEditing === null ? true : paramEditing !== "0";
    const lang = paramLang || language;
    const uiTheme = paramTheme || theme;

    let editor: DocEditor | null = null;
    let saveInProgress = false;

    const postBridgeMessage = (
      type: string,
      payload: Record<string, unknown> = {},
      transfer: Transferable[] = []
    ) => {
      if (!bridgeEnabled || !parentOrigin || !requestId) return;
      window.parent.postMessage(
        { type, requestId, ...payload },
        parentOrigin,
        transfer
      );
    };

    const requestExport = () => {
      if (!bridgeEnabled || !editor || saveInProgress) return;
      saveInProgress = true;
      postBridgeMessage(BRIDGE_SAVING);
      try {
        editor.downloadAs(server.getDocument().fileType);
      } catch (error) {
        saveInProgress = false;
        console.error("Failed to save embedded document", error);
        postBridgeMessage(BRIDGE_ERROR, {
          message: "Failed to generate the updated document",
        });
      }
    };

    server.setDownloadHandler(({ data, fileName, fileType }) => {
      if (!bridgeEnabled || !saveInProgress) return false;
      saveInProgress = false;
      isDirty.current = false;
      postBridgeMessage(
        BRIDGE_SAVED,
        {
          buffer: data,
          fileName,
          mimeType:
            MIME_TYPES[fileType.toLowerCase()] || "application/octet-stream",
        },
        [data]
      );
      return true;
    });

    MockSocket.on("connect", server.handleConnect);
    MockSocket.on("disconnect", server.handleDisconnect);

    const onAppReady = () => {
      const iframe = document.querySelector<HTMLIFrameElement>(
        'iframe[name="frameEditor"]'
      );
      const win = iframe?.contentWindow as typeof window;
      const iframeDoc = iframe?.contentDocument;
      if (!iframeDoc || !win) {
        throw new Error("Iframe not loaded");
      }

      const xhr = createXHRProxy(win.XMLHttpRequest);
      const fetchProxy = createFetchProxy(win);
      const _Worker = win.Worker;

      xhr.use((request: Request) => {
        return server.handleRequest(request);
      });
      fetchProxy.use((request: Request) => {
        return server.handleRequest(request);
      });
      Object.assign(win, {
        io: io,
        XMLHttpRequest: xhr,
        fetch: fetchProxy,
        Worker: function (url: string, options?: WorkerOptions) {
          const u = new URL(url, location.origin);
          return new _Worker(
            u.href.replace(u.origin, location.origin),
            options
          );
        },
      });

      // const script = iframeDoc.createElement("script");
      // script.src = apiUrl;
      // iframeDoc.body.appendChild(script);
    };

    const createEditor = () => {
      const doc = server.getDocument();
      const user = server.getUser();
      const documentType = getDocumentType(doc.fileType);

      server.setClient({
        buildVersion: window.DocsAPI!.DocEditor.version(),
      });
      editor = new window.DocsAPI!.DocEditor("placeholder", {
        document: {
          fileType: doc.fileType,
          key: doc.key,
          title: doc.title,
          url: doc.url,

          permissions: {
            edit: editing && doc.fileType !== "pdf",
            chat: false,
            rename: editing,
            protect: editing,
            review: false,
            print: false,
          },
        },
        documentType: documentType,
        editorConfig: {
          lang: lang,
          coEditing: {
            mode: "fast",
            change: false,
          },
          user: {
            ...user,
          },
          customization: {
            uiTheme: uiTheme,
            features: {
              spellcheck: {
                change: false,
              },
            },
            logo: {
              image: location.origin + "/logo-name_black.svg",
              imageDark: location.origin + "/logo-name_white.svg",
              url: location.origin,
            },
          },
        },
        events: {
          onAppReady: async (e: unknown) => {
            console.log("App ready", e, editor);
            onAppReady();
          },
          onDocumentReady: (e: unknown) => {
            console.log("Document ready", e);
            postBridgeMessage(BRIDGE_OPENED);
          },
          onDocumentStateChange: (e: { data: boolean; target: unknown }) => {
            console.log("Document state change", e);
            if (e.data) {
              isDirty.current = true;
              postBridgeMessage(BRIDGE_DIRTY, { dirty: true });
            }
          },
          onRequestOpen: (e: unknown) => {
            console.log("onRequestOpen", e);
          },
          onError: (e: unknown) => {
            console.log("Error", e);
          },
          onInfo: (e: unknown) => {
            console.log("Info", e);
          },
          onWarning: (e: unknown) => {
            console.log("onWarning", e);
          },
          onRequestSaveAs: (e: unknown) => {
            console.log("onRequestSaveAs", e);
          },
          onSaveDocument: (e: unknown) => {
            console.log("onSaveDocument", e);
            if (bridgeEnabled && isDirty.current) {
              requestExport();
            } else if (!bridgeEnabled) {
              isDirty.current = false;
            }
          },
          onDownloadAs: (e: unknown) => {
            console.log("onDownloadAs", e);
          },
          onSave: (e: unknown) => {
            console.log("onSave", e);
            if (bridgeEnabled && isDirty.current) {
              requestExport();
            } else if (!bridgeEnabled) {
              isDirty.current = false;
            }
          },
          writeFile: async (e: unknown) => {
            console.log("writeFile", e);
            if (bridgeEnabled && isDirty.current) {
              requestExport();
            } else if (!bridgeEnabled) {
              isDirty.current = false;
            }
          },
        },
        type: "desktop",
        width: "100%",
        height: "100%",
      });
      Object.assign(window, {
        editor,
      });
      return editor;
    };

    const loadEditor = () => {
      if (window.DocsAPI && window.DocsAPI.DocEditor) {
        createEditor();
        return;
      }
      let script = document.querySelector<HTMLScriptElement>(
        `script[src="${apiUrl}"]`
      );
      if (!script) {
        script = document.createElement("script");
        script.src = apiUrl;
        document.head.appendChild(script);
      }
      script.onload = () => {
        createEditor();
      };
      script.onerror = (e) => {
        console.error("Failed to load DocsAPI script", e);
      };
    };

    const handleBridgeMessage = async (event: MessageEvent) => {
      if (
        !bridgeEnabled ||
        event.source !== window.parent ||
        event.origin !== parentOrigin ||
        event.data?.requestId !== requestId
      ) {
        return;
      }

      if (event.data.type === BRIDGE_OPEN) {
        const { buffer, fileName, fileType, mimeType } = event.data;
        if (!(buffer instanceof ArrayBuffer) || typeof fileName !== "string") {
          postBridgeMessage(BRIDGE_ERROR, { message: "Invalid document data" });
          return;
        }
        try {
          await server.open(
            new File([buffer], fileName, {
              type:
                typeof mimeType === "string"
                  ? mimeType
                  : "application/octet-stream",
            }),
            {
              fileType: typeof fileType === "string" ? fileType : undefined,
              fileName,
            }
          );
          loadEditor();
        } catch (error) {
          console.error("Failed to open embedded document", error);
          postBridgeMessage(BRIDGE_ERROR, {
            message: "Failed to open the document",
          });
        }
        return;
      }

      if (event.data.type === BRIDGE_SAVE) requestExport();
    };

    window.addEventListener("message", handleBridgeMessage);

    const init = async () => {
      if (bridgeEnabled) {
        postBridgeMessage(BRIDGE_READY);
        return;
      }
      if (newDoc) {
        server.openNew(newDoc);
      }
      if (fileUrl && !fileId) {
        const { loader, tryDirect } = createExtensionLoader({
          onWaiting: () => setShowInstallHint(true),
          onReady: () => setShowInstallHint(false),
        });
        tryDirectRef.current = tryDirect;
        server.openUrl(fileUrl, {
          fileType: searchParams.get("fileType") || "",
          fileName: searchParams.get("fileName") || "",
          loader,
        });
      }
      loadEditor();
    };

    init();

    return () => {
      window.removeEventListener("message", handleBridgeMessage);
      server.setDownloadHandler(null);
      MockSocket.off("connect", server.handleConnect);
      MockSocket.off("disconnect", server.handleDisconnect);
      editor?.destroyEditor?.();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hasHydrated]);

  return (
    <>
      <InstallExtensionDialog
        open={showInstallHint}
        onClose={() => setShowInstallHint(false)}
        onTryDirect={tryDirectRef.current || undefined}
      />
      <div>
        <div className="w-screen h-screen">
          <div id="placeholder">
            <iframe
              className="w-0 h-0 hidden"
              src={APP_ROOT + PRELOAD_HTML}
            ></iframe>
          </div>
        </div>
      </div>
    </>
  );
}
