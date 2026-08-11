"use client";

import { useLayoutEffect, useRef, useEffect } from "react";
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
import { DocEditor, DocumentType } from "@/utils/editor/types";

const BRIDGE_READY = "xinghuo-office-ready";
const BRIDGE_OPEN = "xinghuo-office-open";
const BRIDGE_OPENED = "xinghuo-office-opened";
const BRIDGE_DIRTY = "xinghuo-office-dirty";
const BRIDGE_SAVE = "xinghuo-office-save";
const BRIDGE_SAVING = "xinghuo-office-saving";
const BRIDGE_SAVED = "xinghuo-office-saved";
const BRIDGE_ERROR = "xinghuo-office-error";
const BRIDGE_CANCEL_SAVE = "xinghuo-office-cancel-save";

const MIME_TYPES: Record<string, string> = {
  doc: "application/msword",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xls: "application/vnd.ms-excel",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ppt: "application/vnd.ms-powerpoint",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  pdf: "application/pdf",
};

type SpreadsheetEditorWindow = Window & {
  Asc?: {
    editor?: {
      asc_closeCellEditor?: (cancel?: boolean) => boolean;
    };
  };
};

const getEditorFrame = () =>
  document.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');

export default function Page() {
  const server = useAppStore((state) => state.server);
  const language = useResolvedLanguage();
  const theme = useAppStore((state) => state.theme);
  const hasHydrated = useHasHydrated();
  const isDirty = useRef(false);

  useEffect(() => {
    const updateViewportHeight = () => {
      const height = Math.round(window.visualViewport?.height ?? window.innerHeight);
      document.documentElement.style.setProperty("--office-viewport-height", `${height}px`);
    };

    updateViewportHeight();
    window.addEventListener("resize", updateViewportHeight, { passive: true });
    window.addEventListener("orientationchange", updateViewportHeight, {
      passive: true,
    });
    window.visualViewport?.addEventListener("resize", updateViewportHeight, {
      passive: true,
    });
    window.visualViewport?.addEventListener("scroll", updateViewportHeight, {
      passive: true,
    });

    return () => {
      window.removeEventListener("resize", updateViewportHeight);
      window.removeEventListener("orientationchange", updateViewportHeight);
      window.visualViewport?.removeEventListener("resize", updateViewportHeight);
      window.visualViewport?.removeEventListener("scroll", updateViewportHeight);
      document.documentElement.style.removeProperty("--office-viewport-height");
    };
  }, []);

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

    const newDoc = searchParams.get("new");
    const paramEditing = searchParams.get("editing");
    const paramLang = searchParams.get("lang");
    const paramTheme = searchParams.get("theme");
    const mobileMode =
      searchParams.get("mobile") === "1" ||
      window.matchMedia(
        "(max-width: 750px), (max-height: 520px) and (pointer: coarse)"
      ).matches;
    const compactToolbar =
      searchParams.get("compactToolbar") === "1" || mobileMode;
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
    let activeSaveId: string | undefined;
    let userInteracted = false;

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

    const commitPendingSpreadsheetEdit = () => {
      if (
        getDocumentType(server.getDocument().fileType) !== DocumentType.Cell
      ) {
        return;
      }

      const iframe = getEditorFrame();
      const spreadsheetApi = (iframe?.contentWindow as SpreadsheetEditorWindow)
        ?.Asc?.editor;
      if (typeof spreadsheetApi?.asc_closeCellEditor !== "function") {
        throw new Error("Spreadsheet editor API is not ready");
      }

      // OnlyOffice keeps the value being typed in a temporary cell editor
      // until the edit is explicitly closed. Exporting before this call would
      // save the previous cell value when Ctrl/Cmd+S is pressed mid-entry.
      if (spreadsheetApi.asc_closeCellEditor() === false) {
        throw new Error("The active spreadsheet cell could not be committed");
      }
    };

    const releaseEditorFocus = () => {
      try {
        editor?.blurFocus?.({});
      } catch (error) {
        console.warn("Failed to release the public Office focus", error);
      }

      const iframe = getEditorFrame();
      const activeElement = iframe?.contentDocument?.activeElement;
      if (activeElement instanceof HTMLElement) activeElement.blur();
      iframe?.blur();
    };

    const requestExport = (saveId?: string) => {
      if (!bridgeEnabled || !editor || saveInProgress) return;
      try {
        commitPendingSpreadsheetEdit();
        releaseEditorFocus();
        saveInProgress = true;
        activeSaveId = saveId;
        postBridgeMessage(BRIDGE_SAVING, { saveId });
        editor.downloadAs(server.getDocument().fileType);
      } catch (error) {
        saveInProgress = false;
        const failedSaveId = activeSaveId ?? saveId;
        activeSaveId = undefined;
        console.error("Failed to save embedded document", error);
        postBridgeMessage(BRIDGE_ERROR, {
          saveId: failedSaveId,
          message: "Failed to generate the updated document",
        });
      }
    };

    const resetEditorModifiedState = () => {
      try {
        // Public Docs API: clear the saved snapshot so the next edit emits a
        // fresh onDocumentStateChange(true) and Save becomes available again.
        editor?.setDocumentModified?.(false);
      } catch (error) {
        console.warn("Failed to reset the public Office modified state", error);
      }
    };

    server.setDownloadHandler(({ data, fileName, fileType }) => {
      if (!bridgeEnabled) return false;
      // A cancelled or superseded bridge export must still be consumed here;
      // otherwise EditorServer falls back to a browser download.
      if (!saveInProgress) return true;
      const completedSaveId = activeSaveId;
      saveInProgress = false;
      activeSaveId = undefined;
      isDirty.current = false;
      // downloadAs exports a snapshot but does not always reset OnlyOffice's
      // internal modified flag. Reset it so the next edit produces a fresh
      // onDocumentStateChange(true) event and can be saved again.
      resetEditorModifiedState();
      postBridgeMessage(
        BRIDGE_SAVED,
        {
          saveId: completedSaveId,
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
      const iframe = getEditorFrame();
      const win = iframe?.contentWindow as typeof window;
      const iframeDoc = iframe?.contentDocument;
      if (!iframeDoc || !win) {
        throw new Error("Iframe not loaded");
      }

      const markUserInteraction = () => {
        userInteracted = true;
      };
      iframeDoc.addEventListener("pointerdown", markUserInteraction, {
        capture: true,
        passive: true,
      });
      iframeDoc.addEventListener("keydown", markUserInteraction, true);
      iframeDoc.addEventListener("beforeinput", markUserInteraction, true);

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
            compactToolbar,
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
              if (bridgeEnabled && !userInteracted) return;
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
        type: mobileMode ? "mobile" : "desktop",
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

      if (event.data.type === BRIDGE_SAVE) {
        requestExport(
          typeof event.data.saveId === "string" ? event.data.saveId : undefined
        );
        return;
      }

      if (event.data.type === BRIDGE_CANCEL_SAVE) {
        const saveId =
          typeof event.data.saveId === "string" ? event.data.saveId : undefined;
        if (!saveId || !activeSaveId || saveId === activeSaveId) {
          saveInProgress = false;
          activeSaveId = undefined;
        }
      }
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
      <div>
        <div
          className="w-screen overflow-hidden"
          style={{ height: "var(--office-viewport-height, 100dvh)" }}
        >
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
