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
const BRIDGE_SOURCE_BEGIN = "xinghuo-office-source-begin";
const BRIDGE_SOURCE_CHUNK = "xinghuo-office-source-chunk";
const BRIDGE_SOURCE_CHUNK_RECEIVED = "xinghuo-office-source-chunk-received";
const BRIDGE_SOURCE_END = "xinghuo-office-source-end";
const BRIDGE_SOURCE_RECEIVED = "xinghuo-office-source-received";
const BRIDGE_OPENED = "xinghuo-office-opened";
const BRIDGE_DIRTY = "xinghuo-office-dirty";
const BRIDGE_SAVE = "xinghuo-office-save";
const BRIDGE_SAVING = "xinghuo-office-saving";
const BRIDGE_SAVED = "xinghuo-office-saved";
const BRIDGE_ERROR = "xinghuo-office-error";
const BRIDGE_DIAGNOSTIC = "xinghuo-office-diagnostic";
const BRIDGE_CANCEL_SAVE = "xinghuo-office-cancel-save";
const BRIDGE_PROTOCOL_VERSION = 2;
const MAX_CHUNKED_SOURCE_BYTES = 256 * 1024 * 1024;
const MAX_CHUNK_COUNT = 4096;

const summarizeDiagnosticValue = (value: unknown): string => {
  let text: string;
  if (value instanceof Error) {
    text = `${value.name}: ${value.message}${value.stack ? `\n${value.stack}` : ""}`;
  } else if (typeof value === "string") {
    text = value;
  } else {
    try {
      const candidate =
        typeof value === "object" && value !== null && "data" in value
          ? { data: (value as { data?: unknown }).data }
          : value;
      text = JSON.stringify(candidate);
    } catch {
      text = String(value);
    }
  }
  return text
    .replace(/(access[_-]?token|authorization|cookie|password|secret|key)=([^&\s]+)/gi, "$1=[redacted]")
    .replace(/Bearer\s+[^\s]+/gi, "Bearer [redacted]")
    .slice(0, 1200);
};

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

type ChunkedSource = {
  fileName: string;
  fileType?: string;
  mimeType: string;
  byteLength: number;
  chunkCount: number;
  chunkSize?: number;
  buffer: Uint8Array;
  receivedChunks: Uint8Array;
  receivedBytes: number;
};

const decodeBase64Chunk = (encoded: string): Uint8Array => {
  const binary = window.atob(encoded);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
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
      const height = Math.round(
        window.visualViewport?.height ?? window.innerHeight
      );
      document.documentElement.style.setProperty(
        "--office-viewport-height",
        `${height}px`
      );
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
      window.visualViewport?.removeEventListener(
        "resize",
        updateViewportHeight
      );
      window.visualViewport?.removeEventListener(
        "scroll",
        updateViewportHeight
      );
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
    const requestedMobileMode = searchParams.get("mobile");
    const mobileMode =
      requestedMobileMode === null
        ? window.matchMedia(
            "(max-width: 750px), (max-height: 520px) and (pointer: coarse)"
          ).matches
        : requestedMobileMode === "1";
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
    let chunkedSource: ChunkedSource | null = null;
    let embeddedDocumentReady = false;
    let activeDocumentType: DocumentType | null = null;
    let renderedPreviewErrorObserver: MutationObserver | null = null;
    let removeReadOnlySpreadsheetGuards: (() => void) | null = null;
    const renderedPreviewErrorTimers = new Set<number>();

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

    const postDiagnostic = (
      stage: string,
      level: "info" | "warning" | "error" = "info",
      value?: unknown
    ) => {
      postBridgeMessage(BRIDGE_DIAGNOSTIC, {
        stage,
        level,
        ...(value === undefined
          ? {}
          : { message: summarizeDiagnosticValue(value) }),
      });
    };

    const handleDiagnosticWindowError = (event: ErrorEvent) => {
      postDiagnostic(
        "office_window_error",
        "error",
        event.error || `${event.message} (${event.lineno}:${event.colno})`
      );
    };
    const handleDiagnosticUnhandledRejection = (event: PromiseRejectionEvent) => {
      postDiagnostic("office_unhandled_rejection", "error", event.reason);
    };
    window.addEventListener("error", handleDiagnosticWindowError);
    window.addEventListener(
      "unhandledrejection",
      handleDiagnosticUnhandledRejection
    );
    postDiagnostic("office_bridge_initialized", "info", {
      online: navigator.onLine,
      viewport: `${window.innerWidth}x${window.innerHeight}`,
      mobileMode,
      editing,
    });

    const dismissRenderedMobileSpreadsheetError = () => {
      if (
        !bridgeEnabled ||
        editing ||
        !mobileMode ||
        !embeddedDocumentReady
      ) {
        return;
      }

      const iframeDoc = getEditorFrame()?.contentDocument;
      const candidateDocuments = [document, iframeDoc].filter(
        (candidate): candidate is Document => Boolean(candidate)
      );

      const processingErrorPattern =
        /在处理此文档时发生了错误|error occurred while (?:processing|working with) (?:this|the) document/i;
      const communityNoticePattern =
        /使用免费的社区版本|要访问移动web编辑器，需要商业许可证|using the free community version|mobile web editors?,? a commercial license is required/i;
      const isProcessingError = activeDocumentType === DocumentType.Cell;
      const dialogSelector =
        '[role="dialog"], [aria-modal="true"], .asc-window, .asc-popup, .dialog, .modal, .popup, [class*="dialog"], [class*="modal"], [class*="popup"], [class*="window"]';
      const dialogCandidates = candidateDocuments.flatMap((candidateDocument) =>
        Array.from(candidateDocument.querySelectorAll<HTMLElement>(dialogSelector))
      );
      const findMatchingDialog = (pattern: RegExp) =>
        dialogCandidates.find((dialog) => pattern.test(dialog.textContent || ""));
      const processingError = isProcessingError
        ? findMatchingDialog(processingErrorPattern)
        : undefined;
      const communityNotice = findMatchingDialog(communityNoticePattern);
      const targetDialog = processingError || communityNotice;
      if (!targetDialog) return;

      const actionPattern = /^(?:确定|好|关闭|OK|Close)$/i;
      const actions = Array.from(
        targetDialog.querySelectorAll<HTMLElement>(
          'button, [role="button"], .asc-button, .button, [class*="button"], input[type="button"], input[type="submit"]'
        )
      );
      let confirmAction = actions.find((action) =>
        actionPattern.test((action.textContent || (action as HTMLInputElement).value || "").trim())
      );
      if (!confirmAction) {
        const allActions = candidateDocuments.flatMap((candidateDocument) =>
          Array.from(
            candidateDocument.querySelectorAll<HTMLElement>(
              'button, [role="button"], .asc-button, .button, [class*="button"], input[type="button"], input[type="submit"]'
            )
          )
        );
        confirmAction = allActions.find((action) => {
          if (
            !actionPattern.test(
              (action.textContent || (action as HTMLInputElement).value || "").trim()
            )
          ) {
            return false;
          }
          let parent: HTMLElement | null = action.parentElement;
          for (let depth = 0; parent && depth < 8; depth += 1) {
            if (processingErrorPattern.test(parent.textContent || "") || communityNoticePattern.test(parent.textContent || "")) {
              return true;
            }
            parent = parent.parentElement;
          }
          return false;
        });
      }
      if (!confirmAction) return;

      console.warn(
        processingError
          ? "Dismissed a non-fatal mobile spreadsheet preview error after the document rendered"
          : "Dismissed the Community mobile preview notice in read-only mode"
      );
      confirmAction.click();
    };

    const scheduleRenderedPreviewErrorDismissal = () => {
      [0, 100, 250, 500, 1000, 2000, 4000, 8000, 12000, 20000, 30000].forEach((delay) => {
        const timer = window.setTimeout(() => {
          renderedPreviewErrorTimers.delete(timer);
          dismissRenderedMobileSpreadsheetError();
        }, delay);
        renderedPreviewErrorTimers.add(timer);
      });
    };

    const installReadOnlySpreadsheetTapGuard = (iframe: HTMLIFrameElement) => {
      if (
        !bridgeEnabled ||
        editing ||
        !mobileMode ||
        activeDocumentType !== DocumentType.Cell
      ) {
        return;
      }

      const frameWindow = iframe.contentWindow;
      if (!frameWindow) return;

      let pointerStart: { x: number; y: number; id: number } | null = null;
      let touchStart: { x: number; y: number } | null = null;
      const tapThreshold = 8;
      const isSpreadsheetCanvas = (target: EventTarget | null) => {
        const element = target as Element | null;
        if (!element || typeof element.closest !== "function") return false;
        return Boolean(
          element.closest(
            'canvas, #ws-canvas-outer, #ws-canvas, #ws-canvas-overlay, [id*="ws-canvas"], .ws-canvas-outer'
          )
        );
      };
      const isTap = (
        start: { x: number; y: number } | null,
        x: number,
        y: number
      ) =>
        Boolean(
          start &&
            Math.abs(x - start.x) <= tapThreshold &&
            Math.abs(y - start.y) <= tapThreshold
        );
      const blockCellTap = (event: Event) => {
        if (!isSpreadsheetCanvas(event.target)) return;
        event.preventDefault();
        event.stopImmediatePropagation();
      };
      const handlePointerDown = (event: PointerEvent) => {
        pointerStart = { x: event.clientX, y: event.clientY, id: event.pointerId };
      };
      const handlePointerUp = (event: PointerEvent) => {
        if (
          pointerStart?.id === event.pointerId &&
          isTap(pointerStart, event.clientX, event.clientY)
        ) {
          blockCellTap(event);
        }
        pointerStart = null;
      };
      const handleTouchStart = (event: TouchEvent) => {
        const touch = event.touches[0];
        touchStart = touch ? { x: touch.clientX, y: touch.clientY } : null;
      };
      const handleTouchEnd = (event: TouchEvent) => {
        const touch = event.changedTouches[0];
        if (touch && isTap(touchStart, touch.clientX, touch.clientY)) {
          blockCellTap(event);
        }
        touchStart = null;
      };

      frameWindow.addEventListener("pointerdown", handlePointerDown, true);
      frameWindow.addEventListener("pointerup", handlePointerUp, true);
      frameWindow.addEventListener("touchstart", handleTouchStart, true);
      frameWindow.addEventListener("touchend", handleTouchEnd, true);
      frameWindow.addEventListener("click", blockCellTap, true);
      frameWindow.addEventListener("dblclick", blockCellTap, true);
      removeReadOnlySpreadsheetGuards = () => {
        frameWindow.removeEventListener("pointerdown", handlePointerDown, true);
        frameWindow.removeEventListener("pointerup", handlePointerUp, true);
        frameWindow.removeEventListener("touchstart", handleTouchStart, true);
        frameWindow.removeEventListener("touchend", handleTouchEnd, true);
        frameWindow.removeEventListener("click", blockCellTap, true);
        frameWindow.removeEventListener("dblclick", blockCellTap, true);
      };
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

      renderedPreviewErrorObserver?.disconnect();
      renderedPreviewErrorObserver = new MutationObserver(() => {
        scheduleRenderedPreviewErrorDismissal();
      });
      renderedPreviewErrorObserver.observe(iframeDoc.body, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ["class", "style", "aria-hidden"],
      });
      renderedPreviewErrorObserver.observe(document.body, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ["class", "style", "aria-hidden"],
      });
      removeReadOnlySpreadsheetGuards?.();
      installReadOnlySpreadsheetTapGuard(iframe);

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
      activeDocumentType = documentType;
      postDiagnostic("onlyoffice_editor_create_start", "info", {
        documentType,
        fileType: doc.fileType,
        mobileMode,
        editing,
      });

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
          // A read-only mobile preview must be a real view session. Relying
          // only on permissions.edit still lets Community mobile bundles
          // initialize their edit/license path and show a warning dialog.
          mode: editing && doc.fileType !== "pdf" ? "edit" : "view",
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
            mobile: {
              forceView: !editing,
            },
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
            postDiagnostic("onlyoffice_app_ready");
            onAppReady();
          },
          onDocumentReady: (e: unknown) => {
            console.log("Document ready", e);
            postDiagnostic("onlyoffice_document_ready");
            embeddedDocumentReady = true;
            postBridgeMessage(BRIDGE_OPENED);
            scheduleRenderedPreviewErrorDismissal();
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
            postDiagnostic("onlyoffice_error", "error", e);
            scheduleRenderedPreviewErrorDismissal();
          },
          onInfo: (e: unknown) => {
            console.log("Info", e);
          },
          onWarning: (e: unknown) => {
            console.log("onWarning", e);
            postDiagnostic("onlyoffice_warning", "warning", e);
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
        // ONLYOFFICE Community Edition intentionally disables editing in its
        // mobile Web editor. Keep the responsive mobile reader for previews,
        // but use the fully editable desktop engine for mobile edit sessions.
        type: mobileMode && !editing ? "mobile" : "desktop",
        width: "100%",
        height: "100%",
      });
      Object.assign(window, {
        editor,
      });
      postDiagnostic("onlyoffice_editor_created");
      return editor;
    };

    const createEditorSafely = () => {
      try {
        createEditor();
      } catch (error) {
        console.error("Failed to create DocsAPI editor", error);
        postDiagnostic("onlyoffice_editor_create_error", "error", error);
        postBridgeMessage(BRIDGE_ERROR, {
          message: "Failed to initialize the document editor",
        });
      }
    };

    const loadEditor = () => {
      if (window.DocsAPI && window.DocsAPI.DocEditor) {
        postDiagnostic("onlyoffice_api_already_loaded");
        createEditorSafely();
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
        postDiagnostic("onlyoffice_api_loaded");
        createEditorSafely();
      };
      script.onerror = (e) => {
        console.error("Failed to load DocsAPI script", e);
        postDiagnostic("onlyoffice_api_load_error", "error", e);
        postBridgeMessage(BRIDGE_ERROR, {
          message: "Failed to load the document editor assets",
        });
      };
    };

    const openEmbeddedDocument = async (
      buffer: ArrayBuffer,
      fileName: string,
      fileType: string | undefined,
      mimeType: string
    ) => {
      postBridgeMessage(BRIDGE_SOURCE_RECEIVED);
      postDiagnostic("document_conversion_start", "info", {
        byteLength: buffer.byteLength,
        fileType: fileType || "unknown",
        mobileMode,
      });
      if (mobileMode) {
        await server.openBuffer(buffer, {
          fileType,
          fileName,
          transferInput: true,
          waitForLoad: true,
        });
      } else {
        await server.open(new File([buffer], fileName, { type: mimeType }), {
          fileType,
          fileName,
        });
      }
      postDiagnostic("document_conversion_complete");
      loadEditor();
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
          await openEmbeddedDocument(
            buffer,
            fileName,
            typeof fileType === "string" ? fileType : undefined,
            typeof mimeType === "string" ? mimeType : "application/octet-stream"
          );
        } catch (error) {
          console.error("Failed to open embedded document", error);
          postDiagnostic("document_open_error", "error", error);
          postBridgeMessage(BRIDGE_ERROR, {
            message: "Failed to open the document",
          });
        }
        return;
      }

      if (event.data.type === BRIDGE_SOURCE_BEGIN) {
        const { fileName, fileType, mimeType, byteLength, chunkCount } =
          event.data;
        const requestedChunkSize = event.data.chunkSize;
        if (
          typeof fileName !== "string" ||
          typeof byteLength !== "number" ||
          !Number.isSafeInteger(byteLength) ||
          byteLength <= 0 ||
          byteLength > MAX_CHUNKED_SOURCE_BYTES ||
          typeof chunkCount !== "number" ||
          !Number.isSafeInteger(chunkCount) ||
          chunkCount <= 0 ||
          chunkCount > MAX_CHUNK_COUNT ||
          (requestedChunkSize !== undefined &&
            (!Number.isSafeInteger(requestedChunkSize) ||
              requestedChunkSize <= 0 ||
              requestedChunkSize > 256 * 1024))
        ) {
          chunkedSource = null;
          postBridgeMessage(BRIDGE_ERROR, {
            message: "Invalid chunked document metadata",
          });
          return;
        }
        chunkedSource = {
          fileName,
          fileType: typeof fileType === "string" ? fileType : undefined,
          mimeType:
            typeof mimeType === "string"
              ? mimeType
              : "application/octet-stream",
          byteLength,
          chunkCount,
          chunkSize: requestedChunkSize,
          buffer: new Uint8Array(byteLength),
          receivedChunks: new Uint8Array(chunkCount),
          receivedBytes: 0,
        };
        postDiagnostic("chunked_source_accepted", "info", {
          byteLength,
          chunkCount,
          chunkSize: requestedChunkSize || null,
        });
        return;
      }

      if (event.data.type === BRIDGE_SOURCE_CHUNK) {
        const { chunkIndex, chunkData } = event.data;
        if (
          !chunkedSource ||
          typeof chunkIndex !== "number" ||
          !Number.isSafeInteger(chunkIndex) ||
          chunkIndex < 0 ||
          chunkIndex >= chunkedSource.chunkCount ||
          typeof chunkData !== "string"
        ) {
          postBridgeMessage(BRIDGE_ERROR, {
            message: "Invalid chunked document data",
          });
          return;
        }
        try {
          if (!chunkedSource.receivedChunks[chunkIndex]) {
            const bytes = decodeBase64Chunk(chunkData);
            if (!chunkedSource.chunkSize) {
              if (chunkIndex !== 0) {
                throw new Error("The first document chunk was not received");
              }
              chunkedSource.chunkSize = bytes.byteLength;
            }
            const offset = chunkIndex * chunkedSource.chunkSize;
            const expectedLength =
              chunkIndex === chunkedSource.chunkCount - 1
                ? chunkedSource.byteLength - offset
                : chunkedSource.chunkSize;
            if (
              expectedLength <= 0 ||
              bytes.byteLength !== expectedLength ||
              offset + bytes.byteLength > chunkedSource.byteLength
            ) {
              throw new Error("Invalid document chunk size");
            }
            chunkedSource.buffer.set(bytes, offset);
            chunkedSource.receivedChunks[chunkIndex] = 1;
            chunkedSource.receivedBytes += bytes.byteLength;
          }
          postBridgeMessage(BRIDGE_SOURCE_CHUNK_RECEIVED, { chunkIndex });
        } catch (error) {
          console.error("Failed to decode document chunk", error);
          chunkedSource = null;
          postBridgeMessage(BRIDGE_ERROR, {
            message: "Failed to decode chunked document data",
          });
        }
        return;
      }

      if (event.data.type === BRIDGE_SOURCE_END) {
        postDiagnostic("chunked_source_end_received");
        const completedSource = chunkedSource;
        chunkedSource = null;
        if (
          !completedSource ||
          completedSource.receivedBytes !== completedSource.byteLength ||
          completedSource.receivedChunks.some((received) => !received)
        ) {
          postBridgeMessage(BRIDGE_ERROR, {
            message: "Chunked document transfer is incomplete",
          });
          return;
        }
        try {
          // This view was allocated with `new Uint8Array(byteLength)`, so its
          // backing store is a real ArrayBuffer. Pass it through directly;
          // copying it here can exhaust an Android WebView before conversion.
          const completedBuffer = completedSource.buffer.buffer as ArrayBuffer;
          await openEmbeddedDocument(
            completedBuffer,
            completedSource.fileName,
            completedSource.fileType,
            completedSource.mimeType
          );
        } catch (error) {
          console.error("Failed to open chunked embedded document", error);
          postDiagnostic("chunked_document_open_error", "error", error);
          postBridgeMessage(BRIDGE_ERROR, {
            message: "Failed to open the chunked document",
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
        postDiagnostic("bridge_ready_sent");
        postBridgeMessage(BRIDGE_READY, {
          protocolVersion: BRIDGE_PROTOCOL_VERSION,
          supportsChunkedSource: true,
        });
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
      window.removeEventListener("error", handleDiagnosticWindowError);
      window.removeEventListener(
        "unhandledrejection",
        handleDiagnosticUnhandledRejection
      );
      renderedPreviewErrorObserver?.disconnect();
      removeReadOnlySpreadsheetGuards?.();
      renderedPreviewErrorTimers.forEach((timer) => window.clearTimeout(timer));
      renderedPreviewErrorTimers.clear();
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
