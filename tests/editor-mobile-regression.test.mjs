import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const readSource = (relativePath) =>
  readFile(new URL(`../${relativePath}`, import.meta.url), "utf8");

test("mobile previews stay responsive while Community Edition editing uses the desktop engine", async () => {
  const editorSource = await readSource("app/editor/page.tsx");
  const layoutSource = await readSource("app/layout.tsx");
  const serverSource = await readSource("utils/editor/server.ts");
  const converterSource = await readSource("utils/editor/x2t.ts");

  assert.match(
    editorSource,
    /const requestedMobileMode = searchParams\.get\("mobile"\)/
  );
  assert.match(editorSource, /requestedMobileMode === "1"/);
  assert.match(
    editorSource,
    /type: mobileMode && !editing \? "mobile" : "desktop"/
  );
  assert.match(
    editorSource,
    /mode: editing && doc\.fileType !== "pdf" \? "edit" : "view"/
  );
  assert.match(editorSource, /forceView: !editing/);
  assert.match(
    editorSource,
    /Community Edition intentionally disables editing/
  );
  assert.match(editorSource, /BRIDGE_SOURCE_RECEIVED/);
  assert.match(editorSource, /postBridgeMessage\(BRIDGE_SOURCE_RECEIVED\)/);
  assert.match(editorSource, /BRIDGE_SOURCE_BEGIN/);
  assert.match(editorSource, /BRIDGE_SOURCE_CHUNK_RECEIVED/);
  assert.match(editorSource, /decodeBase64Chunk/);
  assert.match(editorSource, /supportsChunkedSource: true/);
  assert.match(editorSource, /protocolVersion: BRIDGE_PROTOCOL_VERSION/);
  assert.match(editorSource, /BRIDGE_DIAGNOSTIC/);
  assert.match(editorSource, /postDiagnostic\(\s*"office_window_error"/);
  assert.match(editorSource, /postDiagnostic\("office_unhandled_rejection"/);
  assert.match(editorSource, /postDiagnostic\("document_conversion_start"/);
  assert.match(editorSource, /postDiagnostic\("document_conversion_complete"/);
  assert.match(editorSource, /postDiagnostic\("onlyoffice_error"/);
  assert.match(editorSource, /postDiagnostic\("onlyoffice_document_ready"/);
  assert.match(editorSource, /summarizeDiagnosticValue/);
  assert.match(editorSource, /buffer: new Uint8Array\(byteLength\)/);
  assert.match(editorSource, /receivedChunks: new Uint8Array\(chunkCount\)/);
  assert.match(editorSource, /completedSource\.buffer\.buffer as ArrayBuffer/);
  assert.doesNotMatch(editorSource, /new Uint8Array\(completedBuffer\)\.set/);
  assert.match(editorSource, /server\.openBuffer\(buffer/);
  assert.match(editorSource, /transferInput: true/);
  assert.match(editorSource, /waitForLoad: true/);
  assert.match(serverSource, /async openBuffer\(/);
  assert.match(serverSource, /this\.loadDocument\(buffer, this\.fileType, transferInput\)/);
  assert.match(converterSource, /transferInput \? data : data\.slice\(0\)/);
  assert.match(editorSource, /attributes: true/);
  assert.match(editorSource, /renderedPreviewErrorObserver\.observe\(document\.body/);
  assert.match(editorSource, /scheduleRenderedPreviewErrorDismissal/);
  assert.match(editorSource, /Chunked document transfer is incomplete/);
  assert.match(editorSource, /dismissRenderedMobileSpreadsheetError/);
  assert.match(editorSource, /embeddedDocumentReady/);
  assert.match(editorSource, /activeDocumentType === DocumentType\.Cell/);
  assert.match(editorSource, /installReadOnlySpreadsheetTapGuard/);
  assert.match(editorSource, /ws-canvas-overlay/);
  assert.match(editorSource, /event\.stopImmediatePropagation\(\)/);
  assert.match(editorSource, /non-fatal mobile spreadsheet preview error/);
  assert.match(editorSource, /Failed to initialize the document editor/);
  assert.match(editorSource, /Failed to load the document editor assets/);
  assert.match(editorSource, /compactToolbar,/);
  assert.match(editorSource, /window\.visualViewport\?\.height/);
  assert.match(editorSource, /--office-viewport-height/);
  assert.match(layoutSource, /interactiveWidget: "resizes-content"/);
});

test("saving commits spreadsheet input and releases mobile keyboard focus", async () => {
  const source = await readSource("app/editor/page.tsx");
  const requestExportStart = source.indexOf("const requestExport =");
  const requestExportEnd = source.indexOf(
    "const resetEditorModifiedState",
    requestExportStart
  );
  const requestExportSource = source.slice(
    requestExportStart,
    requestExportEnd
  );

  assert.match(source, /asc_closeCellEditor/);
  assert.match(source, /editor\?\.blurFocus\?\.\(\{\}\)/);
  assert.match(source, /activeElement instanceof HTMLElement/);
  assert.match(requestExportSource, /commitPendingSpreadsheetEdit\(\)/);
  assert.match(requestExportSource, /releaseEditorFocus\(\)/);
  assert.ok(
    requestExportSource.indexOf("commitPendingSpreadsheetEdit()") <
      requestExportSource.indexOf("editor.downloadAs")
  );
});

test("the mobile fixes publish under a new immutable Office image tag", async () => {
  const workflowSource = await readSource(
    ".github/workflows/build-office-image.yml"
  );

  assert.match(workflowSource, /type=raw,value=9\.4\.0\.1-7/);
  assert.doesNotMatch(workflowSource, /type=raw,value=9\.4\.0\.1-[123456]/);
  assert.match(workflowSource, /load: true/);
  assert.match(workflowSource, /push: false/);
  assert.match(workflowSource, /cache-from: type=gha,scope=xinghuo-office-9\.4/);
  assert.match(workflowSource, /cache-to: type=gha,mode=min/);
  assert.match(workflowSource, /provenance: false/);
  assert.match(workflowSource, /sbom: false/);
  assert.doesNotMatch(workflowSource, /docker pull "\$IMAGE_REF"/);
  assert.match(workflowSource, /docker push "\$IMAGE_REF"/);
});
