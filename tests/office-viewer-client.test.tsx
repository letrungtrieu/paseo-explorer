import React from "react";
import { act, create, type ReactTestRenderer } from "react-test-renderer";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { PluginFile } from "@getpaseo/plugin/client";
import { Pressable, ScrollView, Text, View } from "react-native";
import { DocumentFrame } from "@getpaseo/plugin/client/ui";
import { OfficeFilePanel, OfficeViewerPanel } from "../client/office-viewer-panel";
import { officeStore } from "../shared/office-store";
import type { OfficeFileData } from "../shared/viewer";
import { createOfficeDocumentUrl, readSpreadsheetViewport } from "../client/web";
import { WORKSHEET_WINDOW_ROWS, WORKSHEET_WINDOW_COLUMNS } from "../shared/spreadsheet-window";

const hostMocks = vi.hoisted(() => ({ readOfficeFile: vi.fn(), scrollTo: vi.fn(), fileUrl: vi.fn(),
  platform: "web", dataPosition: { x: 0, y: 0 },
  scrollEvents: [] as { pane: unknown; x?: number; y?: number; animated?: boolean }[],
}));
vi.mock("react-native", () => ({
  // Render actual component trees without loading React Native's platform-specific Flow runtime.
  Platform: { get OS() { return hostMocks.platform; } }, View: "View", Text: "Text", Pressable: "Pressable",
  ScrollView: React.forwardRef((props: Record<string, unknown>, ref) => {
    React.useImperativeHandle(ref, () => ({ scrollTo: (offsets: { x?: number; y?: number; animated?: boolean }) => {
      // Identify the data owner separately from header mirrors when checking scroll synchronization.
      hostMocks.scrollEvents.push({ pane: props.testID, ...offsets });
      hostMocks.scrollTo(offsets);
    } }));
    return React.createElement("ScrollView", props);
  }),
}));
vi.mock("@getpaseo/plugin/client", () => ({
  useRpc: () => hostMocks.readOfficeFile,
  useWorkspace: () => "/workspace",
  useFileUrl: (...args: unknown[]) => hostMocks.fileUrl(...args),
}));
vi.mock("@getpaseo/plugin/client/ui", () => ({
  DocumentFrame: "DocumentFrame", Loading: "Loading",
}));

let renderedViewer: ReactTestRenderer | undefined;
let queryClient: QueryClient;
const palette = { colors: { foreground: "#eee", foregroundMuted: "#aaa", surface0: "#111", accent: "#38f" } };

async function flushQueries() {
  // React Query batches notifications onto a timer; await that boundary before reading the mounted tree.
  await act(async () => { await new Promise((resolve) => setTimeout(resolve, 15)); });
}
async function renderViewer(viewer: React.ReactElement = <OfficeViewerPanel theme={palette} workspaceId="workspace-1" />) {
  await act(async () => {
    renderedViewer = create(<QueryClientProvider client={queryClient}>{viewer}</QueryClientProvider>);
  });
  await flushQueries();
  return renderedViewer!;
}

beforeEach(() => {
  (globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
  queryClient = new QueryClient({ defaultOptions: { queries: { retry: false, gcTime: 0 } } });
  officeStore.reset();
  hostMocks.readOfficeFile.mockReset();
  hostMocks.scrollTo.mockReset();
  hostMocks.platform = "web";
  hostMocks.dataPosition = { x: 0, y: 0 };
  hostMocks.scrollEvents = [];
  hostMocks.scrollTo.mockImplementation((offsets: { x?: number; y?: number }) => {
    // Browser scroll events report both axes, including the offset retained from the preceding operation.
    if (offsets.x !== undefined) hostMocks.dataPosition.x = offsets.x;
    if (offsets.y !== undefined) hostMocks.dataPosition.y = offsets.y;
  });
  hostMocks.fileUrl.mockReturnValue({ status: "ready", url: "blob:host-pdf" });
});
afterEach(async () => {
  // Unmount first so store broadcasts cannot notify components after their query provider is gone.
  await act(async () => renderedViewer?.unmount());
  renderedViewer = undefined;
  queryClient.clear();
  officeStore.reset();
  vi.restoreAllMocks();
});

function spreadsheetPreview(input: { filePath: string; sheet?: string; rowStart?: number; columnStart?: number; rowLimit?: number; columnLimit?: number }): OfficeFileData {
  const rowStart = input.rowStart ?? 0;
  const columnStart = input.columnStart ?? 0;
  // Simulate a huge sheet while returning the same bounded window contract as the server.
  const sheetNames = ["Giao dịch", "Tóm tắt"];
  return { filePath: input.filePath, fileType: "xlsx", sheets: sheetNames.map((name) => ({
    name, rowStart, columnStart, rowCount: 100_000, columnCount: 1_000,
    rows: name === (input.sheet ?? sheetNames[0])
      ? Array.from({ length: input.rowLimit ?? WORKSHEET_WINDOW_ROWS }, (_, visibleRow) => Array.from({ length: input.columnLimit ?? WORKSHEET_WINDOW_COLUMNS }, (_, visibleColumn) =>
        `${name}:${rowStart + visibleRow}:${columnStart + visibleColumn}`)) : [],
  })) };
}

async function scrollViewer(viewer: ReactTestRenderer, axis: "vertical" | "horizontal", offset: number) {
  const scrollPane = viewer.root.findAll((node) => node.type === ScrollView &&
    node.props.testID === "office-sheet-data-scroll")[0] ?? viewer.root.findAll((node) => node.type === ScrollView &&
    node.props.testID === `office-sheet-${axis}-scroll`)[0];
  // Send the same viewport measurements that native/browser scroll events deliver to the production hook.
  await act(async () => scrollPane.props.onScroll({ nativeEvent: {
    contentOffset: { ...hostMocks.dataPosition, ...(axis === "vertical" ? { y: offset } : { x: offset }) },
    layoutMeasurement: { width: 600, height: 600 },
  } }));
  await flushQueries();
}

describe("spreadsheet viewport ownership", () => {
  it("measures the usable data rectangle without treating native scrollbar space as cell pixels", () => {
    // RN Web's event measurements include scrollbars; fixed gutters need the smaller client rectangle.
    expect(readSpreadsheetViewport({ getScrollableNode: () => ({
      clientWidth: 480, clientHeight: 360, offsetWidth: 495, offsetHeight: 375,
    }) })).toEqual({ width: 480, height: 360, right: 15, bottom: 15 });
    hostMocks.platform = "android";
    expect(readSpreadsheetViewport({ getScrollableNode: () => { throw new Error("DOM access on native"); } })).toBeNull();
  });

  it("bounds the panel and synchronizes fixed gutters with a single data scroller", async () => {
    hostMocks.readOfficeFile.mockImplementation(async (input) => spreadsheetPreview(input));
    officeStore.openAndJump({ filePath: "wide.xlsx", ext: ".xlsx", cell: "A1" });
    const viewer = await renderViewer();
    expect(viewer.root.findByProps({ testID: "office-viewer-panel" }).props.style).toMatchObject({
      minWidth: 0, minHeight: 0, overflow: "hidden",
    });
    const scrollPanes = viewer.root.findAllByType(ScrollView);
    const dataPane = scrollPanes.find((pane) => pane.props.testID === "office-sheet-data-scroll")!;
    const columnGutter = scrollPanes.find((pane) => pane.props.testID === "office-sheet-column-headings")!;
    const rowGutter = scrollPanes.find((pane) => pane.props.testID === "office-sheet-row-headings")!;
    // Only the data owner scrolls interactively; fixed gutters remain separate siblings with hidden scrollbars.
    expect(dataPane.props.style).toMatchObject({ flex: 1, minWidth: 0, minHeight: 0, overflow: "scroll" });
    expect(columnGutter.props.scrollEnabled).toBe(false);
    expect(rowGutter.props.scrollEnabled).toBe(false);
    hostMocks.scrollEvents = [];
    await scrollViewer(viewer, "vertical", 400);
    await scrollViewer(viewer, "horizontal", 200);
    expect(hostMocks.scrollEvents).toContainEqual({ pane: "office-sheet-row-headings", y: 400, animated: false });
    expect(hostMocks.scrollEvents).toContainEqual({ pane: "office-sheet-column-headings", x: 200, animated: false });
  });

  it("constrains the native nested data scrollers to the measured viewport height", async () => {
    hostMocks.platform = "android";
    hostMocks.readOfficeFile.mockImplementation(async (input) => spreadsheetPreview(input));
    officeStore.openAndJump({ filePath: "native.xlsx", ext: ".xlsx", cell: "A1" });
    const viewer = await renderViewer();
    // A native horizontal container must not inherit the height of the entire buffered worksheet.
    const nativeFrame = viewer.root.findAll((node) => typeof node.props.onLayout === "function" && node.type === View)[0];
    await act(async () => nativeFrame.props.onLayout({ nativeEvent: { layout: { width: 320, height: 480 } } }));
    const verticalData = viewer.root.findAllByType(ScrollView).find((pane) => pane.props.testID === "office-sheet-vertical-scroll")!;
    const horizontalData = viewer.root.findAllByType(ScrollView).find((pane) => pane.props.testID === "office-sheet-horizontal-scroll")!;
    expect(verticalData.props.style.height).toBe(480);
    expect(horizontalData.props.style.height).toBe(480);
    expect(horizontalData.props.contentContainerStyle.height).toBe(480);
  });
});

describe("spreadsheet navigation and request ordering", () => {
  it("automatically fills a wide viewport across hidden columns without clicking the last column", async () => {
    hostMocks.readOfficeFile.mockImplementation(async (input) => {
      const preview = spreadsheetPreview(input);
      // Keep the real 106-column boundary while simulating narrow visible columns separated by hidden pairs.
      for (const sheet of preview.sheets!) {
        sheet.columnCount = 106;
        sheet.rows = sheet.rows.map((row) => row.slice(0, 106 - sheet.columnStart));
        if (sheet.rows.length) sheet.presentation = {
          columnWidths: sheet.rows[0].map((_, column) => (sheet.columnStart + column) % 3 === 2 ? 48 : 0),
          rowHeights: sheet.rows.map(() => 32),
        };
      }
      return preview;
    });
    officeStore.openAndJump({ filePath: "hidden.xlsx", ext: ".xlsx", cell: "A1" });
    const viewer = await renderViewer();
    const dataPane = viewer.root.findAllByType(ScrollView).find((pane) => pane.props.testID === "office-sheet-data-scroll")!;
    // An onLayout event is the only trigger: the user should not have to focus the last loaded cell.
    await act(async () => dataPane.props.onLayout({ nativeEvent: { layout: { width: 1200, height: 600 } } }));
    await flushQueries();
    expect(hostMocks.readOfficeFile).toHaveBeenLastCalledWith(expect.objectContaining({ columnStart: 0, columnLimit: 106 }));
    expect(viewer.root.findAllByType(Pressable).some((cell) => cell.props.accessibilityLabel === "DA1")).toBe(true);
  });

  it("buffers 50 rows around a jump and replaces paging buttons with automatic scrolling", async () => {
    hostMocks.readOfficeFile.mockImplementation(async (input) => spreadsheetPreview(input));
    officeStore.openAndJump({ filePath: "large.xlsx", ext: ".xlsx", sheet: "Giao dịch", cell: "AA5001" });
    const viewer = await renderViewer();
    expect(hostMocks.readOfficeFile).toHaveBeenLastCalledWith(expect.objectContaining({ rowStart: 4950, rowLimit: 101, columnStart: 20 }));
    const visibleCells = viewer.root.findAll((node) => node.type === Pressable && /^[A-Z]+\d+$/.test(node.props.accessibilityLabel ?? ""));
    expect(visibleCells).toHaveLength(101 * 12);
    expect(JSON.stringify(viewer.toJSON())).not.toContain("Dòng trước");
    // A local click neither re-reads the workbook nor moves the viewport away from the clicked cell.
    hostMocks.scrollTo.mockClear();
    await act(async () => visibleCells.find((cell) => cell.props.accessibilityLabel === "AA5001")!.props.onPress());
    await flushQueries();
    expect(hostMocks.readOfficeFile).toHaveBeenCalledTimes(1);
    expect(hostMocks.scrollTo).not.toHaveBeenCalled();
    await scrollViewer(viewer, "vertical", 2600);
    expect(hostMocks.readOfficeFile).toHaveBeenLastCalledWith(expect.objectContaining({ rowStart: 4950, rowLimit: 151 }));
    // Once the buffer reaches its cap, drop only rows above the viewport and retain the same visible pixel anchor.
    await scrollViewer(viewer, "vertical", 4200);
    expect(hostMocks.readOfficeFile).toHaveBeenLastCalledWith(expect.objectContaining({ rowStart: 5000, rowLimit: 151 }));
    expect(hostMocks.scrollTo).toHaveBeenCalledWith({ y: 2600, animated: false });
    await scrollViewer(viewer, "vertical", 30);
    expect(hostMocks.readOfficeFile).toHaveBeenLastCalledWith(expect.objectContaining({ rowStart: 4950, rowLimit: 151 }));
    expect(hostMocks.scrollTo).toHaveBeenCalledWith({ y: 1630, animated: false });
    // Horizontal scrolling remains usable after removing the column paging buttons as well.
    await scrollViewer(viewer, "horizontal", 700);
    expect(hostMocks.readOfficeFile).toHaveBeenLastCalledWith(expect.objectContaining({ columnStart: 20, columnLimit: 24 }));
  });

  it("keeps cells mounted during prefetch and suppresses duplicate reads while momentum continues", async () => {
    let finishPrefetch!: (document: OfficeFileData) => void;
    hostMocks.readOfficeFile.mockImplementation((input) => input.rowLimit === 151
      ? new Promise<OfficeFileData>((resolve) => { finishPrefetch = resolve; }) : Promise.resolve(spreadsheetPreview(input)));
    officeStore.openAndJump({ filePath: "large.xlsx", ext: ".xlsx" });
    const viewer = await renderViewer();
    await scrollViewer(viewer, "vertical", 2600);
    const prefetchInput = hostMocks.readOfficeFile.mock.calls[1][0];
    await scrollViewer(viewer, "vertical", 2620);
    // The first-sheet name was initially implicit; resolving it during scrolling must still preserve the old grid.
    expect(hostMocks.readOfficeFile).toHaveBeenCalledTimes(2);
    expect(viewer.root.findAllByType(Pressable).some((cell) => cell.props.accessibilityLabel === "A1")).toBe(true);
    expect(viewer.root.findAll((node) => String(node.type) === "Loading")).toHaveLength(0);
    hostMocks.scrollTo.mockClear();
    await act(async () => finishPrefetch(spreadsheetPreview(prefetchInput)));
    await flushQueries();
    expect(hostMocks.scrollTo).toHaveBeenCalledWith({ y: 2620, animated: false });
  });

  it("preserves the visible row when prepending rows with custom heights", async () => {
    hostMocks.readOfficeFile.mockImplementation(async (input) => {
      const preview = spreadsheetPreview(input);
      // Height depends on the absolute row so overlapping reads describe exactly the same worksheet geometry.
      for (const sheet of preview.sheets!) if (sheet.rows.length) sheet.presentation = {
        columnWidths: sheet.rows[0].map(() => 115), rowHeights: sheet.rows.map((_, index) =>
          (sheet.rowStart + index) % 2 === 0 ? 20 : 40),
      };
      return preview;
    });
    officeStore.openAndJump({ filePath: "large.xlsx", ext: ".xlsx", cell: "A1001" });
    const viewer = await renderViewer();
    hostMocks.scrollTo.mockClear();
    await scrollViewer(viewer, "vertical", 30);
    expect(hostMocks.readOfficeFile).toHaveBeenLastCalledWith(expect.objectContaining({ rowStart: 900, rowLimit: 151 }));
    expect(hostMocks.scrollTo).toHaveBeenCalledWith({ y: 1530, animated: false });
  });

  it("prefetches when a chat jump reaches a loaded edge without another wheel event", async () => {
    hostMocks.readOfficeFile.mockImplementation(async (input) => spreadsheetPreview(input));
    officeStore.openAndJump({ filePath: "large.xlsx", ext: ".xlsx", cell: "A1" });
    const viewer = await renderViewer();
    await scrollViewer(viewer, "vertical", 200);
    // Row 101 is already loaded, but reaching it clamps the scroll position at the old content boundary.
    await act(async () => officeStore.openAndJump({ filePath: "large.xlsx", ext: ".xlsx", cell: "A101" }));
    await flushQueries();
    expect(hostMocks.readOfficeFile).toHaveBeenLastCalledWith(expect.objectContaining({ rowStart: 0, rowLimit: 151 }));
    expect(hostMocks.scrollTo).toHaveBeenCalledWith({ y: 2632, animated: false });
  });

  it("lets a distant chat jump supersede a late scroll request in the same file", async () => {
    let finishPrefetch!: (document: OfficeFileData) => void;
    hostMocks.readOfficeFile.mockImplementation((input) => input.rowLimit === 151
      ? new Promise<OfficeFileData>((resolve) => { finishPrefetch = resolve; }) : Promise.resolve(spreadsheetPreview(input)));
    officeStore.openAndJump({ filePath: "large.xlsx", ext: ".xlsx", cell: "A1" });
    const viewer = await renderViewer();
    await scrollViewer(viewer, "vertical", 2600);
    const prefetchInput = hostMocks.readOfficeFile.mock.calls[1][0];
    // Query isolation must apply to a different range of the same workbook, not only to different files.
    await act(async () => officeStore.openAndJump({ filePath: "large.xlsx", ext: ".xlsx", sheet: "Giao dịch", cell: "AZ9001" }));
    await flushQueries();
    await act(async () => finishPrefetch(spreadsheetPreview(prefetchInput)));
    await flushQueries();
    const selectedCell = viewer.root.findAllByType(Pressable).find((cell) => cell.props.accessibilityLabel === "AZ9001")!;
    expect(selectedCell.findByType(Text).props.children).toBe("Giao dịch:9000:51");
  });

  it("switches sheets and returns to a distant chat jump", async () => {
    hostMocks.readOfficeFile.mockImplementation(async (input) => spreadsheetPreview(input));
    officeStore.openAndJump({ filePath: "large.xlsx", ext: ".xlsx", cell: "A1" });
    const viewer = await renderViewer();
    const summaryTab = viewer.root.findAllByType(Pressable).find((node) =>
      node.findAllByType(Text).some((textNode) => textNode.props.children === "Tóm tắt"))!;
    await act(async () => summaryTab.props.onPress());
    await flushQueries();
    expect(hostMocks.readOfficeFile).toHaveBeenLastCalledWith(expect.objectContaining({ sheet: "Tóm tắt", rowStart: 0 }));
    // A new issue supersedes manual navigation and loads its exact distant row/column window.
    await act(async () => officeStore.openAndJump({ filePath: "large.xlsx", ext: ".xlsx", sheet: "Giao dịch", cell: "AZ9001" }));
    await flushQueries();
    expect(hostMocks.readOfficeFile).toHaveBeenLastCalledWith(expect.objectContaining({ sheet: "Giao dịch", rowStart: 8950, columnStart: 45 }));
  });

  it("ignores an old file response that completes after the newly selected file", async () => {
    let completeFirstRead!: (document: OfficeFileData) => void;
    hostMocks.readOfficeFile.mockImplementation((input) => input.filePath === "first.xlsx"
      ? new Promise<OfficeFileData>((resolve) => { completeFirstRead = resolve; })
      : Promise.resolve(spreadsheetPreview(input)));
    officeStore.openAndJump({ filePath: "first.xlsx", ext: ".xlsx" });
    const viewer = await renderViewer();
    await act(async () => officeStore.openAndJump({ filePath: "second.xlsx", ext: ".xlsx", sheet: "Tóm tắt" }));
    await flushQueries();
    await act(async () => completeFirstRead(spreadsheetPreview({ filePath: "first.xlsx" })));
    await flushQueries();
    // Assert the visible cells, not merely query state: stale content must never reach the rendered pane.
    const firstVisibleCell = viewer.root.findAllByType(Pressable).find((node) => node.props.accessibilityLabel === "A1")!;
    expect(firstVisibleCell.findByType(Text).props.children).toBe("Tóm tắt:0:0");
  });

  it("shows read errors with a retry action", async () => {
    hostMocks.readOfficeFile.mockRejectedValue(new Error("File không tồn tại"));
    officeStore.openAndJump({ filePath: "missing.xlsx", ext: ".xlsx" });
    const viewer = await renderViewer();
    expect(JSON.stringify(viewer.toJSON())).toContain("File không tồn tại");
    expect(viewer.root.findAllByType(Pressable).some((node) =>
      node.findAllByType(Text).some((textNode) => textNode.props.children === "Thử lại"))).toBe(true);
  });

  it("keeps the current grid on a scroll read failure and can retry the same window", async () => {
    let rejectPrefetch = true;
    hostMocks.readOfficeFile.mockImplementation(async (input) => {
      if (input.rowLimit === 151 && rejectPrefetch) throw new Error("Không tải được dòng tiếp theo");
      return spreadsheetPreview(input);
    });
    officeStore.openAndJump({ filePath: "large.xlsx", ext: ".xlsx", cell: "A1" });
    const viewer = await renderViewer();
    await scrollViewer(viewer, "vertical", 2600);
    // A transient scroll failure must not remove all rows or lose the pending position correction.
    expect(viewer.root.findAllByType(Pressable).some((cell) => cell.props.accessibilityLabel === "A1")).toBe(true);
    expect(JSON.stringify(viewer.toJSON())).toContain("Không tải được dòng tiếp theo");
    const retry = viewer.root.findAllByType(Pressable).find((button) =>
      button.findAllByType(Text).some((textNode) => textNode.props.children === "Thử lại"))!;
    rejectPrefetch = false;
    await act(async () => retry.props.onPress());
    await flushQueries();
    expect(hostMocks.scrollTo).toHaveBeenCalledWith({ y: 2600, animated: false });
    expect(JSON.stringify(viewer.toJSON())).not.toContain("Không tải được dòng tiếp theo");
  });
});

describe("document frames and resource lifetime", () => {
  it("uses Paseo's PDF URL directly and includes the requested page without a text-extraction RPC", async () => {
    const stopWatching = vi.fn();
    const file = { path: "/workspace/report.pdf", extension: "pdf", watch: vi.fn(() => stopWatching) } as unknown as PluginFile;
    officeStore.openAndJump({ filePath: file.path, ext: ".pdf", page: 2 });
    const viewer = await renderViewer(<OfficeFilePanel file={file} workspaceId="workspace-1" theme={palette as never}
      host={{ id: "local", label: "Local" }} layout={{ compact: true, platform: "web" }} params={{ path: file.path }} tab={{} as never} revision={0} />);
    expect(viewer.root.findByType(DocumentFrame).props.url).toBe("blob:host-pdf#page=2");
    expect(hostMocks.readOfficeFile).not.toHaveBeenCalled();
    await act(async () => viewer.unmount());
    expect(stopWatching).toHaveBeenCalledTimes(1);
    renderedViewer = undefined;
  });

  it("frames Word HTML and releases replaced blob URLs on unmount", async () => {
    const revokeUrl = vi.spyOn(URL, "revokeObjectURL");
    hostMocks.readOfficeFile.mockResolvedValue({ filePath: "contract.docx", fileType: "docx",
      htmlContent: "<h1>Hợp đồng</h1><table><tr><td>Giá</td></tr></table>", textContent: "Hợp đồng" });
    officeStore.openAndJump({ filePath: "contract.docx", ext: ".docx" });
    const viewer = await renderViewer();
    const documentUrl = viewer.root.findByType(DocumentFrame).props.url;
    const previewHtml = await (await fetch(documentUrl)).text();
    expect(previewHtml).toContain("<table>");
    expect(previewHtml).toContain("default-src 'none'");
    await act(async () => viewer.unmount());
    expect(revokeUrl).toHaveBeenCalledWith(documentUrl);
    renderedViewer = undefined;
  });

  it("creates a typed PDF blob for fallback workspace previews", async () => {
    // A MIME-typed blob is required for DocumentFrame to invoke the PDF viewer instead of treating it as HTML.
    const pdfBytes = Buffer.from("%PDF-1.4\nExample");
    const preview = createOfficeDocumentUrl({ documentBase64: pdfBytes.toString("base64") })!;
    try {
      const downloadedPdf = await fetch(preview.url);
      expect(downloadedPdf.headers.get("content-type")).toBe("application/pdf");
      expect(Buffer.from(await downloadedPdf.arrayBuffer())).toEqual(pdfBytes);
    } finally {
      preview.dispose();
    }
  });
});
