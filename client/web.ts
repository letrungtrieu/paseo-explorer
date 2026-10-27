import { Platform } from "react-native";
import {
  findIssuesForFilePath,
  isMatchingOfficePath,
  officeStore,
  type OfficeStoreSnapshot,
} from "../shared/office-store";
import { ICON_CONFIG, type FileBadgePalette } from "./file-icons";

// Minimal DOM interfaces scoped strictly to client/web.ts so tsconfig.json stays free of global DOM lib pollution.
interface WebStyleDeclaration {
  [property: string]: string | undefined;
}

interface WebDomElement {
  textContent: string | null;
  style: WebStyleDeclaration;
  parentElement: WebDomElement | null;
  previousElementSibling: WebDomElement | null;
  firstElementChild: WebDomElement | null;
  getAttribute(attributeName: string): string | null;
  setAttribute(attributeName: string, attributeValue: string): void;
  removeAttribute(attributeName: string): void;
  querySelector(selector: string): WebDomElement | null;
  querySelectorAll(selector: string): ArrayLike<WebDomElement>;
  appendChild(childNode: WebDomElement): WebDomElement;
  remove(): void;
  click(): void;
}

interface WebDomDocument {
  body: WebDomElement | null;
  createElement(tagName: string): WebDomElement;
  querySelector(selector: string): WebDomElement | null;
  querySelectorAll(selector: string): ArrayLike<WebDomElement>;
}

interface WebMutationObserver {
  observe(
    targetNode: WebDomElement,
    options: { childList?: boolean; subtree?: boolean; characterData?: boolean },
  ): void;
  disconnect(): void;
}

interface WebGlobalEnvironment {
  document?: WebDomDocument;
  MutationObserver?: new (callback: () => void) => WebMutationObserver;
  getComputedStyle?: (element: WebDomElement) => WebStyleDeclaration;
  requestAnimationFrame?: (callback: () => void) => number;
  cancelAnimationFrame?: (handle: number) => void;
}

const LEGACY_EXPLORER_TAB_CLOSE_SELECTOR =
  '[data-testid="workspace-plugin-close-plugin_workspace_14_paseo-explorer_15_office-explorer"]';
const NATIVE_FILES_TAB_SELECTOR = '[data-testid="workspace-tab-files"]';
const NATIVE_FILE_ROW_SELECTOR =
  '[data-testid^="file-explorer-row-"]:not([data-testid$="-name"])';

// Only override extensions that lack dedicated icons in Paseo's built-in MaterialFileIcon table.
const NATIVE_TREE_OFFICE_EXTENSIONS = new Set([
  ".xlsx",
  ".xls",
  ".csv",
  ".docx",
  ".doc",
  ".pdf",
]);

// Extract lowercase dot-prefixed extension from a file or folder label in Paseo's native Files tree.
export function resolveNativeTreeOfficeBadge(entryDisplayName: string): {
  extension: string;
  badgePalette: FileBadgePalette;
} | null {
  const trimmedName = entryDisplayName.trim();
  const lastDotPosition = trimmedName.lastIndexOf(".");
  if (lastDotPosition <= 0 || lastDotPosition === trimmedName.length - 1) {
    return null;
  }
  const normalizedExtension = trimmedName.slice(lastDotPosition).toLowerCase();
  if (!NATIVE_TREE_OFFICE_EXTENSIONS.has(normalizedExtension)) {
    return null;
  }
  const badgePalette = ICON_CONFIG[normalizedExtension];
  if (!badgePalette) {
    return null;
  }
  return { extension: normalizedExtension, badgePalette };
}

// Parse the numeric FlatList index from Paseo's `file-explorer-row-<index>` testID so rows sort in tree order.
function parseExplorerRowOrderIndex(rowElement: WebDomElement): number {
  const rawTestId = rowElement.getAttribute("data-testid") ?? "";
  const matchResult = rawTestId.match(/^file-explorer-row-(\d+)$/);
  return matchResult ? Number.parseInt(matchResult[1], 10) : Number.MAX_SAFE_INTEGER;
}

// Compute zero-based directory nesting depth from Paseo's `12 + depth * 12` pixel row indentation formula.
function resolveExplorerRowDepth(
  rowElement: WebDomElement,
  webEnv: WebGlobalEnvironment,
): number {
  const inlinePaddingLeft = rowElement.style?.paddingLeft;
  const computedPaddingLeft =
    inlinePaddingLeft || webEnv.getComputedStyle?.(rowElement)?.paddingLeft || "12px";
  const paddingPixels = Number.parseFloat(computedPaddingLeft);
  if (!Number.isFinite(paddingPixels) || paddingPixels <= 12) {
    return 0;
  }
  return Math.max(0, Math.round((paddingPixels - 12) / 12));
}

// Detect directory rows by checking for Paseo's Lucide ChevronRight path inside the leading icon slot.
function isDirectoryExplorerRow(iconContainerElement: WebDomElement): boolean {
  return iconContainerElement.querySelector('path[d*="9 18 6-6-6-6"]') !== null;
}

// Apply or remove the colored XLS/CSV/DOC/PDF badge inside a native Files tree row's icon container.
function syncRowOfficeIconBadge(
  domDocument: WebDomElement | WebDomDocument,
  iconContainerElement: WebDomElement,
  officeBadgeInfo: { extension: string; badgePalette: FileBadgePalette } | null,
  isDirectoryRow: boolean,
): void {
  const existingBadgeNode = iconContainerElement.querySelector('[data-office-file-badge="true"]');
  const nativeMaterialSvgNode = iconContainerElement.querySelector("svg");

  if (!officeBadgeInfo || isDirectoryRow) {
    // Clean up custom badge when virtualized FlatList recycles an Office row for a code file or folder.
    if (existingBadgeNode) {
      existingBadgeNode.remove();
      iconContainerElement.style.width = "";
      iconContainerElement.style.minWidth = "";
      if (nativeMaterialSvgNode) {
        nativeMaterialSvgNode.style.display = "";
      }
    }
    return;
  }

  const { extension, badgePalette } = officeBadgeInfo;
  if (nativeMaterialSvgNode && nativeMaterialSvgNode.style.display !== "none") {
    nativeMaterialSvgNode.style.display = "none";
  }
  if (iconContainerElement.style.width !== "26px") {
    iconContainerElement.style.width = "26px";
    iconContainerElement.style.minWidth = "26px";
  }

  let targetBadgeElement = existingBadgeNode;
  if (!targetBadgeElement && "createElement" in domDocument) {
    targetBadgeElement = domDocument.createElement("span");
    targetBadgeElement.setAttribute("data-office-file-badge", "true");
    iconContainerElement.appendChild(targetBadgeElement);
  }
  if (!targetBadgeElement) return;

  // Skip style writes if this row already displays the exact badge for the current extension.
  if (targetBadgeElement.getAttribute("data-office-ext") === extension) {
    return;
  }

  targetBadgeElement.setAttribute("data-office-ext", extension);
  targetBadgeElement.textContent = badgePalette.label;
  targetBadgeElement.style.display = "inline-flex";
  targetBadgeElement.style.alignItems = "center";
  targetBadgeElement.style.justifyContent = "center";
  targetBadgeElement.style.width = "26px";
  targetBadgeElement.style.height = "15px";
  targetBadgeElement.style.backgroundColor = badgePalette.bg;
  targetBadgeElement.style.border = `1px solid ${badgePalette.border}`;
  targetBadgeElement.style.borderRadius = "3px";
  targetBadgeElement.style.color = badgePalette.fg;
  targetBadgeElement.style.fontSize = "8px";
  targetBadgeElement.style.fontWeight = "800";
  targetBadgeElement.style.letterSpacing = "0.3px";
  targetBadgeElement.style.lineHeight = "1";
  targetBadgeElement.style.fontFamily = "system-ui, -apple-system, sans-serif";
  targetBadgeElement.style.userSelect = "none";
  targetBadgeElement.style.boxSizing = "border-box";
}

// Apply or remove the amber `⚠ N` issue counter pill beside affected filenames in Paseo's native Files tree.
function syncRowIssueWarningPill(
  domDocument: WebDomDocument,
  entryInfoElement: WebDomElement,
  issueCount: number,
  isDirectoryRow: boolean,
): void {
  const existingPillNode = entryInfoElement.querySelector('[data-office-issue-pill="true"]');
  if (isDirectoryRow || issueCount <= 0) {
    if (existingPillNode) {
      existingPillNode.remove();
    }
    return;
  }

  let targetPillElement = existingPillNode;
  if (!targetPillElement) {
    targetPillElement = domDocument.createElement("span");
    targetPillElement.setAttribute("data-office-issue-pill", "true");
    entryInfoElement.appendChild(targetPillElement);
  }

  const serializedCount = String(issueCount);
  if (targetPillElement.getAttribute("data-issue-count") === serializedCount) {
    return;
  }

  targetPillElement.setAttribute("data-issue-count", serializedCount);
  targetPillElement.textContent = `⚠ ${issueCount}`;
  targetPillElement.style.display = "inline-flex";
  targetPillElement.style.alignItems = "center";
  targetPillElement.style.backgroundColor = "rgba(245, 158, 11, 0.18)";
  targetPillElement.style.border = "1px solid rgba(245, 158, 11, 0.45)";
  targetPillElement.style.color = "#f59e0b";
  targetPillElement.style.fontSize = "9px";
  targetPillElement.style.fontWeight = "700";
  targetPillElement.style.padding = "1px 5px";
  targetPillElement.style.borderRadius = "8px";
  targetPillElement.style.marginLeft = "6px";
  targetPillElement.style.flexShrink = "0";
  targetPillElement.style.lineHeight = "1.2";
  targetPillElement.style.fontFamily = "system-ui, -apple-system, sans-serif";
}

// Decorate all visible rows in Paseo's native Files tab with Office icons and Agent issue count badges.
export function decorateNativeFilesTreeDom(
  domDocument: WebDomDocument,
  storeSnapshot: OfficeStoreSnapshot,
  webEnv: WebGlobalEnvironment = globalThis as unknown as WebGlobalEnvironment,
): void {
  // Automatically close any leftover custom "Office Explorer" sidebar tab from older plugin sessions.
  const legacyExplorerCloseButton = domDocument.querySelector(LEGACY_EXPLORER_TAB_CLOSE_SELECTOR);
  if (legacyExplorerCloseButton) {
    legacyExplorerCloseButton.click();
    domDocument.querySelector(NATIVE_FILES_TAB_SELECTOR)?.click();
  }

  const rawRowNodeList = domDocument.querySelectorAll(NATIVE_FILE_ROW_SELECTOR);
  const sortedRowElements = Array.from(rawRowNodeList).sort(
    (firstRow, secondRow) =>
      parseExplorerRowOrderIndex(firstRow) - parseExplorerRowOrderIndex(secondRow),
  );

  const depthPathStack: string[] = [];

  for (const rowElement of sortedRowElements) {
    const nameElement = rowElement.querySelector('[data-testid$="-name"]');
    const entryInfoElement = nameElement?.parentElement ?? null;
    const iconContainerElement = nameElement?.previousElementSibling ?? null;
    if (!nameElement || !entryInfoElement || !iconContainerElement) {
      continue;
    }

    const entryName = (nameElement.textContent ?? "").trim();
    if (!entryName) continue;

    const rowDepth = resolveExplorerRowDepth(rowElement, webEnv);
    depthPathStack[rowDepth] = entryName;
    depthPathStack.length = rowDepth + 1;
    // Filter out empty ancestor slots when virtualized FlatList scrolls past collapsed/offscreen parent folders.
    const reconstructedRelativePath = depthPathStack.filter(Boolean).join("/");

    const isDirectoryRow = isDirectoryExplorerRow(iconContainerElement);
    const officeBadgeInfo = resolveNativeTreeOfficeBadge(entryName);

    // Record reconstructed path on the row so Chat issue links can click the native row directly when visible.
    if (rowElement.getAttribute("data-office-relative-path") !== reconstructedRelativePath) {
      rowElement.setAttribute("data-office-relative-path", reconstructedRelativePath);
    }

    syncRowOfficeIconBadge(domDocument, iconContainerElement, officeBadgeInfo, isDirectoryRow);

    const matchedIssuesByRelativePath = findIssuesForFilePath(
      storeSnapshot.issuesByFile,
      reconstructedRelativePath,
    );
    const matchedIssues =
      matchedIssuesByRelativePath.length > 0
        ? matchedIssuesByRelativePath
        : officeBadgeInfo
          ? findIssuesForFilePath(storeSnapshot.issuesByFile, entryName)
          : [];

    syncRowIssueWarningPill(
      domDocument,
      entryInfoElement,
      matchedIssues.length,
      isDirectoryRow,
    );
  }
}

// Attempt to focus an already-open native file tab or click the file row directly in Paseo's native Files tree.
export function tryOpenNativeFileTabFromDom(
  targetFilePath: string,
  domDocument?: WebDomDocument,
): boolean {
  if (Platform.OS !== "web" && !domDocument) return false;
  const webEnv = globalThis as unknown as WebGlobalEnvironment;
  const activeDocument = domDocument ?? webEnv.document;
  if (!activeDocument) return false;

  // 1. Check if a native Paseo file tab (`workspace-tab-file_<path>`) is already open in the workspace tab bar.
  const openFileTabs = Array.from(
    activeDocument.querySelectorAll('[data-testid^="workspace-tab-file_"]'),
  );
  for (const fileTabElement of openFileTabs) {
    const tabTestId = fileTabElement.getAttribute("data-testid") ?? "";
    const tabFilePath = tabTestId.replace(/^workspace-tab-file_/, "");
    if (isMatchingOfficePath(tabFilePath, targetFilePath)) {
      fileTabElement.click();
      return true;
    }
  }

  // 2. Otherwise check if the file row is currently visible inside Paseo's native Files tree and click it.
  const treeRows = Array.from(activeDocument.querySelectorAll(NATIVE_FILE_ROW_SELECTOR));
  for (const rowElement of treeRows) {
    const rowRelativePath =
      rowElement.getAttribute("data-office-relative-path") ??
      rowElement.querySelector('[data-testid$="-name"]')?.textContent?.trim() ??
      "";
    if (rowRelativePath && isMatchingOfficePath(rowRelativePath, targetFilePath)) {
      rowElement.click();
      return true;
    }
  }

  return false;
}

// Start a MutationObserver + store subscription that keeps Paseo's native Files tab decorated with Office badges.
export function startNativeFilesTreeEnhancer(): () => void {
  if (Platform.OS !== "web") {
    return () => {};
  }

  const webEnv = globalThis as unknown as WebGlobalEnvironment;
  const domDocument = webEnv.document;
  if (!domDocument?.body || !webEnv.MutationObserver) {
    return () => {};
  }

  let scheduledAnimationFrame: number | null = null;
  let isApplyingDecorations = false;

  const runDecorationPass = () => {
    scheduledAnimationFrame = null;
    if (isApplyingDecorations) return;
    isApplyingDecorations = true;
    try {
      decorateNativeFilesTreeDom(domDocument, officeStore.getState(), webEnv);
    } finally {
      isApplyingDecorations = false;
    }
  };

  const scheduleDecorationPass = () => {
    if (isApplyingDecorations || scheduledAnimationFrame !== null) return;
    if (webEnv.requestAnimationFrame) {
      scheduledAnimationFrame = webEnv.requestAnimationFrame(runDecorationPass);
    } else {
      runDecorationPass();
    }
  };

  const domObserver = new webEnv.MutationObserver(() => {
    scheduleDecorationPass();
  });

  domObserver.observe(domDocument.body, {
    childList: true,
    subtree: true,
  });

  const unsubscribeStore = officeStore.subscribe(() => {
    scheduleDecorationPass();
  });

  // Perform an immediate decoration pass on startup so already-mounted Files rows get badges right away.
  runDecorationPass();

  return () => {
    domObserver.disconnect();
    unsubscribeStore();
    if (scheduledAnimationFrame !== null && webEnv.cancelAnimationFrame) {
      webEnv.cancelAnimationFrame(scheduledAnimationFrame);
    }
  };
}
