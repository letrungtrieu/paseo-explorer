// Coordinates active file tabs, grid selections (cell/row/col/range), and Agent-reported issues across Chat, Tree, and Viewer.
export interface OfficeTarget {
  filePath: string;
  ext: string;
  sheet?: string;
  // Supports single cells ("E3"), ranges ("B4:D8"), whole rows ("Row:3" / "Dòng 3"), or whole columns ("Col:E" / "Cột E").
  cell?: string;
  page?: number;
  reason?: string;
  timestamp: number;
}

export type OfficeIssueInput = Omit<OfficeTarget, "timestamp" | "filePath">;

export interface OfficeStoreSnapshot {
  activeTarget: OfficeTarget | null;
  openTargets: OfficeTarget[];
  issuesByFile: Record<string, OfficeTarget[]>;
}

export interface ParsedGridSelection {
  kind: "cell" | "row" | "col" | "range";
  label: string;
  startRow?: number;
  endRow?: number;
  startCol?: number;
  endCol?: number;
}

type StoreSubscriber = (state: OfficeStoreSnapshot) => void;
type WorkspaceTabOpener = (workspaceId?: string) => void;

const storeSubscribers = new Set<StoreSubscriber>();
let currentActiveTarget: OfficeTarget | null = null;
let currentOpenTargets: OfficeTarget[] = [];
let reportedIssuesByFile: Record<string, OfficeTarget[]> = {};
let registeredWorkspaceTabOpener: WorkspaceTabOpener | null = null;

// Normalize path separators so relative paths ("examples/a.xlsx") match absolute workspace paths ("/repo/examples/a.xlsx").
export function isMatchingOfficePath(firstPath: string, secondPath: string): boolean {
  const normalizedFirst = firstPath.replace(/\\/g, "/").replace(/^\.\//, "").trim();
  const normalizedSecond = secondPath.replace(/\\/g, "/").replace(/^\.\//, "").trim();
  if (!normalizedFirst || !normalizedSecond) return false;
  if (normalizedFirst === normalizedSecond) return true;
  return (
    normalizedFirst.endsWith(`/${normalizedSecond}`) ||
    normalizedSecond.endsWith(`/${normalizedFirst}`)
  );
}

// Look up reported issues for a file even if one side used a relative path and the other used an absolute path.
export function findIssuesForFilePath(
  issuesMap: Record<string, OfficeTarget[]>,
  candidateFilePath?: string,
): OfficeTarget[] {
  if (!candidateFilePath) return [];
  if (issuesMap[candidateFilePath]) return issuesMap[candidateFilePath];
  for (const [registeredPath, registeredIssues] of Object.entries(issuesMap)) {
    if (isMatchingOfficePath(registeredPath, candidateFilePath)) {
      return registeredIssues;
    }
  }
  return [];
}

// Broadcast immutable state snapshots so React subscribers in the Sidebar and Main Tab re-render predictably.
function notifyStoreSubscribers(): void {
  const currentSnapshot: OfficeStoreSnapshot = {
    activeTarget: currentActiveTarget,
    openTargets: currentOpenTargets,
    issuesByFile: reportedIssuesByFile,
  };
  storeSubscribers.forEach((subscriberCallback) => subscriberCallback(currentSnapshot));
}

export const officeStore = {
  // Allow index.client.tsx to bind client.openPanel(..., { location: "workspace" }) for automatic tab opening.
  setWorkspaceTabOpener(openerCallback: WorkspaceTabOpener | null): void {
    registeredWorkspaceTabOpener = openerCallback;
  },

  // Merge or replace issues for a file path (matching relative and absolute variants) so badges stay accurate.
  registerIssues(filePath: string, issues: Array<OfficeIssueInput>): void {
    const registrationTimestamp = Date.now();
    const existingMatchedKey = Object.keys(reportedIssuesByFile).find((existingKey) =>
      isMatchingOfficePath(existingKey, filePath),
    );
    const canonicalPathKey =
      existingMatchedKey && existingMatchedKey.length > filePath.length
        ? existingMatchedKey
        : filePath;

    const nextIssuesMap = { ...reportedIssuesByFile };
    if (existingMatchedKey && existingMatchedKey !== canonicalPathKey) {
      delete nextIssuesMap[existingMatchedKey];
    }

    nextIssuesMap[canonicalPathKey] = issues.map((issueDetail) => ({
      ...issueDetail,
      filePath: canonicalPathKey,
      timestamp: registrationTimestamp,
    }));
    reportedIssuesByFile = nextIssuesMap;
    notifyStoreSubscribers();
  },

  // Activate a file/cell target and open the Office Viewer in a Main Workspace Tab.
  openAndJump(
    target: Omit<OfficeTarget, "timestamp">,
    options?: { workspaceId?: string; openMainTab?: boolean },
  ): void {
    const nextTarget: OfficeTarget = {
      ...target,
      timestamp: Date.now(),
    };
    currentActiveTarget = nextTarget;

    const existingTabIndex = currentOpenTargets.findIndex((openEntry) =>
      isMatchingOfficePath(openEntry.filePath, nextTarget.filePath),
    );
    if (existingTabIndex >= 0) {
      const preservedFilePath =
        currentOpenTargets[existingTabIndex].filePath.length > nextTarget.filePath.length
          ? currentOpenTargets[existingTabIndex].filePath
          : nextTarget.filePath;
      const mergedTarget = { ...nextTarget, filePath: preservedFilePath };
      currentActiveTarget = mergedTarget;
      currentOpenTargets = currentOpenTargets.map((openEntry, entryIndex) =>
        entryIndex === existingTabIndex ? mergedTarget : openEntry,
      );
    } else {
      currentOpenTargets = [...currentOpenTargets, nextTarget];
    }

    notifyStoreSubscribers();

    if (options?.openMainTab !== false && registeredWorkspaceTabOpener) {
      registeredWorkspaceTabOpener(options?.workspaceId);
    }
  },

  // Close a file tab in the main viewer and fall back to the nearest remaining open file tab.
  closeFileTab(filePath: string): void {
    currentOpenTargets = currentOpenTargets.filter(
      (openEntry) => !isMatchingOfficePath(openEntry.filePath, filePath),
    );
    if (currentActiveTarget && isMatchingOfficePath(currentActiveTarget.filePath, filePath)) {
      currentActiveTarget =
        currentOpenTargets.length > 0 ? currentOpenTargets[currentOpenTargets.length - 1] : null;
    }
    notifyStoreSubscribers();
  },

  reset(): void {
    currentActiveTarget = null;
    currentOpenTargets = [];
    reportedIssuesByFile = {};
    notifyStoreSubscribers();
  },

  getState(): OfficeStoreSnapshot {
    return {
      activeTarget: currentActiveTarget,
      openTargets: currentOpenTargets,
      issuesByFile: reportedIssuesByFile,
    };
  },

  // Return a void cleanup function rather than Set.delete's boolean to satisfy React 19 useEffect destructor types.
  subscribe(subscriberCallback: StoreSubscriber): () => void {
    storeSubscribers.add(subscriberCallback);
    return () => {
      storeSubscribers.delete(subscriberCallback);
    };
  },
};

// Decode Excel column letters ("A" -> 0, "E" -> 4, "Z" -> 25, "AA" -> 26).
export function parseColumnLetters(columnLetters: string): number | null {
  const normalizedLetters = columnLetters.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(normalizedLetters)) return null;
  let oneBasedColumnNumber = 0;
  for (let charIndex = 0; charIndex < normalizedLetters.length; charIndex++) {
    oneBasedColumnNumber = oneBasedColumnNumber * 26 + (normalizedLetters.charCodeAt(charIndex) - 64);
  }
  return oneBasedColumnNumber - 1;
}

// Convert A1-style Excel references (e.g. "F14", "E3") into zero-based row/col indices.
export function parseCellRef(cellRef?: string): { col: number; row: number } | null {
  if (!cellRef) return null;
  const trimmedReference = cellRef.trim().toUpperCase();
  const referenceMatch = trimmedReference.match(/^([A-Z]+)(\d+)$/);
  if (!referenceMatch) return null;

  const [, columnLetters, rowDigits] = referenceMatch;
  const parsedRowNumber = Number.parseInt(rowDigits, 10);
  if (!Number.isInteger(parsedRowNumber) || parsedRowNumber < 1) return null;

  const zeroBasedCol = parseColumnLetters(columnLetters);
  if (zeroBasedCol === null) return null;

  return {
    col: zeroBasedCol,
    row: parsedRowNumber - 1,
  };
}

// Parse cell ("E3"), range ("B4:E8"), whole row ("Row:3", "Dòng 3", "3:3"), or whole column ("Col:E", "Cột E", "E:E").
export function parseGridSelection(selectionRef?: string): ParsedGridSelection | null {
  if (!selectionRef) return null;
  const rawRef = selectionRef.trim();
  if (!rawRef) return null;

  // 1. Check single cell (e.g. "E3", "F14")
  const singleCell = parseCellRef(rawRef);
  if (singleCell) {
    return {
      kind: "cell",
      label: rawRef.toUpperCase(),
      startRow: singleCell.row,
      endRow: singleCell.row,
      startCol: singleCell.col,
      endCol: singleCell.col,
    };
  }

  // 2. Check cell range (e.g. "B4:E8")
  const rangeMatch = rawRef.toUpperCase().match(/^([A-Z]+\d+)\s*:\s*([A-Z]+\d+)$/);
  if (rangeMatch) {
    const startCoords = parseCellRef(rangeMatch[1]);
    const endCoords = parseCellRef(rangeMatch[2]);
    if (startCoords && endCoords) {
      return {
        kind: "range",
        label: `${rangeMatch[1]}:${rangeMatch[2]}`,
        startRow: Math.min(startCoords.row, endCoords.row),
        endRow: Math.max(startCoords.row, endCoords.row),
        startCol: Math.min(startCoords.col, endCoords.col),
        endCol: Math.max(startCoords.col, endCoords.col),
      };
    }
  }

  // 3. Check whole row (e.g. "Row:3", "Row 3", "Dòng 3", "Hàng 3", "3:3")
  const rowMatch =
    rawRef.match(/^(?:row|dòng|dong|hàng|hang)\s*[:#-]?\s*(\d+)$/i) ??
    rawRef.match(/^(\d+)\s*:\s*(\d+)$/);
  if (rowMatch) {
    const firstRowNum = Number.parseInt(rowMatch[1], 10);
    const secondRowNum = rowMatch[2] ? Number.parseInt(rowMatch[2], 10) : firstRowNum;
    if (firstRowNum >= 1 && secondRowNum >= 1) {
      const minRow = Math.min(firstRowNum, secondRowNum) - 1;
      const maxRow = Math.max(firstRowNum, secondRowNum) - 1;
      return {
        kind: "row",
        label: minRow === maxRow ? `Dòng ${minRow + 1}` : `Dòng ${minRow + 1}:${maxRow + 1}`,
        startRow: minRow,
        endRow: maxRow,
      };
    }
  }

  // 4. Check whole column (e.g. "Col:E", "Column E", "Cột E", "E:E")
  const colMatch =
    rawRef.match(/^(?:col|column|cột|cot)\s*[:#-]?\s*([A-Za-z]+)$/i) ??
    rawRef.match(/^([A-Za-z]+)\s*:\s*([A-Za-z]+)$/);
  if (colMatch) {
    const firstColIndex = parseColumnLetters(colMatch[1]);
    const secondColIndex = colMatch[2] ? parseColumnLetters(colMatch[2]) : firstColIndex;
    if (firstColIndex !== null && secondColIndex !== null) {
      const minCol = Math.min(firstColIndex, secondColIndex);
      const maxCol = Math.max(firstColIndex, secondColIndex);
      return {
        kind: "col",
        label:
          minCol === maxCol
            ? `Cột ${formatColumnLabel(minCol)}`
            : `Cột ${formatColumnLabel(minCol)}:${formatColumnLabel(maxCol)}`,
        startCol: minCol,
        endCol: maxCol,
      };
    }
  }

  return null;
}

// Parse one or more comma/semicolon-separated grid coordinates from a single clicked link (e.g. "E3", "E3, G5", "B4:D8").
export function parseGridSelections(selectionRef?: string): ParsedGridSelection[] {
  if (!selectionRef) return [];
  return selectionRef
    .split(/[,;]+/)
    .map((segment) => parseGridSelection(segment))
    .filter((parsedItem): parsedItem is ParsedGridSelection => parsedItem !== null);
}

// Determine whether a specific (rowIndex, colIndex) cell falls inside a parsed cell/row/col/range selection.
export function isCellSelected(
  selection: ParsedGridSelection | null,
  rowIndex: number,
  colIndex: number,
): boolean {
  if (!selection) return false;
  if (selection.kind === "row") {
    return (
      selection.startRow !== undefined &&
      selection.endRow !== undefined &&
      rowIndex >= selection.startRow &&
      rowIndex <= selection.endRow
    );
  }
  if (selection.kind === "col") {
    return (
      selection.startCol !== undefined &&
      selection.endCol !== undefined &&
      colIndex >= selection.startCol &&
      colIndex <= selection.endCol
    );
  }
  return (
    selection.startRow !== undefined &&
    selection.endRow !== undefined &&
    selection.startCol !== undefined &&
    selection.endCol !== undefined &&
    rowIndex >= selection.startRow &&
    rowIndex <= selection.endRow &&
    colIndex >= selection.startCol &&
    colIndex <= selection.endCol
  );
}

// Check whether (rowIndex, colIndex) belongs to any selection described by the currently clicked link.
export function isCellInSelections(
  selections: readonly ParsedGridSelection[],
  rowIndex: number,
  colIndex: number,
): boolean {
  return selections.some((selectionItem) => isCellSelected(selectionItem, rowIndex, colIndex));
}

// Convert a zero-based column index back into Excel's A, B, ..., Z, AA header label.
export function formatColumnLabel(zeroBasedColIndex: number): string {
  let remainingColNumber = Math.max(0, zeroBasedColIndex) + 1;
  let columnHeaderLabel = "";
  while (remainingColNumber > 0) {
    const alphabetOffset = (remainingColNumber - 1) % 26;
    columnHeaderLabel = String.fromCharCode(65 + alphabetOffset) + columnHeaderLabel;
    remainingColNumber = Math.floor((remainingColNumber - 1) / 26);
  }
  return columnHeaderLabel;
}
