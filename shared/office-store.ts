// Coordinates an active file selection and agent-reported validation issues across Chat, Tree, and Viewer.
export interface OfficeTarget {
  filePath: string;
  ext: string;
  sheet?: string;
  cell?: string;
  page?: number;
  reason?: string;
  timestamp: number;
}

export type OfficeIssueInput = Omit<OfficeTarget, "timestamp" | "filePath">;

export interface OfficeStoreSnapshot {
  activeTarget: OfficeTarget | null;
  issuesByFile: Record<string, OfficeTarget[]>;
}

type StoreSubscriber = (state: OfficeStoreSnapshot) => void;

const storeSubscribers = new Set<StoreSubscriber>();
let currentActiveTarget: OfficeTarget | null = null;
let reportedIssuesByFile: Record<string, OfficeTarget[]> = {};

// Broadcast state snapshots immutably so React subscribers always detect reference updates.
function notifyStoreSubscribers(): void {
  const currentSnapshot: OfficeStoreSnapshot = {
    activeTarget: currentActiveTarget,
    issuesByFile: reportedIssuesByFile,
  };
  storeSubscribers.forEach((subscriberCallback) => subscriberCallback(currentSnapshot));
}

export const officeStore = {
  // Attach a timestamp to each issue so repeated agent scans trigger fresh UI badge updates.
  registerIssues(filePath: string, issues: Array<OfficeIssueInput>): void {
    const registrationTimestamp = Date.now();
    reportedIssuesByFile = {
      ...reportedIssuesByFile,
      [filePath]: issues.map((issueDetail) => ({
        ...issueDetail,
        filePath,
        timestamp: registrationTimestamp,
      })),
    };
    notifyStoreSubscribers();
  },

  // Stamp each navigation request with Date.now() so clicking the same cell twice still re-scrolls the viewer.
  openAndJump(target: Omit<OfficeTarget, "timestamp">): void {
    currentActiveTarget = {
      ...target,
      timestamp: Date.now(),
    };
    notifyStoreSubscribers();
  },

  // Allow resetting state between workspace switches or test runs without leaking stale targets.
  reset(): void {
    currentActiveTarget = null;
    reportedIssuesByFile = {};
    notifyStoreSubscribers();
  },

  getState(): OfficeStoreSnapshot {
    return {
      activeTarget: currentActiveTarget,
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

// Convert A1-style Excel references (e.g. "F14", "AA10") into zero-based row/col indices for grid scrolling.
export function parseCellRef(cellRef?: string): { col: number; row: number } | null {
  if (!cellRef) return null;
  const trimmedReference = cellRef.trim().toUpperCase();
  const referenceMatch = trimmedReference.match(/^([A-Z]+)(\d+)$/);
  if (!referenceMatch) return null;

  const [, columnLetters, rowDigits] = referenceMatch;
  const parsedRowNumber = Number.parseInt(rowDigits, 10);
  // Reject row 0 ("A0") because Excel row numbering is strictly 1-based.
  if (!Number.isInteger(parsedRowNumber) || parsedRowNumber < 1) return null;

  // Base-26 bijective decoding where 'A' is 1 and 'Z' is 26.
  let oneBasedColumnNumber = 0;
  for (let charIndex = 0; charIndex < columnLetters.length; charIndex++) {
    oneBasedColumnNumber = oneBasedColumnNumber * 26 + (columnLetters.charCodeAt(charIndex) - 64);
  }
  return {
    col: oneBasedColumnNumber - 1,
    row: parsedRowNumber - 1,
  };
}

// Convert a zero-based column index back into Excel's A, B, ..., Z, AA header label.
export function formatColumnLabel(zeroBasedColIndex: number): string {
  let remainingColNumber = Math.max(0, zeroBasedColIndex) + 1;
  let columnHeaderLabel = "";
  // Repeatedly extract the least-significant base-26 digit with 1-based offset adjustment.
  while (remainingColNumber > 0) {
    const alphabetOffset = (remainingColNumber - 1) % 26;
    columnHeaderLabel = String.fromCharCode(65 + alphabetOffset) + columnHeaderLabel;
    remainingColNumber = Math.floor((remainingColNumber - 1) / 26);
  }
  return columnHeaderLabel;
}
