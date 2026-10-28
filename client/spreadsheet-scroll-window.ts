import type { WorksheetSheet } from "../shared/viewer";
import {
  WORKSHEET_MAX_WINDOW_COLUMNS, WORKSHEET_MAX_WINDOW_ROWS, WORKSHEET_ROW_BUFFER,
  WORKSHEET_WINDOW_COLUMNS, type SpreadsheetWindow,
} from "../shared/spreadsheet-window";
import { visibleSpreadsheetIndex, type SpreadsheetGeometry } from "./spreadsheet-geometry";

export type SpreadsheetScrollAxis = "rows" | "columns";

export function nextSpreadsheetWindow(worksheet: WorksheetSheet, geometry: SpreadsheetGeometry,
  axis: SpreadsheetScrollAxis, offset: number, viewportSize: number, direction: number,
): SpreadsheetWindow | null {
  const vertical = axis === "rows";
  const loadedStart = vertical ? worksheet.rowStart : worksheet.columnStart;
  const boundaries = vertical ? geometry.rowOffsets : geometry.columnOffsets;
  const loadedCount = boundaries.length - 1;
  const sheetCount = vertical ? worksheet.rowCount : worksheet.columnCount;
  // Fixed gutters are outside the data scroller, so prefetch and anchoring use only worksheet pixels.
  const contentSize = boundaries[boundaries.length - 1];
  const bufferPixels = Math.max(150, viewportSize / 2);
  // Fit narrow/hidden columns by pixel coverage instead of waiting for many tiny 12-column reads.
  const requiredColumnCount = Math.ceil(loadedCount * (viewportSize + bufferPixels) / Math.max(1, contentSize));
  const loadCount = vertical ? WORKSHEET_ROW_BUFFER : Math.max(WORKSHEET_WINDOW_COLUMNS, requiredColumnCount - loadedCount);
  const maximumCount = vertical ? WORKSHEET_MAX_WINDOW_ROWS : WORKSHEET_MAX_WINDOW_COLUMNS;
  let nextStart = loadedStart;
  let nextEnd = loadedStart + loadedCount;
  // Direction gates prevent an automatic correction near one edge immediately requesting the opposite edge.
  if (direction < 0 && offset < bufferPixels && loadedStart > 0) {
    nextStart = Math.max(0, loadedStart - loadCount);
    nextEnd = Math.min(nextEnd, nextStart + maximumCount);
  } else if (direction > 0 && contentSize - offset - viewportSize < bufferPixels && nextEnd < sheetCount) {
    // An underfilled horizontal window must grow in place rather than prune columns still visible at its beginning.
    const requestedEnd = !vertical && contentSize <= viewportSize
      ? loadedStart + Math.min(maximumCount, loadedCount + loadCount) : nextEnd + loadCount;
    nextEnd = Math.min(sheetCount, requestedEnd);
    nextStart = Math.max(loadedStart, nextEnd - maximumCount);
  } else return null;
  if (nextStart === loadedStart && nextEnd === loadedStart + loadedCount) return null;
  const visibleIndex = visibleSpreadsheetIndex(boundaries, offset);
  // Never prune a row/column still on screen, including an unusually wide viewport.
  if (nextStart > loadedStart + visibleIndex) return null;
  return { sheet: worksheet.name,
    rowStart: vertical ? nextStart : worksheet.rowStart,
    columnStart: vertical ? worksheet.columnStart : nextStart,
    rowLimit: vertical ? nextEnd - nextStart : geometry.rowHeights.length,
    columnLimit: vertical ? geometry.columnWidths.length : nextEnd - nextStart,
  };
}
