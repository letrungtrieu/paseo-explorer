import type { WorksheetSheet } from "../shared/viewer";

// Older/CSV responses have no saved geometry; keep their existing default sizes for compatibility.
export const COL_WIDTH = 115;
export const ROW_HEIGHT = 32;
// Heading dimensions belong to fixed gutters rather than the scrollable worksheet coordinates.
export const COLUMN_HEADING_HEIGHT = 26;
export const ROW_HEADING_WIDTH = 40;

function accumulatedSizes(sizes: number[]): number[] {
  const boundaries = [0];
  for (const size of sizes) boundaries.push(boundaries[boundaries.length - 1] + size);
  return boundaries;
}

export function createSpreadsheetGeometry(worksheet: WorksheetSheet) {
  const columnWidths = worksheet.presentation?.columnWidths ??
    Array.from({ length: worksheet.rows[0]?.length ?? 0 }, () => COL_WIDTH);
  const rowHeights = worksheet.presentation?.rowHeights ?? worksheet.rows.map(() => ROW_HEIGHT);
  // Prefix positions keep cell drawing, image anchors, and scroll-to-selection on the same coordinate system.
  const columnOffsets = accumulatedSizes(columnWidths);
  const rowOffsets = accumulatedSizes(rowHeights);
  return { columnWidths, rowHeights, columnOffsets, rowOffsets,
    width: columnOffsets[columnOffsets.length - 1], height: rowOffsets[rowOffsets.length - 1] };
}
export type SpreadsheetGeometry = ReturnType<typeof createSpreadsheetGeometry>;

export function visibleSpreadsheetIndex(boundaries: number[], position: number): number {
  let lowerIndex = 0;
  let upperIndex = boundaries.length - 1;
  // Equal boundaries belong to hidden rows/columns; choose the final boundary at that position.
  while (lowerIndex < upperIndex) {
    const middleIndex = Math.ceil((lowerIndex + upperIndex) / 2);
    if (boundaries[middleIndex] <= position) lowerIndex = middleIndex;
    else upperIndex = middleIndex - 1;
  }
  return Math.min(lowerIndex, Math.max(0, boundaries.length - 2));
}
