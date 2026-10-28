import type ExcelJS from "exceljs";
import type { WorkSheet } from "xlsx";
import type { SpreadsheetPresentation } from "../shared/spreadsheet-presentation";

export interface SpreadsheetAxis {
  defaultSize: number;
  overrides: Map<number, number>;
}
export interface SpreadsheetLayout {
  rows: SpreadsheetAxis;
  columns: SpreadsheetAxis;
  images: NonNullable<SpreadsheetPresentation["images"]>;
  imageRowCount: number;
  imageColumnCount: number;
}

export function axisPosition(axis: SpreadsheetAxis, boundary: number): number {
  let position = boundary * axis.defaultSize;
  // Sparse overrides avoid walking a million default-height rows when jumping to a distant cell.
  for (const [index, size] of axis.overrides) {
    if (index < boundary) position += size - axis.defaultSize;
  }
  return position;
}

function containingBoundary(axis: SpreadsheetAxis, position: number, maximum: number): number {
  let lowerBoundary = 0;
  let upperBoundary = maximum;
  // A binary search also handles long runs of hidden rows without dividing by a zero-sized cell.
  while (lowerBoundary < upperBoundary) {
    const middleBoundary = Math.floor((lowerBoundary + upperBoundary) / 2);
    if (axisPosition(axis, middleBoundary) < position) lowerBoundary = middleBoundary + 1;
    else upperBoundary = middleBoundary;
  }
  return lowerBoundary;
}

export function readSpreadsheetLayout(workbook: ExcelJS.Workbook, worksheet: ExcelJS.Worksheet, sourceSheet: WorkSheet): SpreadsheetLayout {
  // Excel stores row heights in points and column widths in default-font character units, not pixels.
  const rows: SpreadsheetAxis = { defaultSize: (worksheet.properties.defaultRowHeight ?? 15) * 96 / 72,
    overrides: new Map() };
  const columns: SpreadsheetAxis = { defaultSize: Math.floor((worksheet.properties.defaultColWidth ?? 8.43) * 7 + 5),
    overrides: new Map() };
  // Sparse OOXML row metadata includes hidden empty rows, which ExcelJS's value/model iterators can omit.
  sourceSheet["!rows"]?.forEach((row, rowIndex) => {
    // hpt is authoritative: SheetJS's derived hpx uses a configurable DPI that can differ from the screen's 96 DPI.
    const height = row.hpt === undefined ? row.hpx : row.hpt * 96 / 72;
    if (row.hidden || height !== undefined) rows.overrides.set(rowIndex, row.hidden ? 0 : height!);
  });
  // Column metadata is sparse in many exports; absent entries retain the sheet's default width.
  worksheet.columns?.forEach((column, columnIndex) => {
    if (column.hidden || column.width !== undefined) {
      columns.overrides.set(columnIndex, column.hidden ? 0 : Math.floor(column.width! * 7 + 5));
    }
  });
  const images: SpreadsheetLayout["images"] = [];
  for (const embeddedImage of worksheet.getImages()) {
    const image = workbook.getImage(Number(embeddedImage.imageId));
    // Never resolve a filename or a remote URL from a workbook: only actual embedded raster bytes are displayed.
    if (!image?.buffer || !["png", "jpeg", "gif"].includes(image.extension)) continue;
    const range = embeddedImage.range as ExcelJS.ImageRange & { ext?: { width: number; height: number } };
    const left = axisPosition(columns, range.tl.nativeCol) + range.tl.nativeColOff / 9525;
    const top = axisPosition(rows, range.tl.nativeRow) + range.tl.nativeRowOff / 9525;
    // Drawing offsets are EMUs (9525 per screen pixel); fractional Anchor.col uses different units internally.
    const right = range.br ? axisPosition(columns, range.br.nativeCol) + range.br.nativeColOff / 9525 : left + (range.ext?.width ?? 0);
    const bottom = range.br ? axisPosition(rows, range.br.nativeRow) + range.br.nativeRowOff / 9525 : top + (range.ext?.height ?? 0);
    if (right <= left || bottom <= top) continue;
    images.push({ uri: `data:image/${image.extension};base64,${Buffer.from(image.buffer).toString("base64")}`,
      left, top, width: right - left, height: bottom - top });
  }
  // Image-only sheets have no cell dimension; extend their bounds so the illustration can still be opened.
  return { rows, columns, images,
    imageRowCount: Math.max(0, ...images.map((image) => containingBoundary(rows, image.top + image.height, 1_048_576))),
    imageColumnCount: Math.max(0, ...images.map((image) => containingBoundary(columns, image.left + image.width, 16_384))),
  };
}
