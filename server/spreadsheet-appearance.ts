import ExcelJS from "exceljs";
import * as XLSX from "xlsx";
import type { SpreadsheetCellStyle, SpreadsheetPresentation } from "../shared/spreadsheet-presentation";
import { readSpreadsheetTheme, resolveSpreadsheetColor } from "./spreadsheet-colors";
import { axisPosition, readSpreadsheetLayout, type SpreadsheetLayout } from "./spreadsheet-layout";
import { normalizeSpreadsheetPackage } from "./spreadsheet-package";

export interface SpreadsheetAppearance {
  worksheet: ExcelJS.Worksheet;
  layout: SpreadsheetLayout;
  theme: string[];
}

export async function loadSpreadsheetAppearance(fileBuffer: Buffer, sourceWorkbook: XLSX.WorkBook): Promise<Map<string, SpreadsheetAppearance>> {
  const workbook = new ExcelJS.Workbook();
  const compatibleBuffer = await normalizeSpreadsheetPackage(fileBuffer);
  // SheetJS already owns merge ranges. Skipping ExcelJS's expansion prevents huge merges allocating millions of cells.
  await workbook.xlsx.load(compatibleBuffer as unknown as Parameters<typeof workbook.xlsx.load>[0], { ignoreNodes: ["mergeCells"] });
  const theme = readSpreadsheetTheme(Object.values(workbook.model.themes ?? {})[0]);
  return new Map(workbook.worksheets.map((worksheet) => [worksheet.name, {
    worksheet, layout: readSpreadsheetLayout(workbook, worksheet, sourceWorkbook.Sheets[worksheet.name]), theme,
  }]));
}

function cellBorder(border: Partial<ExcelJS.Border> | undefined, theme: string[]): SpreadsheetCellStyle["top"] {
  if (!border?.style) return undefined;
  // Excel's many dash variants collapse to the native renderer's dash/dot primitives while keeping their weight.
  return { color: resolveSpreadsheetColor(border.color, theme) ?? "#000000",
    width: border.style === "thick" ? 3 : border.style.startsWith("medium") || border.style === "double" ? 2 : 1,
    style: border.style === "double" ? "double" : border.style === "dotted" || border.style === "hair" ? "dotted"
      : /dash/i.test(border.style) ? "dashed" : "solid" };
}

function readCellStyle(appearance: SpreadsheetAppearance, rowIndex: number, columnIndex: number): SpreadsheetCellStyle {
  // ExcelJS exposes row.style at runtime but omits that aggregate property from its Row declaration.
  const row = appearance.worksheet.findRow(rowIndex + 1) as (ExcelJS.Row & { style: Partial<ExcelJS.Style> }) | undefined;
  const column = appearance.worksheet.columns?.[columnIndex];
  const cell = row?.findCell(columnIndex + 1);
  // findCell avoids allocating blank workbook cells every time the user browses another preview window.
  const style = { ...column?.style, ...row?.style, ...cell?.style };
  const font = style.font;
  const alignment = style.alignment;
  const fill = style.fill;
  const cellValue = cell?.value;
  const numericResult = typeof cellValue === "number" || cellValue instanceof Date ||
    (typeof cellValue === "object" && cellValue !== null && "result" in cellValue && typeof cellValue.result === "number");
  const horizontal = alignment?.horizontal;
  // Numeric cells keep Excel's default right alignment even when no explicit alignment style was saved.
  return {
    fontFamily: font?.name, fontSize: font?.size === undefined ? undefined : font.size * 96 / 72,
    bold: font?.bold, italic: font?.italic, underline: font?.underline ? font.underline !== "none" : undefined,
    strike: font?.strike, color: resolveSpreadsheetColor(font?.color, appearance.theme),
    background: fill?.type === "pattern" && fill.pattern === "solid"
      ? resolveSpreadsheetColor(fill.fgColor, appearance.theme) : undefined,
    horizontal: horizontal === "centerContinuous" || horizontal === "distributed" ? "center"
      : horizontal === "fill" ? "left" : horizontal ?? (numericResult ? "right" : undefined),
    vertical: alignment?.vertical === "middle" ? "center"
      : alignment?.vertical === "distributed" || alignment?.vertical === "justify" ? "center" : alignment?.vertical,
    wrap: alignment?.wrapText, indent: alignment?.indent,
    top: cellBorder(style.border?.top, appearance.theme), right: cellBorder(style.border?.right, appearance.theme),
    bottom: cellBorder(style.border?.bottom, appearance.theme), left: cellBorder(style.border?.left, appearance.theme),
  };
}

export function readSpreadsheetPresentation(
  appearance: SpreadsheetAppearance, worksheet: XLSX.WorkSheet,
  window: { rowStart: number; columnStart: number; rowLength: number; columnLength: number },
): SpreadsheetPresentation {
  const { rowStart, columnStart, rowLength, columnLength } = window;
  const { rows, columns, images } = appearance.layout;
  const columnWidths = Array.from({ length: columnLength }, (_, visibleColumn) =>
    columns.overrides.get(columnStart + visibleColumn) ?? columns.defaultSize);
  const rowHeights = Array.from({ length: rowLength }, (_, visibleRow) =>
    rows.overrides.get(rowStart + visibleRow) ?? rows.defaultSize);
  const styles: SpreadsheetCellStyle[] = [{}];
  const styleIds = new Map<string, number>([["{}", 0]]);
  const cellStyleId = (rowIndex: number, columnIndex: number) => {
    const style = readCellStyle(appearance, rowIndex, columnIndex);
    const styleKey = JSON.stringify(style);
    const existingStyleId = styleIds.get(styleKey);
    if (existingStyleId !== undefined) return existingStyleId;
    // Intern styles only in this window; a giant workbook's unused formatting never crosses the RPC boundary.
    const assignedStyleId = styles.length;
    styles.push(style);
    styleIds.set(styleKey, assignedStyleId);
    return assignedStyleId;
  };
  const cellStyleIds = Array.from({ length: rowLength }, (_, visibleRow) =>
    // Hidden cells never paint; merged anchors below still explicitly resolve their saved styles when needed.
    Array.from({ length: columnLength }, (_, visibleColumn) => rowHeights[visibleRow] === 0 || columnWidths[visibleColumn] === 0
      ? 0 : cellStyleId(rowStart + visibleRow, columnStart + visibleColumn)));
  const merges = (worksheet["!merges"] ?? []).filter((merge) => merge.e.r >= rowStart && merge.s.r < rowStart + rowLength &&
    merge.e.c >= columnStart && merge.s.c < columnStart + columnLength).map((merge) => {
    const anchor = worksheet[XLSX.utils.encode_cell(merge.s)];
    // Keep the real anchor text/style when a merged region continues into a later row or column window.
    return { startRow: merge.s.r, endRow: merge.e.r, startColumn: merge.s.c, endColumn: merge.e.c,
      text: anchor ? XLSX.utils.format_cell(anchor) : "", styleId: cellStyleId(merge.s.r, merge.s.c) };
  });
  const windowLeft = axisPosition(columns, columnStart);
  const windowTop = axisPosition(rows, rowStart);
  const windowRight = axisPosition(columns, columnStart + columnLength);
  const windowBottom = axisPosition(rows, rowStart + rowLength);
  // Send overlapping drawings, including images anchored outside the visible cells, and let the canvas clip them.
  const visibleImages = images.filter((image) => image.left < windowRight && image.left + image.width > windowLeft &&
    image.top < windowBottom && image.top + image.height > windowTop).map((image) => ({ ...image,
      left: image.left - windowLeft, top: image.top - windowTop }));
  return {
    columnWidths, rowHeights,
    ...(styles.length > 1 ? { styles, cellStyleIds } : {}),
    ...(merges.length ? { merges } : {}), ...(visibleImages.length ? { images: visibleImages } : {}),
    // Some valid exports omit sheetViews entirely; Excel shows gridlines by default in that case.
    showGridLines: appearance.worksheet.views?.[0]?.showGridLines !== false,
  };
}
