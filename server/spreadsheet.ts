import { readFile, stat } from "node:fs/promises";
import { extname } from "node:path";
import * as XLSX from "xlsx";
import type { WorksheetSheet } from "../shared/viewer";
import { WORKSHEET_WINDOW_COLUMNS, WORKSHEET_WINDOW_ROWS } from "../shared/spreadsheet-window";
import { loadSpreadsheetAppearance, readSpreadsheetPresentation, type SpreadsheetAppearance } from "./spreadsheet-appearance";

interface LoadedWorkbook {
  workbook: XLSX.WorkBook;
  appearance?: Map<string, SpreadsheetAppearance>;
}

interface CachedWorkbook {
  revision: string;
  contents: LoadedWorkbook;
  cellCount: number;
  sourceBytes: number;
}

// Normal multi-workbook budgets count parsed cells; a compressed ZIP size alone cannot predict workbook memory.
const NORMAL_CACHE_CELL_BUDGET = 500_000;
const MAX_CACHED_WORKBOOKS = 3;
// Embedded photos can dominate memory even in a workbook containing very few cells.
const NORMAL_CACHE_BYTE_BUDGET = 32 * 1024 * 1024;
const workbookCache = new Map<string, CachedWorkbook>();
const pendingWorkbookReads = new Map<string, Promise<LoadedWorkbook>>();

export function clearWorkbookCache(): void {
  // Plugin teardown releases workbook graphs immediately rather than waiting for process exit.
  workbookCache.clear();
}

async function loadWorkbook(filePath: string): Promise<LoadedWorkbook> {
  const fileMetadata = await stat(filePath);
  const revision = `${fileMetadata.mtimeMs}:${fileMetadata.ctimeMs}:${fileMetadata.size}`;
  const cachedWorkbook = workbookCache.get(filePath);
  if (cachedWorkbook?.revision === revision) {
    // Refresh insertion order so the least recently viewed workbook is evicted first.
    workbookCache.delete(filePath);
    workbookCache.set(filePath, cachedWorkbook);
    return cachedWorkbook.contents;
  }
  workbookCache.delete(filePath);

  const pendingKey = `${filePath}\0${revision}`;
  const pendingRead = pendingWorkbookReads.get(pendingKey);
  if (pendingRead) return pendingRead;

  // Only the current large workbook is retained: release it before opening a different file to avoid two large graphs.
  for (const [cachedPath, cachedEntry] of workbookCache) {
    if (cachedEntry.cellCount > NORMAL_CACHE_CELL_BUDGET || cachedEntry.sourceBytes > NORMAL_CACHE_BYTE_BUDGET) workbookCache.delete(cachedPath);
  }

  // Concurrent tabs share one parse; parsing remains in Paseo's isolated plugin subprocess.
  const workbookRead = readFile(filePath).then(async (fileBuffer) => {
    const workbook = XLSX.read(fileBuffer, {
      type: "buffer", cellHTML: false, cellNF: true, cellText: false, cellFormula: false, cellStyles: true,
    });
    // Defer display formatting until a cell is requested and skip formulas in this read-only viewer.
    const cellCount = Object.values(workbook.Sheets).reduce(
      (totalCells, worksheet) => totalCells + Object.keys(worksheet).length, 0,
    );
    // ExcelJS supplies XLSX styles/drawings; SheetJS retains cached formula formatting and legacy XLS/CSV support.
    const contents: LoadedWorkbook = { workbook, appearance: extname(filePath).toLowerCase() === ".xlsx"
      ? await loadSpreadsheetAppearance(fileBuffer, workbook) : undefined };
    // A large active workbook must stay warm too; otherwise every small scroll read reparses millions of cells.
    workbookCache.set(filePath, { revision, contents, cellCount, sourceBytes: fileBuffer.byteLength });
    let cachedCellCount = [...workbookCache.values()].reduce(
      (totalCells, cachedEntry) => totalCells + cachedEntry.cellCount, 0,
    );
    let cachedSourceBytes = [...workbookCache.values()].reduce((totalBytes, cachedEntry) => totalBytes + cachedEntry.sourceBytes, 0);
    // Over budget, retain just the most recently read workbook rather than a growing collection of large files.
    while (workbookCache.size > MAX_CACHED_WORKBOOKS || (workbookCache.size > 1 &&
      (cachedCellCount > NORMAL_CACHE_CELL_BUDGET || cachedSourceBytes > NORMAL_CACHE_BYTE_BUDGET))) {
      const oldestFilePath = workbookCache.keys().next().value!;
      cachedCellCount -= workbookCache.get(oldestFilePath)!.cellCount;
      cachedSourceBytes -= workbookCache.get(oldestFilePath)!.sourceBytes;
      workbookCache.delete(oldestFilePath);
    }
    return contents;
  });
  pendingWorkbookReads.set(pendingKey, workbookRead);
  try {
    return await workbookRead;
  } finally {
    // Failed reads must be retryable after the file is repaired or replaced.
    pendingWorkbookReads.delete(pendingKey);
  }
}

interface WorksheetWindow {
  sheet?: string;
  rowStart?: number;
  columnStart?: number;
  rowLimit?: number;
  columnLimit?: number;
}

export async function readSpreadsheetWindow(
  filePath: string,
  requestedWindow: WorksheetWindow,
): Promise<WorksheetSheet[]> {
  const { workbook, appearance } = await loadWorkbook(filePath);
  const selectedSheetName = workbook.SheetNames.find(
    (sheetName) => sheetName.toLowerCase() === requestedWindow.sheet?.toLowerCase(),
  ) ?? workbook.SheetNames[0];

  return workbook.SheetNames.map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const sheetAppearance = appearance?.get(sheetName);
    const sheetRange = worksheet["!ref"] ? XLSX.utils.decode_range(worksheet["!ref"]) : null;
    // Anchor at A1 rather than the first populated cell so sparse sheets keep their real coordinates.
    // Drawings can extend beyond the last populated cell, including sheets containing only a photo.
    const rowCount = Math.max(sheetRange ? sheetRange.e.r + 1 : 0, sheetAppearance?.layout.imageRowCount ?? 0);
    const columnCount = Math.max(sheetRange ? sheetRange.e.c + 1 : 0, sheetAppearance?.layout.imageColumnCount ?? 0);
    const rowStart = Math.min(requestedWindow.rowStart ?? 0, Math.max(0, rowCount - 1));
    const columnStart = Math.min(requestedWindow.columnStart ?? 0, Math.max(0, columnCount - 1));
    const rows: string[][] = [];

    // Never walk !ref in full: whole-column formatting can otherwise allocate millions of blanks.
    if (sheetName === selectedSheetName) {
      // Requested limits permit overlapping scroll windows without exposing unbounded RPC allocations.
      const rowEnd = Math.min(rowCount, rowStart + (requestedWindow.rowLimit ?? WORKSHEET_WINDOW_ROWS));
      const columnEnd = Math.min(columnCount, columnStart + (requestedWindow.columnLimit ?? WORKSHEET_WINDOW_COLUMNS));
      for (let rowIndex = rowStart; rowIndex < rowEnd; rowIndex++) {
        const displayedRow: string[] = [];
        for (let columnIndex = columnStart; columnIndex < columnEnd; columnIndex++) {
          const cell = worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
          displayedRow.push(cell ? XLSX.utils.format_cell(cell) : "");
        }
        rows.push(displayedRow);
      }
    }
    return { name: sheetName, rowCount, columnCount, rowStart, columnStart, rows,
      ...(sheetAppearance && sheetName === selectedSheetName ? {
        presentation: readSpreadsheetPresentation(sheetAppearance, worksheet, {
          rowStart, columnStart, rowLength: rows.length, columnLength: rows[0]?.length ?? 0,
        }),
      } : {}),
    };
  });
}
