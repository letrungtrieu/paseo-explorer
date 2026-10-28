// A jump starts with 50 rows on each side; smaller scroll chunks reduce formatting, transport and render work.
export const WORKSHEET_ROW_BUFFER = 50;
export const WORKSHEET_WINDOW_ROWS = WORKSHEET_ROW_BUFFER * 2 + 1;
export const WORKSHEET_MAX_WINDOW_ROWS = WORKSHEET_WINDOW_ROWS + WORKSHEET_ROW_BUFFER;
export const WORKSHEET_WINDOW_COLUMNS = 12;
// Hidden/narrow columns can consume many workbook coordinates before the visible viewport is filled.
export const WORKSHEET_MAX_WINDOW_COLUMNS = 256;

export interface SpreadsheetWindow {
  sheet?: string;
  rowStart: number;
  columnStart: number;
  rowLimit: number;
  columnLimit: number;
}
