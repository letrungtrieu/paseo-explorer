import { defineRpc } from "@getpaseo/plugin";
import { z } from "zod";
import { SpreadsheetPresentationSchema } from "./spreadsheet-presentation";
import { WORKSHEET_MAX_WINDOW_COLUMNS, WORKSHEET_MAX_WINDOW_ROWS } from "./spreadsheet-window";

// Timeline item kind and version registered with Paseo so agent issue reports render as interactive cards.
export const OFFICE_ISSUES_TIMELINE_KIND = "office-explorer-issues";
export const OFFICE_ISSUES_TIMELINE_VERSION = 1;

export const WorksheetSheetSchema = z.object({
  name: z.string(),
  rowCount: z.number().int().nonnegative(),
  columnCount: z.number().int().nonnegative(),
  // Offsets preserve Excel coordinates when only a small window of a sheet is transported.
  rowStart: z.number().int().nonnegative(),
  columnStart: z.number().int().nonnegative(),
  rows: z.array(z.array(z.string())),
  // CSV and legacy readers can omit presentation while XLSX carries its original visual layout.
  presentation: SpreadsheetPresentationSchema.optional(),
});

export type WorksheetSheet = z.infer<typeof WorksheetSheetSchema>;

export const OfficeFileDataSchema = z.object({
  filePath: z.string(),
  // All spreadsheet variants (.xlsx, .xls, .csv) normalize to "xlsx" so the grid viewer handles them uniformly.
  fileType: z.enum(["xlsx", "docx", "pdf"]),
  sheets: z.array(WorksheetSheetSchema).optional(),
  textContent: z.string().optional(),
  // Word HTML is shown inside Paseo's sandbox; PDF bytes retain graphics and scanned pages.
  htmlContent: z.string().optional(),
  documentBase64: z.string().optional(),
  // Optional per-page text breakdown for PDFs and multi-section documents so page jumps can highlight target pages.
  pages: z.array(z.string()).optional(),
});

export type OfficeFileData = z.infer<typeof OfficeFileDataSchema>;

// RPC contract for parsing an Office document on the server and returning structured content to the viewer.
// Lowercase snake_case is required by @getpaseo/plugin's RPC_NAME validation regex (/^[a-z][a-z0-9._-]*$/).
export const readOfficeFileRpc = defineRpc({
  name: "office.viewer.read_file",
  input: z.object({
    filePath: z.string().min(1),
    // Optional workspace directory used to resolve relative paths like "examples/Ngân sách hàng tháng.xlsx".
    cwd: z.string().optional(),
    // Other sheets return metadata only, so switching tabs never transfers an entire workbook.
    sheet: z.string().optional(),
    rowStart: z.number().int().nonnegative().optional(),
    columnStart: z.number().int().nonnegative().optional(),
    // Scroll prefetch can grow the initial window, but never request an entire large sheet in one call.
    rowLimit: z.number().int().positive().max(WORKSHEET_MAX_WINDOW_ROWS).optional(),
    columnLimit: z.number().int().positive().max(WORKSHEET_MAX_WINDOW_COLUMNS).optional(),
  }),
  output: OfficeFileDataSchema,
});

export const OfficeIssueItemSchema = z.object({
  sheet: z.string().optional(),
  cell: z.string().optional(),
  page: z.number().int().positive().optional(),
  reason: z.string(),
});

export type OfficeIssueItem = z.infer<typeof OfficeIssueItemSchema>;

export const OfficeIssueReportPayloadSchema = z.object({
  filePath: z.string().min(1),
  ext: z.string(),
  issues: z.array(OfficeIssueItemSchema),
  // Preserves the chat message body when transforming assistant_message or user_message items.
  originalMessageText: z.string().optional(),
});

export type OfficeIssueReportPayload = z.infer<typeof OfficeIssueReportPayloadSchema>;
