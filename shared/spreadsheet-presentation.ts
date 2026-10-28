import { z } from "zod";

// Transport visual properties explicitly; browser/native views must not depend on a server-side workbook object.
const CellBorderSchema = z.object({
  color: z.string(), width: z.number(), style: z.enum(["solid", "dashed", "dotted", "double"]),
});
export const SpreadsheetCellStyleSchema = z.object({
  fontFamily: z.string().optional(), fontSize: z.number().optional(),
  bold: z.boolean().optional(), italic: z.boolean().optional(),
  underline: z.boolean().optional(), strike: z.boolean().optional(),
  color: z.string().optional(), background: z.string().optional(),
  horizontal: z.enum(["left", "center", "right", "justify"]).optional(),
  vertical: z.enum(["top", "center", "bottom"]).optional(),
  wrap: z.boolean().optional(), indent: z.number().optional(),
  // Each edge is independent because invoices commonly use only a bottom rule or an outside border.
  top: CellBorderSchema.optional(), right: CellBorderSchema.optional(),
  bottom: CellBorderSchema.optional(), left: CellBorderSchema.optional(),
});
export type SpreadsheetCellStyle = z.infer<typeof SpreadsheetCellStyleSchema>;

export const SpreadsheetPresentationSchema = z.object({
  columnWidths: z.array(z.number().nonnegative()),
  rowHeights: z.array(z.number().nonnegative()),
  // Shared style IDs keep repeated formatting from multiplying the bounded preview payload.
  styles: z.array(SpreadsheetCellStyleSchema).optional(),
  cellStyleIds: z.array(z.array(z.number().int().nonnegative())).optional(),
  merges: z.array(z.object({
    startRow: z.number(), endRow: z.number(), startColumn: z.number(), endColumn: z.number(),
    text: z.string(), styleId: z.number(),
  })).optional(),
  // Coordinates are relative to the requested window, allowing images crossing a page boundary to be clipped.
  images: z.array(z.object({
    uri: z.string(), left: z.number(), top: z.number(), width: z.number(), height: z.number(),
  })).optional(),
  showGridLines: z.boolean().optional(),
});
export type SpreadsheetPresentation = z.infer<typeof SpreadsheetPresentationSchema>;
