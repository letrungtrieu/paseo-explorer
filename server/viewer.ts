import { execFile } from "node:child_process";
import { readFile } from "node:fs/promises";
import { extname } from "node:path";
import { promisify } from "node:util";
import type { RpcInput, RpcOutput } from "@getpaseo/plugin";
import mammoth from "mammoth";
import * as XLSX from "xlsx";
import { readOfficeFileRpc, type WorksheetSheet } from "../shared/viewer";

const execFileAsync = promisify(execFile);

// Pad ragged spreadsheet rows to a uniform width so every row in the UI grid renders complete cell borders.
function normalizeWorksheetMatrix(rawSheetMatrix: unknown[][]): string[][] {
  let maxColumnCount = 0;
  for (const sheetRow of rawSheetMatrix) {
    if (Array.isArray(sheetRow) && sheetRow.length > maxColumnCount) {
      maxColumnCount = sheetRow.length;
    }
  }
  if (maxColumnCount === 0) {
    return [];
  }

  return rawSheetMatrix.map((sheetRow) => {
    const normalizedRow: string[] = [];
    for (let colIdx = 0; colIdx < maxColumnCount; colIdx++) {
      const rawCellValue = Array.isArray(sheetRow) ? sheetRow[colIdx] : "";
      normalizedRow.push(rawCellValue === null || rawCellValue === undefined ? "" : String(rawCellValue));
    }
    return normalizedRow;
  });
}

// Parse .xlsx, .xls, and .csv files into uniform 2D string grids per worksheet.
function parseSpreadsheetBuffer(fileBuffer: Buffer): WorksheetSheet[] {
  const parsedWorkbook = XLSX.read(fileBuffer, { type: "buffer" });
  return parsedWorkbook.SheetNames.map((sheetName) => {
    const worksheetObject = parsedWorkbook.Sheets[sheetName];
    // raw: false preserves formatted numbers/dates as displayed in Excel (e.g., "1,200.00").
    const rawRows = XLSX.utils.sheet_to_json<unknown[]>(worksheetObject, {
      header: 1,
      defval: "",
      raw: false,
    });
    return {
      name: sheetName,
      rows: normalizeWorksheetMatrix(rawRows),
    };
  });
}

// Extract readable text from Word documents via mammoth, with a fallback for legacy .doc or XML text fixtures.
async function parseWordDocumentBuffer(fileBuffer: Buffer): Promise<string> {
  try {
    const mammothExtraction = await mammoth.extractRawText({ buffer: fileBuffer });
    if (mammothExtraction.value.trim().length > 0) {
      return mammothExtraction.value.trim();
    }
  } catch {
    // Fall through to XML/UTF-8 extraction when the file is a legacy .doc or non-zipped test fixture.
  }

  const rawUtf8Text = fileBuffer.toString("utf8");
  // Strip XML tags if the buffer contains raw WordprocessingML markup.
  const strippedXmlText = rawUtf8Text.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
  return strippedXmlText;
}

// Pure-JS fallback for extracting literal strings from uncompressed PDF text streams when pdftotext is absent.
function extractPdfLiteralsFallback(fileBuffer: Buffer): string[] {
  const latin1Content = fileBuffer.toString("latin1");
  const rawPageSegments = latin1Content.split(/\/Type\s*\/Page\b/);
  const candidateSegments = rawPageSegments.length > 1 ? rawPageSegments.slice(1) : [latin1Content];

  const extractedPages = candidateSegments
    .map((pageSegment) => {
      const literalMatches = [...pageSegment.matchAll(/\(([^()\\]*(?:\\.[^()\\]*)*)\)/g)];
      return literalMatches
        .map((matchGroup) =>
          matchGroup[1]
            .replace(/\\n/g, "\n")
            .replace(/\\r/g, "\r")
            .replace(/\\t/g, "\t")
            .replace(/\\\(/g, "(")
            .replace(/\\\)/g, ")"),
        )
        .join(" ")
        .trim();
    })
    .filter((pageText) => pageText.length > 0);

  if (extractedPages.length > 0) {
    return extractedPages;
  }
  return [fileBuffer.toString("utf8").trim()];
}

// Extract PDF text preserving layout via pdftotext when available, falling back to literal stream parsing.
async function parsePdfFile(filePath: string, fileBuffer: Buffer): Promise<{ textContent: string; pages: string[] }> {
  try {
    const { stdout: extractedStdout } = await execFileAsync("pdftotext", ["-layout", filePath, "-"]);
    const splitPages = extractedStdout
      .split("\f")
      .map((pageBody) => pageBody.trim())
      .filter((pageBody) => pageBody.length > 0);

    if (splitPages.length > 0) {
      return {
        textContent: splitPages.join("\n\n--- Trang tiếp theo ---\n\n"),
        pages: splitPages,
      };
    }
  } catch {
    // Fallback to pure-JS literal extraction when pdftotext fails or the file is a lightweight test buffer.
  }

  const fallbackPages = extractPdfLiteralsFallback(fileBuffer);
  return {
    textContent: fallbackPages.join("\n\n"),
    pages: fallbackPages,
  };
}

export async function readOfficeFile(
  input: RpcInput<typeof readOfficeFileRpc>,
): Promise<RpcOutput<typeof readOfficeFileRpc>> {
  const targetFilePath = input.filePath;
  const fileExtension = extname(targetFilePath).toLowerCase();
  const fileBuffer = await readFile(targetFilePath);

  if (fileExtension === ".xlsx" || fileExtension === ".xls" || fileExtension === ".csv") {
    const parsedSheets = parseSpreadsheetBuffer(fileBuffer);
    return {
      filePath: targetFilePath,
      fileType: "xlsx",
      sheets: parsedSheets,
    };
  }

  if (fileExtension === ".docx" || fileExtension === ".doc") {
    const extractedWordText = await parseWordDocumentBuffer(fileBuffer);
    return {
      filePath: targetFilePath,
      fileType: "docx",
      textContent: extractedWordText,
      pages: [extractedWordText],
    };
  }

  if (fileExtension === ".pdf") {
    const extractedPdf = await parsePdfFile(targetFilePath, fileBuffer);
    return {
      filePath: targetFilePath,
      fileType: "pdf",
      textContent: extractedPdf.textContent,
      pages: extractedPdf.pages,
    };
  }

  // Fallback for unexpected extensions so the viewer still displays UTF-8 text instead of throwing.
  const fallbackTextContent = fileBuffer.toString("utf8");
  return {
    filePath: targetFilePath,
    fileType: "docx",
    textContent: fallbackTextContent,
    pages: [fallbackTextContent],
  };
}
