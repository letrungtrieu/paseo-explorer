import { readFile } from "node:fs/promises";
import { extname, isAbsolute, resolve } from "node:path";
import type { RpcInput, RpcOutput } from "@getpaseo/plugin";
import mammoth from "mammoth";
import { readOfficeFileRpc } from "../shared/viewer";
import { readSpreadsheetWindow } from "./spreadsheet";

export async function readOfficeFile(
  input: RpcInput<typeof readOfficeFileRpc>,
): Promise<RpcOutput<typeof readOfficeFileRpc>> {
  // Resolve chat links on the selected daemon, including workspaces accessed from a remote client.
  const targetFilePath = isAbsolute(input.filePath)
    ? input.filePath
    : resolve(input.cwd ?? process.cwd(), input.filePath);
  const fileExtension = extname(targetFilePath).toLowerCase();

  if ([".xlsx", ".xls", ".csv"].includes(fileExtension)) {
    return {
      filePath: targetFilePath,
      fileType: "xlsx",
      sheets: await readSpreadsheetWindow(targetFilePath, input),
    };
  }

  // Legacy binary Word needs a different decoder; interpreting its bytes as UTF-8 produces garbage.
  if (fileExtension === ".doc") {
    throw new Error("File .doc cũ chưa được hỗ trợ. Hãy lưu thành .docx hoặc PDF để xem.");
  }
  if (fileExtension !== ".docx" && fileExtension !== ".pdf") {
    throw new Error(`Chưa hỗ trợ xem định dạng ${fileExtension || "không có phần mở rộng"}.`);
  }

  const fileBuffer = await readFile(targetFilePath);
  if (fileExtension === ".pdf") {
    // Preserve the actual PDF: text extraction loses charts, fonts, scanned pages, and page numbering.
    if (!fileBuffer.subarray(0, 1024).includes(Buffer.from("%PDF-"))) {
      throw new Error("File PDF không hợp lệ hoặc bị hỏng.");
    }
    return { filePath: targetFilePath, fileType: "pdf", documentBase64: fileBuffer.toString("base64") };
  }

  // Mammoth preserves headings, tables, and embedded images; raw text remains available on native clients.
  const [wordHtml, wordText] = await Promise.all([
    mammoth.convertToHtml({ buffer: fileBuffer }, { externalFileAccess: false }),
    mammoth.extractRawText({ buffer: fileBuffer }),
  ]);
  return {
    filePath: targetFilePath,
    fileType: "docx",
    htmlContent: wordHtml.value,
    textContent: wordText.value.trim(),
  };
}
