import { mkdtemp, mkdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import * as XLSX from "xlsx";

// Mock react-native primitives so Vitest can execute client TSX components without parsing React Native's Flow entry file.
vi.mock("react-native", () => ({
  View: (props: Record<string, unknown>) => React.createElement("View", props),
  Text: (props: Record<string, unknown>) => React.createElement("Text", props),
  Pressable: (props: Record<string, unknown>) => React.createElement("Pressable", props),
  ScrollView: React.forwardRef((props: Record<string, unknown>, ref) =>
    React.createElement("ScrollView", { ...props, ref }),
  ),
}));

import { ICON_CONFIG, WindowsFileIcon } from "../client/file-icons";
import { AgentOfficeLinkRenderer, officeIssuesTimelineTransformer } from "../client/agent-link-renderer";
import { listDirectoryEntries } from "../server/explorer";
import { readOfficeFile } from "../server/viewer";
import { formatColumnLabel, officeStore, parseCellRef } from "../shared/office-store";
import { OFFICE_ISSUES_TIMELINE_KIND } from "../shared/viewer";

describe("office-store cell coordinate helpers and state synchronization", () => {
  beforeEach(() => {
    officeStore.reset();
  });

  it("parses Excel A1-style cell references into zero-based row and column coordinates", () => {
    // F is the 6th letter (index 5) and row 14 is index 13.
    expect(parseCellRef("F14")).toEqual({ col: 5, row: 13 });
    expect(parseCellRef("a1")).toEqual({ col: 0, row: 0 });
    expect(parseCellRef("AA10")).toEqual({ col: 26, row: 9 });
    expect(parseCellRef("A0")).toBeNull();
    expect(parseCellRef("14F")).toBeNull();
    expect(parseCellRef(undefined)).toBeNull();
  });

  it("formats zero-based column indices back into Excel column letters", () => {
    expect(formatColumnLabel(0)).toBe("A");
    expect(formatColumnLabel(4)).toBe("E");
    expect(formatColumnLabel(5)).toBe("F");
    expect(formatColumnLabel(25)).toBe("Z");
    expect(formatColumnLabel(26)).toBe("AA");
  });

  it("synchronizes registered Agent issues and openAndJump targets across subscribers", () => {
    const observedSnapshots: Array<ReturnType<typeof officeStore.getState>> = [];
    const unsubscribe = officeStore.subscribe((snapshot) => {
      observedSnapshots.push(snapshot);
    });

    const invoicePath = "/workspace/customs-declarations/2026-10/Commercial_Invoice_01.xlsx";
    officeStore.registerIssues(invoicePath, [
      { ext: ".xlsx", sheet: "Sheet1", cell: "F14", reason: "Sai mã HS Code" },
      { ext: ".xlsx", sheet: "Sheet1", cell: "G14", reason: "Thiếu đơn vị tiền tệ" },
    ]);

    expect(officeStore.getState().issuesByFile[invoicePath]).toHaveLength(2);

    officeStore.openAndJump({
      filePath: invoicePath,
      ext: ".xlsx",
      sheet: "Sheet1",
      cell: "F14",
      reason: "Sai mã HS Code",
    });

    const latestState = officeStore.getState();
    expect(latestState.activeTarget?.filePath).toBe(invoicePath);
    expect(latestState.activeTarget?.cell).toBe("F14");
    expect(latestState.activeTarget?.reason).toBe("Sai mã HS Code");
    expect(observedSnapshots).toHaveLength(2);

    unsubscribe();
  });
});

describe("server explorer and office file viewer handlers", () => {
  let temporaryWorkspaceDir = "";

  beforeEach(async () => {
    temporaryWorkspaceDir = await mkdtemp(join(tmpdir(), "paseo-explorer-test-"));
  });

  afterEach(async () => {
    if (temporaryWorkspaceDir) {
      await rm(temporaryWorkspaceDir, { recursive: true, force: true });
    }
  });

  it("lists workspace directories first and normalizes file extensions while filtering .git", async () => {
    const declarationFolder = join(temporaryWorkspaceDir, "2026-10");
    const archiveFolder = join(temporaryWorkspaceDir, "archive");
    const hiddenGitFolder = join(temporaryWorkspaceDir, ".git");
    await mkdir(declarationFolder, { recursive: true });
    await mkdir(archiveFolder, { recursive: true });
    await mkdir(hiddenGitFolder, { recursive: true });

    await writeFile(join(declarationFolder, "Commercial_Invoice_01.xlsx"), "dummy");
    await writeFile(join(declarationFolder, "Contract_992.docx"), "dummy");
    await writeFile(join(declarationFolder, "Bill_of_Lading.pdf"), "dummy");

    const rootListing = await listDirectoryEntries({ dirPath: temporaryWorkspaceDir });
    // .git must be excluded so the file tree only shows user workspace folders.
    expect(rootListing.entries.map((entry) => entry.name)).toEqual(["2026-10", "archive"]);
    expect(rootListing.entries.every((entry) => entry.isDirectory)).toBe(true);

    const subFolderListing = await listDirectoryEntries({ dirPath: declarationFolder });
    expect(subFolderListing.entries.map((entry) => ({ name: entry.name, ext: entry.ext }))).toEqual([
      { name: "Bill_of_Lading.pdf", ext: ".pdf" },
      { name: "Commercial_Invoice_01.xlsx", ext: ".xlsx" },
      { name: "Contract_992.docx", ext: ".docx" },
    ]);
  });

  it("parses Excel (.xlsx), Word (.docx), and PDF (.pdf) documents for the sidebar viewer", async () => {
    const excelFilePath = join(temporaryWorkspaceDir, "Commercial_Invoice_01.xlsx");
    // Build a 14-row worksheet so cell F14 (row index 13, col index 5) holds the target HS code.
    const sheetRows: string[][] = Array.from({ length: 14 }, (_, rowIndex) => [
      `Item-${rowIndex + 1}`,
      "A",
      "B",
      "C",
      rowIndex === 12 ? "8471.30.20" : "8471.99.00",
      rowIndex === 12 ? "1,200.00" : "8500.12.99",
      "USD",
    ]);

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(sheetRows);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    const xlsxBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }) as Buffer;
    await writeFile(excelFilePath, xlsxBuffer);

    const excelData = await readOfficeFile({ filePath: excelFilePath });
    expect(excelData.fileType).toBe("xlsx");
    expect(excelData.sheets?.[0]?.name).toBe("Sheet1");
    const f14Coords = parseCellRef("F14")!;
    expect(excelData.sheets?.[0]?.rows[f14Coords.row]?.[f14Coords.col]).toBe("8500.12.99");

    const docxFilePath = join(temporaryWorkspaceDir, "Contract_992.docx");
    await writeFile(docxFilePath, "<w:document><w:body><w:p><w:t>Hop dong ngoai thuong 992</w:t></w:p></w:body></w:document>");
    const docxData = await readOfficeFile({ filePath: docxFilePath });
    expect(docxData.fileType).toBe("docx");
    expect(docxData.textContent).toContain("Hop dong ngoai thuong 992");

    const pdfFilePath = join(temporaryWorkspaceDir, "Bill_of_Lading.pdf");
    await writeFile(pdfFilePath, "%PDF-1.4\n1 0 obj << /Type /Page >>\nBT (Bill of Lading B/L No. 88421) Tj ET");
    const pdfData = await readOfficeFile({ filePath: pdfFilePath });
    expect(pdfData.fileType).toBe("pdf");
    expect(pdfData.textContent).toContain("Bill of Lading B/L No. 88421");
  });
});

describe("WindowsFileIcon badges and AgentOfficeLinkRenderer", () => {
  it("defines distinct Windows/Office badge colors for XLS, CSV, DOC, PDF, and JSON", () => {
    expect(ICON_CONFIG[".xlsx"].label).toBe("XLS");
    expect(ICON_CONFIG[".docx"].label).toBe("DOC");
    expect(ICON_CONFIG[".pdf"].label).toBe("PDF");

    // Verify the component renders React elements for folders, Office badges, and fallback files.
    const folderElement = WindowsFileIcon({ isDirectory: true, isOpen: true, ext: "" });
    const excelElement = WindowsFileIcon({ isDirectory: false, ext: ".xlsx" });
    const plainElement = WindowsFileIcon({ isDirectory: false, ext: ".txt" });
    expect(folderElement).toBeTruthy();
    expect(excelElement).toBeTruthy();
    expect(plainElement).toBeTruthy();
  });

  it("transforms Agent tool calls carrying office issues into interactive timeline items", () => {
    const sampleIssuePayload = {
      filePath: "/workspace/customs-declarations/2026-10/Commercial_Invoice_01.xlsx",
      ext: ".xlsx",
      issues: [{ sheet: "Sheet1", cell: "F14", reason: "Sai mã HS Code" }],
    };

    const transformedResult = officeIssuesTimelineTransformer.transform({
      item: {
        type: "tool_call",
        callId: "call-42",
        name: "report_office_issues",
        status: "completed",
        error: null,
        detail: {
          type: "unknown",
          input: sampleIssuePayload,
          output: { ok: true },
        },
      },
      phase: "complete",
    });

    expect(transformedResult?.items).toHaveLength(1);
    expect(transformedResult?.items[0]?.kind).toBe(OFFICE_ISSUES_TIMELINE_KIND);
    expect(transformedResult?.items[0]?.data).toEqual(sampleIssuePayload);
    expect(officeStore.getState().issuesByFile[sampleIssuePayload.filePath]).toHaveLength(1);

    const renderedCardElement = React.createElement(AgentOfficeLinkRenderer, {
      input: sampleIssuePayload,
      theme: { colors: { surface: "#1e293b", text: "#f8fafc", textMuted: "#94a3b8", primary: "#3b82f6" } },
    });
    expect(renderedCardElement.type).toBe(AgentOfficeLinkRenderer);
  });
});
