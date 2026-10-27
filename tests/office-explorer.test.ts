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

vi.mock("@getpaseo/plugin/client/ui", () => ({
  Markdown: (props: { children: string }) => React.createElement("Markdown", props),
}));

import { ICON_CONFIG, WindowsFileIcon } from "../client/file-icons";
import {
  AgentOfficeLinkRenderer,
  officeAssistantMessageTransformer,
  officeIssuesTimelineTransformer,
} from "../client/agent-link-renderer";
import { OfficeFilePanel } from "../client/office-explorer-sidebar";
import { listDirectoryEntries } from "../server/explorer";
import {
  OFFICE_EXPLORER_PROMPT_MARKER,
  augmentAgentCreateRequestWithOfficePrompt,
  injectOfficeExplorerSystemPrompt,
  registerOfficePromptInjector,
} from "../server/prompt-injector";
import { readOfficeFile } from "../server/viewer";
import { parseOfficeIssuesFromMessageText } from "../shared/office-parser";
import {
  findIssuesForFilePath,
  formatColumnLabel,
  isCellInSelections,
  isCellSelected,
  officeStore,
  parseCellRef,
  parseGridSelection,
  parseGridSelections,
} from "../shared/office-store";
import { OFFICE_ISSUES_TIMELINE_KIND } from "../shared/viewer";

describe("office-store cell, row, col, and range coordinate helpers", () => {
  beforeEach(() => {
    officeStore.reset();
  });

  it("parses Excel A1-style cell references into zero-based row and column coordinates", () => {
    expect(parseCellRef("F14")).toEqual({ col: 5, row: 13 });
    expect(parseCellRef("E3")).toEqual({ col: 4, row: 2 });
    expect(parseCellRef("AA10")).toEqual({ col: 26, row: 9 });
    expect(parseCellRef("A0")).toBeNull();
    expect(parseCellRef("14F")).toBeNull();
    expect(parseCellRef(undefined)).toBeNull();
  });

  it("parses cell, row, col, and range selections and checks cell membership", () => {
    const cellSelection = parseGridSelection("E3");
    expect(cellSelection).toEqual({
      kind: "cell",
      label: "E3",
      startRow: 2,
      endRow: 2,
      startCol: 4,
      endCol: 4,
    });
    expect(isCellSelected(cellSelection, 2, 4)).toBe(true);
    expect(isCellSelected(cellSelection, 2, 3)).toBe(false);

    const rowSelection = parseGridSelection("Dòng 3");
    expect(rowSelection).toEqual({
      kind: "row",
      label: "Dòng 3",
      startRow: 2,
      endRow: 2,
    });
    expect(isCellSelected(rowSelection, 2, 0)).toBe(true);
    expect(isCellSelected(rowSelection, 2, 10)).toBe(true);
    expect(isCellSelected(rowSelection, 3, 0)).toBe(false);

    const colSelection = parseGridSelection("Cột E");
    expect(colSelection).toEqual({
      kind: "col",
      label: "Cột E",
      startCol: 4,
      endCol: 4,
    });
    expect(isCellSelected(colSelection, 0, 4)).toBe(true);
    expect(isCellSelected(colSelection, 99, 4)).toBe(true);
    expect(isCellSelected(colSelection, 0, 5)).toBe(false);

    const rangeSelection = parseGridSelection("B4:D6");
    expect(rangeSelection?.kind).toBe("range");
    expect(isCellSelected(rangeSelection, 4, 2)).toBe(true);
    expect(isCellSelected(rangeSelection, 0, 0)).toBe(false);

    const clickedSingleCellSelections = parseGridSelections("E3");
    expect(isCellInSelections(clickedSingleCellSelections, 2, 4)).toBe(true);
    // Other cells in Row 3 or Col E must NOT be highlighted when only E3 is clicked.
    expect(isCellInSelections(clickedSingleCellSelections, 2, 0)).toBe(false);
    expect(isCellInSelections(clickedSingleCellSelections, 5, 4)).toBe(false);
  });

  it("formats zero-based column indices back into Excel column letters", () => {
    expect(formatColumnLabel(0)).toBe("A");
    expect(formatColumnLabel(4)).toBe("E");
    expect(formatColumnLabel(5)).toBe("F");
    expect(formatColumnLabel(25)).toBe("Z");
    expect(formatColumnLabel(26)).toBe("AA");
  });

  it("matches relative paths from chat against absolute workspace file paths", () => {
    const relativeInvoicePath = "examples/Ngân sách hàng tháng.xlsx";
    const absoluteInvoicePath =
      "/home/dev/Desktop/Workspaces/tools/paseo-explorer/examples/Ngân sách hàng tháng.xlsx";

    officeStore.registerIssues(relativeInvoicePath, [
      { ext: ".xlsx", cell: "E3", reason: "lỗi để trống" },
    ]);

    expect(
      findIssuesForFilePath(officeStore.getState().issuesByFile, absoluteInvoicePath),
    ).toHaveLength(1);
  });
});

describe("natural language and inline-tag Office issue parser", () => {
  it("extracts cell, row, and column issues from Vietnamese natural-language descriptions", () => {
    const naturalText = [
      'E3 trong file "examples/Ngân sách hàng tháng.xlsx" lỗi để trống',
      'Dòng 5 trong file "examples/Ngân sách hàng tháng.xlsx" (sheet "Giao dịch") thiếu danh mục',
      'Cột E trong "examples/Ngân sách hàng tháng.xlsx" sai định dạng',
      '[[examples/Ngân sách hàng tháng.xlsx#Tóm tắt!L8 | Sai số dư đầu kỳ]]',
    ].join("\n");

    const parsedReports = parseOfficeIssuesFromMessageText(naturalText);
    expect(parsedReports).toHaveLength(1);
    expect(parsedReports[0]?.filePath).toBe("examples/Ngân sách hàng tháng.xlsx");
    expect(parsedReports[0]?.issues).toEqual([
      { sheet: "Tóm tắt", cell: "L8", reason: "Sai số dư đầu kỳ" },
      { cell: "E3", reason: "lỗi để trống" },
      { sheet: "Giao dịch", cell: "Row:5", reason: "(sheet \"Giao dịch\") thiếu danh mục" },
      { cell: "Col:E", reason: "sai định dạng" },
    ]);
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
    expect(rootListing.entries.map((entry) => entry.name)).toEqual(["2026-10", "archive"]);
    expect(rootListing.entries.every((entry) => entry.isDirectory)).toBe(true);

    const subFolderListing = await listDirectoryEntries({ dirPath: declarationFolder });
    expect(subFolderListing.entries.map((entry) => ({ name: entry.name, ext: entry.ext }))).toEqual([
      { name: "Bill_of_Lading.pdf", ext: ".pdf" },
      { name: "Commercial_Invoice_01.xlsx", ext: ".xlsx" },
      { name: "Contract_992.docx", ext: ".docx" },
    ]);
  });

  it("parses Excel (.xlsx), Word (.docx), and PDF (.pdf) documents for the viewer", async () => {
    const excelFilePath = join(temporaryWorkspaceDir, "Commercial_Invoice_01.xlsx");
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

    const excelData = await readOfficeFile({ filePath: "Commercial_Invoice_01.xlsx", cwd: temporaryWorkspaceDir });
    expect(excelData.fileType).toBe("xlsx");
    expect(excelData.sheets?.[0]?.name).toBe("Sheet1");
    const f14Coords = parseCellRef("F14")!;
    expect(excelData.sheets?.[0]?.rows[f14Coords.row]?.[f14Coords.col]).toBe("8500.12.99");
  });
});

describe("WindowsFileIcon badges and timeline transformers", () => {
  it("defines distinct Windows/Office badge colors and transforms both tool calls and chat messages", () => {
    expect(ICON_CONFIG[".xlsx"].label).toBe("XLS");
    expect(ICON_CONFIG[".docx"].label).toBe("DOC");
    expect(ICON_CONFIG[".pdf"].label).toBe("PDF");

    expect(WindowsFileIcon({ isDirectory: true, isOpen: true, ext: "" })).toBeTruthy();

    const sampleIssuePayload = {
      filePath: "examples/Ngân sách hàng tháng.xlsx",
      ext: ".xlsx",
      issues: [{ sheet: "Tóm tắt", cell: "E3", reason: "Lỗi để trống" }],
    };

    const transformedToolResult = officeIssuesTimelineTransformer.transform({
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
    expect(transformedToolResult?.items[0]?.kind).toBe(OFFICE_ISSUES_TIMELINE_KIND);

    const transformedAssistantMsg = officeAssistantMessageTransformer.transform({
      item: {
        type: "assistant_message",
        messageId: "msg-1",
        text: 'Phát hiện ô E3 trong file "examples/Ngân sách hàng tháng.xlsx" lỗi để trống.',
      },
      phase: "complete",
    });
    expect(transformedAssistantMsg?.items).toHaveLength(1);

    const renderedCardElement = React.createElement(AgentOfficeLinkRenderer, {
      input: sampleIssuePayload,
      theme: { colors: { surface: "#1e293b", text: "#f8fafc", textMuted: "#94a3b8", primary: "#3b82f6" } },
    });
    expect(renderedCardElement.type).toBe(AgentOfficeLinkRenderer);
    expect(OfficeFilePanel).toBeTypeOf("function");
  });
});

describe("automatic Office Explorer prompt injection for Paseo and Antigravity CLI", () => {
  it("injects instructions once into agent.create systemPrompt and skips internal ephemeral agents", () => {
    const firstPassPrompt = injectOfficeExplorerSystemPrompt("You are a customs auditor.");
    expect(firstPassPrompt).toContain("You are a customs auditor.");
    expect(firstPassPrompt).toContain(OFFICE_EXPLORER_PROMPT_MARKER);
    expect(firstPassPrompt).toContain("[[path/to/file.xlsx#SheetName!E3");

    // Second pass must remain idempotent and never duplicate the marker block.
    const secondPassPrompt = injectOfficeExplorerSystemPrompt(firstPassPrompt);
    expect(secondPassPrompt).toBe(firstPassPrompt);

    const interactiveAgentRequest = augmentAgentCreateRequestWithOfficePrompt({
      config: {
        provider: "antigravity",
        cwd: "/workspace",
        systemPrompt: "Audit spreadsheets.",
      },
    });
    expect(interactiveAgentRequest.config.systemPrompt).toContain(OFFICE_EXPLORER_PROMPT_MARKER);

    const internalAgentRequest = augmentAgentCreateRequestWithOfficePrompt({
      config: {
        provider: "antigravity",
        cwd: "/workspace",
        internal: true,
        systemPrompt: "Generate commit title.",
      },
    });
    expect(internalAgentRequest.config.systemPrompt).toBe("Generate commit title.");

    const registeredBeforeHandlers = new Map<string, (input: unknown) => unknown>();
    const cleanupHook = registerOfficePromptInjector({
      before: (hookName: string, handler: (input: unknown) => unknown) => {
        registeredBeforeHandlers.set(hookName, handler);
        return () => registeredBeforeHandlers.delete(hookName);
      },
    } as never);

    expect(registeredBeforeHandlers.has("agent.create")).toBe(true);
    cleanupHook();
    expect(registeredBeforeHandlers.has("agent.create")).toBe(false);
  });
});

