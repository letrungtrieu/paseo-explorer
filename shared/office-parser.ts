import type { OfficeIssueItem, OfficeIssueReportPayload } from "./viewer";

const OFFICE_EXT_PATTERN = /\.(xlsx|xls|csv|docx|doc|pdf)$/i;

// Extract lowercase extension (with leading dot) without importing node:path into the client/shared bundle.
function extractExtensionFromPath(filePath: string): string {
  const lastDotIndex = filePath.lastIndexOf(".");
  return lastDotIndex >= 0 ? filePath.slice(lastDotIndex).toLowerCase() : ".xlsx";
}

// Clean surrounding quotes/brackets from extracted file paths while preserving internal spaces (e.g. "examples/Ngân sách hàng tháng.xlsx").
function sanitizeExtractedFilePath(rawPathCandidate: string): string {
  return rawPathCandidate
    .trim()
    .replace(/^["'`[\]()<>]+/, "")
    .replace(/["'`[\]()<>,.;:!?]+$/, "")
    .trim();
}

// Normalize a sheet!target expression (e.g. "Tóm tắt!E3", "Row:5", "Cột E", "E3") into { sheet, cell, page }.
function parseTargetSpecifier(rawSpecifier: string): {
  sheet?: string;
  cell?: string;
  page?: number;
} {
  const trimmedSpecifier = rawSpecifier.trim();
  const bangIndex = trimmedSpecifier.indexOf("!");
  const extractedSheet =
    bangIndex > 0 ? trimmedSpecifier.slice(0, bangIndex).trim().replace(/^['"]|['"]$/g, "") : undefined;
  const coordinatePart = (bangIndex > 0 ? trimmedSpecifier.slice(bangIndex + 1) : trimmedSpecifier).trim();

  const pageMatch = coordinatePart.match(/^(?:page|trang)\s*[:#-]?\s*(\d+)$/i);
  if (pageMatch) {
    return {
      ...(extractedSheet ? { sheet: extractedSheet } : {}),
      page: Number.parseInt(pageMatch[1], 10),
    };
  }

  return {
    ...(extractedSheet ? { sheet: extractedSheet } : {}),
    cell: coordinatePart,
  };
}

// Parse natural-language or inline-tag references to Office files + cells/rows/cols from chat messages.
export function parseOfficeIssuesFromMessageText(messageText: string): OfficeIssueReportPayload[] {
  if (!messageText || !/\.(xlsx|xls|csv|docx|doc|pdf)\b/i.test(messageText)) {
    return [];
  }

  const groupedIssuesByFile = new Map<string, { ext: string; issues: OfficeIssueItem[] }>();

  const appendIssue = (filePath: string, issueItem: OfficeIssueItem) => {
    const cleanFilePath = sanitizeExtractedFilePath(filePath);
    if (!OFFICE_EXT_PATTERN.test(cleanFilePath)) return;
    const normalizedExt = extractExtensionFromPath(cleanFilePath);

    const existingGroup = groupedIssuesByFile.get(cleanFilePath) ?? {
      ext: normalizedExt,
      issues: [],
    };
    const isDuplicateIssue = existingGroup.issues.some(
      (existing) =>
        (existing.sheet ?? "") === (issueItem.sheet ?? "") &&
        (existing.cell ?? "").toUpperCase() === (issueItem.cell ?? "").toUpperCase() &&
        existing.page === issueItem.page,
    );
    if (!isDuplicateIssue) {
      existingGroup.issues.push(issueItem);
      groupedIssuesByFile.set(cleanFilePath, existingGroup);
    }
  };

  // 1. Parse explicit double-bracket tags: [[path/to/file.xlsx#Sheet!E3 | Reason]] or [[path/to/file.xlsx#Row:3 | Reason]]
  const bracketTagRegex = /\[\[([^\]|#]+?\.(?:xlsx|xls|csv|docx|doc|pdf))(?:#([^\]|]+))?(?:\|\s*([^\]]+))?\]\]/gi;
  for (const tagMatch of messageText.matchAll(bracketTagRegex)) {
    const [, matchedPath, matchedTarget = "A1", matchedReason] = tagMatch;
    const targetParts = parseTargetSpecifier(matchedTarget);
    appendIssue(matchedPath, {
      ...targetParts,
      reason: (matchedReason ?? "Cần kiểm tra").trim(),
    });
  }

  // 2. Parse line-by-line natural Vietnamese / English descriptions (e.g. 'E3 trong file "examples/Ngân sách hàng tháng.xlsx" lỗi để trống').
  const messageLines = messageText.split(/\r?\n/);
  for (const rawLine of messageLines) {
    const lineWithoutTags = rawLine.replace(bracketTagRegex, "").trim();
    if (!lineWithoutTags) continue;

    // Match either a quoted path ("examples/Ngân sách hàng tháng.xlsx") or an unquoted path token ending in an Office extension.
    const quotedFileMatch = lineWithoutTags.match(
      /["'`]([^"'`\n]+?\.(?:xlsx|xls|csv|docx|doc|pdf))(?:#([^"'`\n]+))?["'`]/i,
    );
    const unquotedFileMatch = !quotedFileMatch
      ? lineWithoutTags.match(/([^\s"'`[\]()<>]+?\.(?:xlsx|xls|csv|docx|doc|pdf))(?![\w.])/i)
      : null;

    const detectedFilePath = quotedFileMatch?.[1] ?? unquotedFileMatch?.[1];
    if (!detectedFilePath) continue;

    // Extract optional sheet name mentioned in the line, e.g. sheet "Tóm tắt" or trang tính 'Giao dịch'.
    const sheetMatch =
      lineWithoutTags.match(/(?:sheet|trang\s+tính|tab)\s+["'`]([^"'`]+)["'`]/i) ??
      lineWithoutTags.match(/\[([^\]!]+)!([A-Za-z]+\d+|\d+:\d+|[A-Za-z]+:[A-Za-z]+)\]/);
    const detectedSheet = sheetMatch?.[1]?.trim();

    // Strip the file path clause from the line so we don't mistake numbers inside the filename for cell references.
    const fullMatchedFileToken = quotedFileMatch?.[0] ?? unquotedFileMatch?.[0] ?? detectedFilePath;
    const textAroundFile = lineWithoutTags.replace(fullMatchedFileToken, " ");

    const extractedCoordinates: string[] = [];

    // 2a. Detect row references: "dòng 3", "hàng 3", "row 3", "Row:3"
    for (const rowMatch of textAroundFile.matchAll(/\b(?:dòng|dong|hàng|hang|row)\s*[:#-]?\s*(\d+)\b/gi)) {
      extractedCoordinates.push(`Row:${rowMatch[1]}`);
    }

    // 2b. Detect column references: "cột E", "col E", "column E", "Col:E"
    for (const colMatch of textAroundFile.matchAll(/\b(?:cột|cot|col|column)\s*[:#-]?\s*([A-Za-z]{1,3})\b/gi)) {
      extractedCoordinates.push(`Col:${colMatch[1].toUpperCase()}`);
    }

    // 2c. Detect cell ranges ("B4:E8") and A1-style cells ("E3", "F14", "L8")
    for (const rangeMatch of textAroundFile.matchAll(/\b([A-Za-z]{1,3}[1-9]\d*\s*:\s*[A-Za-z]{1,3}[1-9]\d*)\b/g)) {
      extractedCoordinates.push(rangeMatch[1].replace(/\s+/g, "").toUpperCase());
    }
    for (const cellMatch of textAroundFile.matchAll(/\b([A-Za-z]{1,3}[1-9]\d{0,4})\b/g)) {
      const candidateCell = cellMatch[1].toUpperCase();
      // Ignore already-captured range endpoints.
      const isPartOfRange = extractedCoordinates.some(
        (existingCoord) => existingCoord.includes(":") && existingCoord.includes(candidateCell),
      );
      if (!isPartOfRange) {
        extractedCoordinates.push(candidateCell);
      }
    }

    if (extractedCoordinates.length === 0 && quotedFileMatch?.[2]) {
      extractedCoordinates.push(quotedFileMatch[2]);
    }
    if (extractedCoordinates.length === 0) continue;

    // Extract the descriptive reason (e.g. "lỗi để trống") by trimming connector phrases around the file & cell.
    let descriptiveReason = textAroundFile
      .replace(/\b(?:trong\s+file|ở\s+file|tại\s+file|trong|file|tập\s+tin)\b/gi, " ")
      .replace(/\b(?:ô|cell|dòng|dong|hàng|hang|row|cột|cot|col|column)\s*[:#-]?\s*[A-Za-z0-9:]+\b/gi, " ")
      .replace(/^[\s,.:;\-–—]+|[\s,.:;\-–—]+$/g, "")
      .replace(/\s+/g, " ")
      .trim();

    for (const coordToken of extractedCoordinates) {
      const escapedToken = coordToken.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      descriptiveReason = descriptiveReason
        .replace(new RegExp(`^${escapedToken}\\b\\s*[,.:;\\-–—]*`, "i"), "")
        .trim();
    }
    if (!descriptiveReason) {
      descriptiveReason = "Cần kiểm tra";
    }

    for (const coordinateToken of extractedCoordinates) {
      appendIssue(detectedFilePath, {
        ...(detectedSheet ? { sheet: detectedSheet } : {}),
        cell: coordinateToken,
        reason: descriptiveReason,
      });
    }
  }

  return Array.from(groupedIssuesByFile.entries()).map(([filePath, groupData]) => ({
    filePath,
    ext: groupData.ext,
    issues: groupData.issues,
  }));
}
