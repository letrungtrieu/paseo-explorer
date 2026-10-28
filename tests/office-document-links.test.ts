import { describe, expect, it } from "vitest";
import {
  convertOfficeTagsToMarkdownLinks,
  parseOfficeIssuesFromMessageText,
  parseOfficeJumpUrl,
} from "../shared/office-parser";

describe("Office document tags without coordinates", () => {
  // The public tag format includes spaces around the separator even when Word has no cell or page target.
  it.each(["docx", "doc", "pdf", "xlsx", "xls", "csv"])(
    "converts and registers a spaced %s tag as the same navigation target",
    (extension) => {
      // Vietnamese file names must survive URL encoding so clicks still resolve to the on-disk path.
      const filePath = `examples/Tài liệu mẫu.${extension}`;
      const chatMessage = `Mở [[${filePath} | Xem tài liệu mẫu]] để test.`;
      const convertedMessage = convertOfficeTagsToMarkdownLinks(chatMessage);
      expect(convertedMessage.hasConvertedInlineTags).toBe(true);
      expect(convertedMessage.markdownText).not.toContain("[[");

      // Conversion and issue extraction must agree or the timeline can discard an otherwise valid hyperlink.
      const jumpHref = convertedMessage.markdownText.match(/\]\(([^)]+)\)/)?.[1];
      expect(parseOfficeJumpUrl(jumpHref ?? "")).toMatchObject({
        filePath, ext: `.${extension}`, reason: "Xem tài liệu mẫu",
      });
      const parsedReports = parseOfficeIssuesFromMessageText(chatMessage);
      expect(parsedReports).toHaveLength(1);
      expect(parsedReports[0]).toMatchObject({
        filePath, ext: `.${extension}`, issues: [{ reason: "Xem tài liệu mẫu" }],
      });
    },
  );

  // Omitted labels and compact separators are also supported by the optional parts of the tag format.
  it.each([
    "[[examples/demo.docx]]",
    "[[examples/demo.docx ]]",
    "[[examples/demo.docx|Mở Word]]",
  ])("keeps the Word path intact for %s", (chatTag) => {
    expect(convertOfficeTagsToMarkdownLinks(chatTag).hasConvertedInlineTags).toBe(true);
    expect(parseOfficeIssuesFromMessageText(chatTag)[0]?.filePath).toBe("examples/demo.docx");
  });
});
