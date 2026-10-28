import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createHash } from "node:crypto";
import { deflateSync } from "node:zlib";
import JSZip from "jszip";
import * as XLSX from "xlsx";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { readOfficeFile } from "../server/viewer";
import { clearWorkbookCache } from "../server/spreadsheet";
import { OfficeFileDataSchema, readOfficeFileRpc } from "../shared/viewer";
import { WORKSHEET_WINDOW_COLUMNS, WORKSHEET_WINDOW_ROWS, WORKSHEET_MAX_WINDOW_ROWS } from "../shared/spreadsheet-window";

let workspaceDirectory: string;
beforeEach(async () => {
  // Isolate file revisions and caches so these regressions do not depend on test ordering.
  workspaceDirectory = await mkdtemp(join(tmpdir(), "office-document-test-"));
  clearWorkbookCache();
});
afterEach(async () => {
  clearWorkbookCache();
  await rm(workspaceDirectory, { recursive: true, force: true });
});

async function saveWorkbook(worksheet: XLSX.WorkSheet, bookType: XLSX.BookType = "xlsx") {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Giao dịch");
  // A second sheet proves unselected tabs transmit metadata instead of duplicating their cell matrices.
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["Tổng cộng"]]), "Tóm tắt");
  const filePath = join(workspaceDirectory, `report.${bookType}`);
  await writeFile(filePath, XLSX.write(workbook, { type: "buffer", bookType, compression: true }));
  return filePath;
}

describe("bounded spreadsheet preview", () => {
  it("keeps absolute sparse coordinates, empty rows, and formatted values on distant windows", async () => {
    const worksheet: XLSX.WorkSheet = {
      M102: { t: "n", v: 1234.5, z: "#,##0.00" },
      N103: { t: "n", v: 0.15, z: "0.0%" },
      O104: { t: "n", v: 10, f: "5+5" },
      P105: { t: "n", v: 45292, z: "yyyy-mm-dd" },
      "!ref": "M102:P105",
    };
    const filePath = await saveWorkbook(worksheet);
    const preview = await readOfficeFile({ filePath, rowStart: 100, columnStart: 12 });
    const worksheetWindow = preview.sheets![0];
    expect(worksheetWindow).toMatchObject({ rowCount: 105, columnCount: 16, rowStart: 100, columnStart: 12 });
    expect(worksheetWindow.rows[0]).toEqual(["", "", "", ""]);
    expect(worksheetWindow.rows[1][0]).toBe("1,234.50");
    expect(worksheetWindow.rows[2][1]).toBe("15.0%");
    // Cached formula results and date number formats must survive the faster read-only parse options.
    expect(worksheetWindow.rows[3][2]).toBe("10");
    expect(worksheetWindow.rows[4][3]).toBe("2024-01-01");
    expect(preview.sheets![1].rows).toEqual([]);
    expect(OfficeFileDataSchema.safeParse(preview).success).toBe(true);
  });

  it("bounds inflated Excel ranges and clamps jumps past the end", async () => {
    const filePath = await saveWorkbook(XLSX.utils.aoa_to_sheet([["First"]]));
    // Patch only the declared dimension: constructing this through SheetJS would itself walk 17 billion cells.
    const workbookArchive = await JSZip.loadAsync(await readFile(filePath));
    const worksheetXml = await workbookArchive.file("xl/worksheets/sheet1.xml")!.async("string");
    workbookArchive.file("xl/worksheets/sheet1.xml", worksheetXml.replace(/<dimension ref="[^"]+"\/>/, '<dimension ref="A1:XFD1048576"/>'));
    await writeFile(filePath, await workbookArchive.generateAsync({ type: "nodebuffer" }));
    const initialPreview = await readOfficeFile({ filePath });
    expect(initialPreview.sheets![0].rows).toHaveLength(WORKSHEET_WINDOW_ROWS);
    expect(initialPreview.sheets![0].rows[0]).toHaveLength(WORKSHEET_WINDOW_COLUMNS);
    // Formatting an entire column must never expand into a million-row JSON response.
    const finalPreview = await readOfficeFile({ filePath, rowStart: 2_000_000, columnStart: 20_000 });
    expect(finalPreview.sheets![0]).toMatchObject({ rowStart: 1_048_575, columnStart: 16_383, rows: [[""]] });
  });

  it("loads another sheet case-insensitively and invalidates cached workbooks after an edit", async () => {
    const filePath = await saveWorkbook(XLSX.utils.aoa_to_sheet([["Before"]]));
    const secondSheetPreview = await readOfficeFile({ filePath, sheet: "tÓM tẮT" });
    expect(secondSheetPreview.sheets![0].rows).toEqual([]);
    expect(secondSheetPreview.sheets![1].rows).toEqual([["Tổng cộng"]]);
    await readOfficeFile({ filePath });
    // Same-path replacements must be detected without relying on a client cache reset.
    await saveWorkbook(XLSX.utils.aoa_to_sheet([["After editing the workbook"]]));
    expect((await readOfficeFile({ filePath })).sheets![0].rows).toEqual([["After editing the workbook"]]);
  });

  it("handles empty sheets and legacy XLS/CSV through the same bounded contract", async () => {
    const emptyFilePath = await saveWorkbook({});
    expect((await readOfficeFile({ filePath: emptyFilePath })).sheets![0]).toMatchObject({ rowCount: 0, columnCount: 0, rows: [] });
    for (const bookType of ["xls", "csv"] as const) {
      const legacyFilePath = await saveWorkbook(XLSX.utils.aoa_to_sheet([["Mã", "Số tiền"], ["VN-01", 42]]), bookType);
      expect((await readOfficeFile({ filePath: legacyFilePath })).sheets![0].rows[1]).toEqual(["VN-01", "42"]);
    }
    // Reject invalid offsets at the RPC boundary instead of allocating an unbounded or negative range.
    expect(readOfficeFileRpc.input.safeParse({ filePath: emptyFilePath, rowStart: -1 }).success).toBe(false);
    // Prefetch may add one chunk, but arbitrary callers cannot turn a scroll request into a full-sheet response.
    expect(readOfficeFileRpc.input.safeParse({ filePath: emptyFilePath, rowLimit: WORKSHEET_MAX_WINDOW_ROWS + 1 }).success).toBe(false);
  });

  it("keeps a megabyte workbook response small and reuses parsed cells on later windows", async () => {
    const workbookRows = Array.from({ length: 6000 }, (_, rowIndex) => Array.from({ length: 16 }, (_, columnIndex) =>
      createHash("sha256").update(`${rowIndex}:${columnIndex}`).digest("hex").slice(0, 16)));
    const filePath = await saveWorkbook(XLSX.utils.aoa_to_sheet(workbookRows));
    const fileBuffer = await readFile(filePath);
    const coldReadStart = performance.now();
    const initialPreview = await readOfficeFile({ filePath });
    const coldReadMs = performance.now() - coldReadStart;
    const warmReadStart = performance.now();
    const distantPreview = await readOfficeFile({ filePath, rowStart: 5900, columnStart: 12 });
    const warmReadMs = performance.now() - warmReadStart;
    expect(initialPreview.sheets![0].rows).toHaveLength(WORKSHEET_WINDOW_ROWS);
    expect(distantPreview.sheets![0].rows[0][0]).toBe(workbookRows[5900][12]);
    expect(fileBuffer.byteLength).toBeGreaterThan(1_000_000);
    // The larger buffer is proportional to visible cells, still far smaller than the full workbook matrix.
    expect(JSON.stringify(initialPreview).length).toBeLessThan(35_000);
    // Report timings without flaky wall-clock assertions; the enforceable regression is bounded work and payload.
    console.log(JSON.stringify({ workbookBytes: fileBuffer.byteLength, previousMatrixBytes: JSON.stringify(workbookRows).length,
      previewBytes: JSON.stringify(initialPreview).length, coldReadMs: Math.round(coldReadMs), warmReadMs: Math.round(warmReadMs) }));
  }, 20_000);
});

describe("Word and PDF document preview", () => {
  it("converts a real zipped DOCX into readable HTML including tables and embedded images", async () => {
    const wordArchive = new JSZip();
    wordArchive.file("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>');
    // Use genuine OOXML relationships, rather than raw XML disguised with a .docx extension.
    wordArchive.file("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    wordArchive.file("word/_rels/document.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="image1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/></Relationships>');
    wordArchive.file("word/media/logo.png", "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=", { base64: true });
    wordArchive.file("word/document.xml", `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hợp đồng Việt Nam</w:t></w:r></w:p>
      <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Đơn giá</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
      <w:p><w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="Logo"/>
      <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="image1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>
      </wp:inline></w:drawing></w:r></w:p></w:body></w:document>`);
    const filePath = join(workspaceDirectory, "contract.docx");
    await writeFile(filePath, await wordArchive.generateAsync({ type: "nodebuffer" }));
    const wordPreview = await readOfficeFile({ filePath });
    expect(wordPreview.htmlContent).toContain("<strong>Hợp đồng Việt Nam</strong>");
    expect(wordPreview.htmlContent).toContain("<table>");
    expect(wordPreview.htmlContent).toContain('src="data:image/png;base64,');
    expect(wordPreview.textContent).toContain("Đơn giá");
  });

  it("preserves compressed PDF bytes so a renderer can show graphics and blank/scanned pages", async () => {
    const pdfStream = deflateSync(Buffer.from("BT /F1 12 Tf 30 70 Td (Compressed PDF) Tj ET"));
    const pdfObjects = ["<< /Type /Catalog /Pages 2 0 R >>", "<< /Type /Pages /Kids [3 0 R 6 0 R] /Count 2 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 100] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
      "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
      Buffer.concat([Buffer.from(`<< /Length ${pdfStream.length} /Filter /FlateDecode >>\nstream\n`), pdfStream, Buffer.from("\nendstream")]),
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 100] >>"];
    // Real cross-reference offsets make this a valid two-page PDF, including a deliberately blank page.
    let pdfFile = Buffer.from("%PDF-1.4\n");
    const objectOffsets = [0];
    pdfObjects.forEach((pdfObject, objectIndex) => {
      objectOffsets.push(pdfFile.length);
      pdfFile = Buffer.concat([pdfFile, Buffer.from(`${objectIndex + 1} 0 obj\n`), Buffer.from(pdfObject), Buffer.from("\nendobj\n")]);
    });
    const xrefOffset = pdfFile.length;
    pdfFile = Buffer.concat([pdfFile, Buffer.from(`xref\n0 7\n0000000000 65535 f \n${objectOffsets.slice(1).map((offset) => `${String(offset).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size 7 /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`)]);
    const filePath = join(workspaceDirectory, "scanned.pdf");
    await writeFile(filePath, pdfFile);
    const pdfPreview = await readOfficeFile({ filePath });
    expect(pdfPreview.fileType).toBe("pdf");
    expect(Buffer.from(pdfPreview.documentBase64!, "base64")).toEqual(pdfFile);
    expect(pdfPreview.textContent).toBeUndefined();
  });

  it("reports invalid documents and legacy DOC instead of displaying binary noise", async () => {
    for (const extension of ["docx", "pdf", "doc"]) {
      const filePath = join(workspaceDirectory, `broken.${extension}`);
      await writeFile(filePath, Buffer.from([0, 255, 10, 21]));
      await expect(readOfficeFile({ filePath })).rejects.toThrow();
    }
    await expect(readOfficeFile({ filePath: join(workspaceDirectory, "missing.xlsx") })).rejects.toThrow();
  });
});
