import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import React from "react";
import { act, create, type ReactTestRenderer } from "react-test-renderer";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { Image, Pressable, Text } from "react-native";
import { readOfficeFile } from "../server/viewer";
import { clearWorkbookCache } from "../server/spreadsheet";
import { resolveSpreadsheetColor, readSpreadsheetTheme } from "../server/spreadsheet-colors";
import { createSpreadsheetGeometry } from "../client/spreadsheet-geometry";
import { SpreadsheetCanvas } from "../client/spreadsheet-canvas";
import { officeStore, parseGridSelections } from "../shared/office-store";
import { OfficeFileDataSchema } from "../shared/viewer";

// Render the real canvas in Node so assertions cover visible styles and drawings, not just serialized metadata.
vi.mock("react-native", () => ({ View: "View", Text: "Text", Image: "Image", Pressable: "Pressable" }));
const embeddedPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=";
let workspaceDirectory: string;
let renderedCanvas: ReactTestRenderer | undefined;

beforeEach(async () => {
  (globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
  workspaceDirectory = await mkdtemp(join(tmpdir(), "office-presentation-"));
  clearWorkbookCache();
  officeStore.reset();
});
afterEach(async () => {
  // Unmount before removing files so no component can retain state from another workbook's preview.
  await act(async () => renderedCanvas?.unmount());
  renderedCanvas = undefined;
  clearWorkbookCache();
  await rm(workspaceDirectory, { recursive: true, force: true });
});

async function saveStyledWorkbook() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Báo giá", { properties: { defaultRowHeight: 15 } });
  worksheet.getColumn(1).width = 24;
  worksheet.getColumn(2).width = 18;
  worksheet.getRow(1).height = 30;
  // Empty spacer rows must retain their heights too; row iteration that only visits values misses them.
  worksheet.getRow(2).height = 24;
  worksheet.getCell("A1").value = "BÁO GIÁ VIỆT NAM";
  worksheet.getCell("A1").font = { name: "Arial", size: 18, bold: true, italic: true, color: { argb: "FFFFFFFF" } };
  worksheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF245A81" } };
  worksheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  worksheet.getCell("A1").border = { bottom: { style: "double", color: { argb: "FF00AA00" } } };
  worksheet.mergeCells("A1:C2");
  worksheet.getCell("A3").value = "Ghi chú dài cần xuống dòng trong cùng một ô";
  worksheet.getCell("A3").alignment = { wrapText: true, horizontal: "left", vertical: "top" };
  worksheet.getRow(3).height = 45;
  worksheet.getCell("B4").value = 1234.5;
  worksheet.getCell("B4").numFmt = "#,##0.00";
  // Theme colors and blank styled cells are independent of the stored display text.
  worksheet.getCell("C4").fill = { type: "pattern", pattern: "solid", fgColor: { theme: 4 } };
  worksheet.getCell("C4").font = { underline: true, strike: true, color: { argb: "FFCC0000" } };
  worksheet.getCell("D5").value = "Ẩn";
  worksheet.getColumn(4).hidden = true;
  worksheet.getRow(5).hidden = true;
  worksheet.getCell("N52").value = "Cuối sheet";
  // This region deliberately crosses both preview axes; its anchor is outside the later requested window.
  worksheet.mergeCells("K49:N52");
  worksheet.getCell("K49").value = "Tiếp tục vùng gộp";
  worksheet.getCell("K49").font = { bold: true };
  const imageId = workbook.addImage({ base64: embeddedPng, extension: "png" });
  worksheet.addImage(imageId, { tl: { col: 1, row: 48 }, ext: { width: 120, height: 100 } });
  worksheet.addImage(imageId, "E8:F10");
  const illustration = workbook.addWorksheet("Ảnh");
  illustration.addImage(imageId, { tl: { col: 2, row: 3 }, ext: { width: 160, height: 90 } });
  const filePath = join(workspaceDirectory, "styled.xlsx");
  await workbook.xlsx.writeFile(filePath);
  return filePath;
}

describe("spreadsheet presentation from actual XLSX files", () => {
  it("keeps fonts, fills, separate borders, alignment, sizes and styled empty cells", async () => {
    const preview = await readOfficeFile({ filePath: await saveStyledWorkbook() });
    expect(OfficeFileDataSchema.safeParse(preview).success).toBe(true);
    const sheet = preview.sheets![0];
    const presentation = sheet.presentation!;
    const titleStyle = presentation.styles![presentation.merges![0].styleId];
    expect(titleStyle).toMatchObject({ fontFamily: "Arial", fontSize: 24, bold: true, italic: true,
      color: "#FFFFFF", background: "#245A81", horizontal: "center", vertical: "center",
      bottom: { style: "double", width: 2, color: "#00AA00" } });
    expect(presentation.columnWidths.slice(0, 4)).toEqual([173, 131, 64, 0]);
    expect(presentation.rowHeights.slice(0, 5)).toEqual([40, 32, 60, 20, 0]);
    expect(sheet.rows[3][1]).toBe("1,234.50");
    expect(presentation.styles![presentation.cellStyleIds![3][1]].horizontal).toBe("right");
    // A blank themed fill must still paint, even though format_cell returns an empty string.
    expect(presentation.styles![presentation.cellStyleIds![3][2]]).toMatchObject({
      background: "#4F81BD", underline: true, strike: true, color: "#CC0000",
    });
    expect(presentation.merges![0]).toMatchObject({ startRow: 0, endRow: 1, startColumn: 0, endColumn: 2, text: "BÁO GIÁ VIỆT NAM" });
    expect(preview.sheets![1].presentation).toBeUndefined();
  });

  it("keeps embedded PNG bytes and both one-cell and two-cell image geometry", async () => {
    const filePath = await saveStyledWorkbook();
    const initialSheet = (await readOfficeFile({ filePath })).sheets![0];
    const image = initialSheet.presentation!.images!.find((drawing) => drawing.width === 120)!;
    expect(image.uri).toBe(`data:image/png;base64,${embeddedPng}`);
    expect(image.left).toBe(173);
    expect(image.top).toBe(1012);
    const rangeImage = initialSheet.presentation!.images!.find((drawing) => drawing.width !== 120)!;
    expect(rangeImage).toMatchObject({ left: 368, top: 192, width: 128, height: 60 });
    // The image starts above the later window but its overlapping lower part must remain visible.
    const laterSheet = (await readOfficeFile({ filePath, rowStart: 50 })).sheets![0];
    expect(laterSheet.presentation!.images![0]).toMatchObject({ top: -40, width: 120, height: 100 });
    const distantSheet = (await readOfficeFile({ filePath, rowStart: 50, columnStart: 12 })).sheets![0];
    expect(distantSheet.presentation!.images).toBeUndefined();
    expect(distantSheet.presentation!.merges![0]).toMatchObject({ startRow: 48, startColumn: 10, text: "Tiếp tục vùng gộp" });
  });

  it("shows an image-only worksheet and extends its dimensions beyond populated cells", async () => {
    const preview = await readOfficeFile({ filePath: await saveStyledWorkbook(), sheet: "Ảnh" });
    const illustration = preview.sheets![1];
    expect(illustration.rowCount).toBe(8);
    expect(illustration.columnCount).toBe(5);
    expect(illustration.presentation!.images![0]).toMatchObject({ left: 128, top: 60, width: 160, height: 90 });
    expect(illustration.rows).toHaveLength(8);
  });

  it("accounts for a hidden empty row when positioning an image below it", async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Ẩn dòng");
    worksheet.getRow(2).hidden = true;
    const imageId = workbook.addImage({ base64: embeddedPng, extension: "png" });
    worksheet.addImage(imageId, { tl: { col: 0, row: 2 }, ext: { width: 20, height: 20 } });
    const filePath = join(workspaceDirectory, "hidden.xlsx");
    await workbook.xlsx.writeFile(filePath);
    // ExcelJS's writer omits valueless hidden rows; put the valid OOXML row in explicitly, as other writers do.
    const archive = await JSZip.loadAsync(await readFile(filePath));
    const sheetXml = await archive.file("xl/worksheets/sheet1.xml")!.async("string");
    archive.file("xl/worksheets/sheet1.xml", sheetXml.replace(/<sheetData\s*\/>|<sheetData>[\s\S]*?<\/sheetData>/,
      '<sheetData><row r="2" hidden="1"/></sheetData>'));
    await writeFile(filePath, await archive.generateAsync({ type: "nodebuffer" }));
    // Empty hidden rows are layout metadata, even though they have no values for a row iterator to visit.
    const sheet = (await readOfficeFile({ filePath })).sheets![0];
    expect(sheet.presentation!.rowHeights).toEqual([20, 0, 20]);
    expect(sheet.presentation!.images![0].top).toBe(20);
  });

  it("reads package-root image paths from other XLSX writers without altering the original file", async () => {
    const filePath = await saveStyledWorkbook();
    const archive = await JSZip.loadAsync(await readFile(filePath));
    // This is the valid relationship form used by openpyxl, rather than ExcelJS's own writer.
    const sheetRelationsPath = "xl/worksheets/_rels/sheet1.xml.rels";
    const drawingRelationsPath = "xl/drawings/_rels/drawing1.xml.rels";
    archive.file(sheetRelationsPath, (await archive.file(sheetRelationsPath)!.async("string"))
      .replace('Target="../drawings/drawing1.xml"', 'Target="/xl/drawings/drawing1.xml"'));
    archive.file(drawingRelationsPath, (await archive.file(drawingRelationsPath)!.async("string"))
      .replaceAll('Target="../media/', 'Target="/xl/media/'));
    // Other writers also use the default drawing namespace, which is semantically identical to an xdr: prefix.
    const drawingPath = "xl/drawings/drawing1.xml";
    archive.file(drawingPath, (await archive.file(drawingPath)!.async("string"))
      .replaceAll("xdr:", "").replace('xmlns:xdr=', 'xmlns='));
    const originalBytes = await archive.generateAsync({ type: "nodebuffer" });
    await writeFile(filePath, originalBytes);
    const sheet = (await readOfficeFile({ filePath })).sheets![0];
    expect(sheet.presentation!.images).toHaveLength(2);
    expect(await readFile(filePath)).toEqual(originalBytes);
  });
});

describe("spreadsheet canvas rendering", () => {
  it("renders saved styles, merged cells and images while keeping chat selection and clicks", async () => {
    const sheet = (await readOfficeFile({ filePath: await saveStyledWorkbook() })).sheets![0];
    const jumpToCoordinate = vi.fn();
    await act(async () => {
      renderedCanvas = create(<SpreadsheetCanvas worksheet={sheet} geometry={createSpreadsheetGeometry(sheet)}
        selections={parseGridSelections("B2")} jumpToCoordinate={jumpToCoordinate} />);
    });
    const root = renderedCanvas!.root;
    const title = root.findAllByType(Pressable).find((cell) => cell.props.accessibilityLabel === "A1")!;
    expect(title.props.style).toMatchObject({ width: 368, height: 72, backgroundColor: "#245A81" });
    expect(title.findByType(Text).props.style).toMatchObject({ fontFamily: "Arial", fontSize: 24,
      fontWeight: "bold", fontStyle: "italic", color: "#FFFFFF", textAlign: "center" });
    // The covered coordinate highlights its merged anchor without creating a duplicate B2 cell.
    expect(root.findAllByType(Pressable).some((cell) => cell.props.accessibilityLabel === "B2")).toBe(false);
    expect(title.findByProps({ testID: "office-selection-A1" })).toBeTruthy();
    expect(title.findByProps({ testID: "office-selection-A1" }).props.style.backgroundColor).toBeUndefined();
    await act(async () => title.props.onPress());
    expect(jumpToCoordinate).toHaveBeenCalledWith("A1");
    const wrappedNote = root.findAllByType(Pressable).find((cell) => cell.props.accessibilityLabel === "A3")!;
    expect(wrappedNote.findByType(Text).props.numberOfLines).toBeUndefined();
    expect(wrappedNote.props.style.justifyContent).toBe("flex-start");
    // Images use the very same geometry as cells, and must not capture clicks intended for cells below them.
    const drawings = root.findAllByType(Image);
    expect(drawings).toHaveLength(2);
    expect(drawings[0].props.source.uri).toBe(`data:image/png;base64,${embeddedPng}`);
    expect(drawings[0].parent!.props.pointerEvents).toBe("none");
    expect(root.findAllByType(Pressable).some((cell) => cell.props.accessibilityLabel === "D5")).toBe(false);
  });

  it("clips a merged region at a later window while preserving its real anchor and selection", async () => {
    const sheet = (await readOfficeFile({ filePath: await saveStyledWorkbook(), rowStart: 50, columnStart: 12 })).sheets![0];
    const geometry = createSpreadsheetGeometry(sheet);
    await act(async () => {
      renderedCanvas = create(<SpreadsheetCanvas worksheet={sheet} geometry={geometry}
        selections={parseGridSelections("N52")} jumpToCoordinate={vi.fn()} />);
    });
    const anchor = renderedCanvas!.root.findAllByType(Pressable).find((cell) => cell.props.accessibilityLabel === "K49")!;
    expect(anchor.props.accessibilityLabel).toBe("K49");
    expect(anchor.props.style).toMatchObject({ left: 0, top: 0, width: 128, height: 40 });
    expect(anchor.findByType(Text).props.children).toBe("Tiếp tục vùng gộp");
    expect(anchor.findByProps({ testID: "office-selection-K49" })).toBeTruthy();
  });
});

describe("spreadsheet color compatibility", () => {
  it("resolves custom theme XML, HSL tints and indexed colors", () => {
    // A nonstandard theme makes hardcoded Office accents fail even if the default workbook happens to look right.
    const theme = readSpreadsheetTheme('<a:clrScheme><a:lt1><a:sysClr val="window" lastClr="EEEEEE"/></a:lt1>' +
      '<a:accent1><a:srgbClr val="123456"/></a:accent1></a:clrScheme>');
    expect(resolveSpreadsheetColor({ theme: 0 }, theme)).toBe("#EEEEEE");
    expect(resolveSpreadsheetColor({ theme: 4 }, theme)).toBe("#123456");
    expect(resolveSpreadsheetColor({ theme: 4, tint: 0.5 }, readSpreadsheetTheme())).toBe("#A7C0DE");
    expect(resolveSpreadsheetColor({ indexed: 2 }, theme)).toBe("#FF0000");
  });
});
