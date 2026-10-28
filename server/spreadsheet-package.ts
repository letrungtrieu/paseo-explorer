import { posix } from "node:path";
import JSZip from "jszip";
import { normalizeSpreadsheetDrawingXml } from "./spreadsheet-drawing-xml";

export async function normalizeSpreadsheetPackage(fileBuffer: Buffer): Promise<Buffer> {
  const archive = await JSZip.loadAsync(fileBuffer);
  let changedArchive = false;
  // ExcelJS assumes relative relationship targets, while openpyxl and other valid OOXML writers use package-root paths.
  for (const relationshipFile of archive.file(/\.rels$/)) {
    const relationshipsXml = await relationshipFile.async("string");
    const sourceDirectory = posix.dirname(posix.dirname(relationshipFile.name));
    const normalizedXml = relationshipsXml.replace(/<Relationship\b[^>]*\/?\s*>/g, (relationshipTag) => {
      // External links are not package files and must keep their original meaning; we never fetch them.
      if (/\bTargetMode=["']External["']/i.test(relationshipTag)) return relationshipTag;
      return relationshipTag.replace(/\bTarget=(["'])\/([^"']*)\1/, (_targetAttribute, quote: string, packagePath: string) =>
        `Target=${quote}${posix.relative(sourceDirectory === "." ? "" : sourceDirectory, packagePath)}${quote}`);
    });
    if (normalizedXml === relationshipsXml) continue;
    archive.file(relationshipFile.name, normalizedXml);
    changedArchive = true;
  }
  // Drawings from default-namespace writers need canonical names as well as compatible relationship paths.
  for (const drawingFile of archive.file(/^xl\/drawings\/[^/]+\.xml$/)) {
    const drawingXml = await drawingFile.async("string");
    if (/<xdr:wsDr[\s>]/.test(drawingXml)) continue;
    archive.file(drawingFile.name, normalizeSpreadsheetDrawingXml(drawingXml));
    changedArchive = true;
  }
  // Repack only incompatible inputs, and only in memory: the user's workbook remains byte-for-byte untouched.
  return changedArchive ? archive.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 1 } }) : fileBuffer;
}
