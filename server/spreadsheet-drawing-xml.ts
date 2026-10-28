import { SaxesParser } from "saxes";

const DRAWING_NAMESPACES: Record<string, string> = {
  "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing": "xdr",
  "http://schemas.openxmlformats.org/drawingml/2006/main": "a",
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships": "r",
};

function escapeXmlText(xmlText: string): string {
  // SAX decodes entities, so serialized attribute/text values must be escaped again to remain valid XML.
  return xmlText.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

export function normalizeSpreadsheetDrawingXml(drawingXml: string): string {
  const parser = new SaxesParser({ xmlns: true });
  const xmlParts: string[] = [];
  let rootElement = true;
  const qualifiedName = (element: { uri: string; local: string; name: string }) =>
    DRAWING_NAMESPACES[element.uri] ? `${DRAWING_NAMESPACES[element.uri]}:${element.local}` : element.name;
  // ExcelJS matches literal xdr:/a: names; valid drawings using a default namespace otherwise fail during loading.
  parser.on("opentag", (element) => {
    const attributes = Object.values(element.attributes).filter((attribute) =>
      !(attribute.uri === "http://www.w3.org/2000/xmlns/" && DRAWING_NAMESPACES[attribute.value]));
    xmlParts.push(`<${qualifiedName(element)}`);
    if (rootElement) {
      // Declare canonical prefixes once while retaining unrelated namespace declarations from the original drawing.
      for (const [namespaceUri, prefix] of Object.entries(DRAWING_NAMESPACES)) {
        xmlParts.push(` xmlns:${prefix}="${namespaceUri}"`);
      }
      rootElement = false;
    }
    for (const attribute of attributes) {
      xmlParts.push(` ${qualifiedName(attribute)}="${escapeXmlText(attribute.value)}"`);
    }
    xmlParts.push(">");
  });
  // Fully qualified closing tags also handle SAX's close event for self-closing elements.
  parser.on("closetag", (element) => { xmlParts.push(`</${qualifiedName(element)}>`); });
  parser.on("text", (textContent) => { xmlParts.push(escapeXmlText(textContent)); });
  parser.on("cdata", (textContent) => { xmlParts.push(escapeXmlText(textContent)); });
  parser.write(drawingXml).close();
  return xmlParts.join("");
}
