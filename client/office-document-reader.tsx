import React, { useEffect, useState } from "react";
import { useFileUrl, type PluginFile } from "@getpaseo/plugin/client";
import { DocumentFrame, Loading } from "@getpaseo/plugin/client/ui";
import { ScrollView, Text } from "react-native";
import type { OfficeFileData } from "../shared/viewer";
import type { ExplorerThemePalette } from "./theme-colors";
import { createOfficeDocumentUrl } from "./web";

export function PdfFileReader({ file, page, palette }: {
  file: PluginFile;
  page?: number;
  palette: ExplorerThemePalette;
}) {
  // Let Paseo own loading, remote file transfer, MIME type, watching, and URL disposal for native file tabs.
  const fileUrl = useFileUrl(file);
  if (fileUrl.status === "loading") return <Loading />;
  if (fileUrl.status === "error") {
    return <Text style={{ color: palette.text, padding: 16 }}>{fileUrl.error.message}</Text>;
  }
  // PDF page fragments retain the original page numbering, including blank and scanned pages.
  return <DocumentFrame url={`${fileUrl.url}${page ? `#page=${page}` : ""}`} testID="office-pdf" />;
}

export function OfficeDocumentReader({ document, page, palette }: {
  document: OfficeFileData;
  page?: number;
  palette: ExplorerThemePalette;
}) {
  const [previewUrl, setPreviewUrl] = useState<{ document: OfficeFileData; url: string } | null>(null);
  useEffect(() => {
    // Build only when document content changes; selecting an issue must not re-encode an entire PDF.
    const preview = createOfficeDocumentUrl(document);
    setPreviewUrl(preview ? { document, url: preview.url } : null);
    return () => preview?.dispose();
  }, [document]);

  if (previewUrl?.document === document) {
    return <DocumentFrame url={`${previewUrl.url}${document.fileType === "pdf" && page ? `#page=${page}` : ""}`}
      testID="office-document" />;
  }
  // Native Word retains a readable text view because browser blob URLs are unavailable on iOS and Android.
  return (
    <ScrollView style={{ flex: 1 }} contentContainerStyle={{ padding: 16 }}>
      <Text selectable style={{ color: palette.text, fontSize: 13, lineHeight: 21 }}>
        {document.textContent || (document.fileType === "pdf"
          ? "Mở PDF từ cây Files để xem trên thiết bị này." : "Tài liệu không có nội dung văn bản.")}
      </Text>
    </ScrollView>
  );
}
