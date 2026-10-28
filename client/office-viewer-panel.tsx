import React, { useEffect, useState } from "react";
import { View, Text, Pressable } from "react-native";
import type { PluginFile, PluginFilePanelProps } from "@getpaseo/plugin/client";
import { Loading } from "@getpaseo/plugin/client/ui";
import {
  findIssuesForFilePath, isMatchingOfficePath, officeStore, type OfficeTarget,
} from "../shared/office-store";
import { resolveExplorerTheme, useResolvedWorkspacePath } from "./theme-colors";
import { useOfficeDocument } from "./use-office-document";
import { SpreadsheetGrid } from "./spreadsheet-grid";
import { OfficeDocumentReader, PdfFileReader } from "./office-document-reader";
import { OfficeIssueToasts } from "./office-issue-toasts";
import { OfficeFileTabs, WorksheetTabs } from "./office-viewer-tabs";

// Preserve the sidebar's public exports while keeping rendering geometry with the grid that owns it.
export { COL_WIDTH, ROW_HEIGHT } from "./spreadsheet-grid";

export interface OfficeViewerPanelProps {
  theme: unknown;
  workspaceId?: string;
  lockedFilePath?: string;
  lockedExt?: string;
  watchedFile?: PluginFile;
}

export function OfficeViewerPanel({
  theme, workspaceId, lockedFilePath, lockedExt, watchedFile,
}: OfficeViewerPanelProps) {
  const palette = resolveExplorerTheme(theme);
  const workspaceDirectory = useResolvedWorkspacePath(workspaceId);
  const [storeSnapshot, setStoreSnapshot] = useState(() => officeStore.getState());
  useEffect(() => officeStore.subscribe(setStoreSnapshot), []);

  const lockedIssues = findIssuesForFilePath(storeSnapshot.issuesByFile, lockedFilePath);
  // A native file tab stays pinned while chat may select another file in a neighboring tab.
  const activeTarget: OfficeTarget | null = lockedFilePath
    ? storeSnapshot.activeTarget && isMatchingOfficePath(storeSnapshot.activeTarget.filePath, lockedFilePath)
      ? { ...storeSnapshot.activeTarget, filePath: lockedFilePath }
      : { filePath: lockedFilePath, ext: lockedExt ?? ".xlsx", sheet: lockedIssues[0]?.sheet,
          cell: lockedIssues[0]?.cell ?? "A1", page: lockedIssues[0]?.page,
          reason: lockedIssues[0]?.reason, timestamp: 0 }
    : storeSnapshot.activeTarget;
  const { documentQuery, currentWorksheet, isNativePdf, navigate } = useOfficeDocument(
    activeTarget, workspaceDirectory, watchedFile,
  );

  if (!activeTarget) {
    return (
      <View style={{ flex: 1, minWidth: 0, minHeight: 0, overflow: "hidden",
        backgroundColor: palette.background, justifyContent: "center", padding: 24 }}>
        <Text style={{ color: palette.textMuted }}>
          Chọn file Excel, Word hoặc PDF từ cây thư mục, hoặc bấm link lỗi trong Chat để xem nội dung.
        </Text>
      </View>
    );
  }

  const activeFileIssues = findIssuesForFilePath(storeSnapshot.issuesByFile, activeTarget.filePath);
  const officeDocument = documentQuery.data;
  // Keep failures visible instead of leaving an empty pane or displaying the previous file's contents.
  const errorMessage = documentQuery.error instanceof Error
    ? documentQuery.error.message : documentQuery.error ? String(documentQuery.error) : null;

  return (
    // The panel owns the available workspace rectangle; sheet content must scroll within that rectangle.
    <View testID="office-viewer-panel" style={{ flex: 1, minWidth: 0, minHeight: 0,
      overflow: "hidden", backgroundColor: palette.background, position: "relative" }}>
      {!lockedFilePath && <OfficeFileTabs openTargets={storeSnapshot.openTargets}
        activeTarget={activeTarget} palette={palette} />}
      {officeDocument?.sheets && <WorksheetTabs sheets={officeDocument.sheets}
        activeSheet={currentWorksheet?.name} palette={palette} selectSheet={(sheet) => navigate(sheet)} />}
      {errorMessage && !isNativePdf && (
        <View style={{ padding: 16, gap: 12 }}>
          <Text selectable style={{ color: palette.text }}>Không thể mở file: {errorMessage}</Text>
          <Pressable accessibilityRole="button" onPress={() => { void documentQuery.refetch(); }}>
            <Text style={{ color: palette.primary }}>Thử lại</Text>
          </Pressable>
        </View>
      )}
      {/* A pending request owns a loading pane, so rapid navigation never paints a stale sheet as the new one. */}
      {!isNativePdf && !errorMessage && documentQuery.isPending && <Loading />}
      {isNativePdf && watchedFile && <PdfFileReader file={watchedFile} page={activeTarget.page} palette={palette} />}
      {/* A failed scroll prefetch keeps the last grid and its anchor visible beside the retry action. */}
      {officeDocument?.fileType === "xlsx" && currentWorksheet && (
        <SpreadsheetGrid key={`${officeDocument.filePath}:${currentWorksheet.name}`}
          worksheet={currentWorksheet} target={activeTarget} issues={activeFileIssues}
          palette={palette} loadingWindow={documentQuery.isFetching}
          loadWindow={(window) => navigate(currentWorksheet.name, window)} />
      )}
      {!errorMessage && officeDocument && officeDocument.fileType !== "xlsx" && (
        <OfficeDocumentReader document={officeDocument} page={activeTarget.page} palette={palette} />
      )}
      {/* Toasts remain independent from file reads: selecting or dismissing one should not reparse a workbook. */}
      <OfficeIssueToasts key={activeTarget.filePath} target={activeTarget} issues={activeFileIssues}
        sheetName={currentWorksheet?.name} palette={palette} />
    </View>
  );
}

export function OfficeFilePanel({ file, workspaceId, theme }: PluginFilePanelProps) {
  // Paseo supplies a stable file handle whose watcher is released by the document hook on unmount.
  return <OfficeViewerPanel theme={theme} workspaceId={workspaceId} lockedFilePath={file.path}
    lockedExt={`.${file.extension.toLowerCase()}`} watchedFile={file} />;
}
