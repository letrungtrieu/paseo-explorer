import React, { useEffect, useState } from "react";
import { View, Text, ScrollView } from "react-native";
import { findIssuesForFilePath, officeStore } from "../shared/office-store";
import {
  COL_WIDTH,
  OfficeFilePanel,
  OfficeViewerPanel,
  type OfficeViewerPanelProps,
  ROW_HEIGHT,
} from "./office-viewer-panel";
import { resolveExplorerTheme, useResolvedWorkspacePath } from "./theme-colors";
import { WindowsTreeNode } from "./windows-tree-node";

export { COL_WIDTH, OfficeFilePanel, OfficeViewerPanel, type OfficeViewerPanelProps, ROW_HEIGHT, WindowsTreeNode };

export interface OfficeExplorerSidebarProps {
  theme: unknown;
  workspaceId?: string;
  cwd?: string;
}

// Right-sidebar Explorer panel: renders ONLY the Windows Explorer file tree; clicking a file opens it in a Main Tab.
export function OfficeExplorerSidebar({
  theme,
  workspaceId,
  cwd: explicitCwd,
}: OfficeExplorerSidebarProps) {
  const palette = resolveExplorerTheme(theme);
  const resolvedWorkspaceDirectory = useResolvedWorkspacePath(workspaceId, explicitCwd);
  const [storeSnapshot, setStoreSnapshot] = useState(() => officeStore.getState());

  useEffect(() => {
    return officeStore.subscribe((nextStoreState) => {
      setStoreSnapshot(nextStoreState);
    });
  }, []);

  const workspaceFolderLabel = resolvedWorkspaceDirectory?.split(/[\\/]/).pop() ?? "Workspace";

  return (
    <View style={{ flex: 1, backgroundColor: palette.background }}>
      <View
        style={{
          flexDirection: "row",
          justifyContent: "space-between",
          alignItems: "center",
          paddingHorizontal: 10,
          paddingVertical: 8,
          backgroundColor: palette.surface,
          borderBottomWidth: 1,
          borderColor: palette.border,
        }}
      >
        <Text style={{ color: palette.text, fontSize: 11, fontWeight: "bold" }}>
          📁 EXPLORER ({workspaceFolderLabel})
        </Text>
      </View>

      {resolvedWorkspaceDirectory ? (
        <ScrollView style={{ flex: 1, paddingVertical: 6 }}>
          <WindowsTreeNode
            dirPath={resolvedWorkspaceDirectory}
            depth={0}
            theme={theme}
            selectedPath={storeSnapshot.activeTarget?.filePath}
            issuesByFile={storeSnapshot.issuesByFile}
            onSelectFile={(selectedFilePath, selectedExt) => {
              const matchedIssues = findIssuesForFilePath(
                storeSnapshot.issuesByFile,
                selectedFilePath,
              );
              const firstReportedIssue = matchedIssues[0];
              if (firstReportedIssue) {
                officeStore.openAndJump(
                  {
                    filePath: selectedFilePath,
                    ext: selectedExt,
                    sheet: firstReportedIssue.sheet,
                    cell: firstReportedIssue.cell ?? "A1",
                    page: firstReportedIssue.page,
                    reason: firstReportedIssue.reason,
                  },
                  { workspaceId, openMainTab: true },
                );
              } else {
                officeStore.openAndJump(
                  {
                    filePath: selectedFilePath,
                    ext: selectedExt,
                    cell: "A1",
                  },
                  { workspaceId, openMainTab: true },
                );
              }
            }}
          />
        </ScrollView>
      ) : (
        <View style={{ padding: 12 }}>
          <Text style={{ color: palette.textMuted, fontSize: 12 }}>
            Đang tải thư mục Workspace...
          </Text>
        </View>
      )}
    </View>
  );
}
