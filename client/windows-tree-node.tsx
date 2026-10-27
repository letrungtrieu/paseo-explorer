import React, { useEffect, useState } from "react";
import { View, Text, Pressable } from "react-native";
import { listDirRpc, SUPPORTED_EXTS } from "../shared/explorer";
import {
  findIssuesForFilePath,
  isMatchingOfficePath,
  type OfficeTarget,
} from "../shared/office-store";
import { WindowsFileIcon } from "./file-icons";
import { resolveExplorerTheme, useRpcMutation } from "./theme-colors";

export interface WindowsTreeNodeProps {
  dirPath: string;
  depth: number;
  theme: unknown;
  selectedPath?: string;
  issuesByFile: Record<string, OfficeTarget[]>;
  onSelectFile: (filePath: string, ext: string) => void;
}

// Check whether a file extension belongs to the supported Office document formats.
function isSupportedOfficeExtension(fileExtension: string): boolean {
  const normalizedExt = fileExtension.toLowerCase();
  return (SUPPORTED_EXTS as readonly string[]).includes(normalizedExt);
}

// Recursive directory tree node with Windows Explorer vertical indent guides and issue count badges.
export function WindowsTreeNode({
  dirPath,
  depth,
  theme,
  selectedPath,
  issuesByFile,
  onSelectFile,
}: WindowsTreeNodeProps) {
  const palette = resolveExplorerTheme(theme);
  const listDirController = useRpcMutation(listDirRpc);
  const [expandedFoldersByPath, setExpandedFoldersByPath] = useState<Record<string, boolean>>({});

  // Fetch directory children lazily whenever a node mounts or its target directory path changes.
  useEffect(() => {
    listDirController.mutate({ dirPath });
  }, [dirPath, listDirController.mutate]);

  const directoryEntries = listDirController.data?.entries ?? [];

  return (
    <View
      style={{
        // Render a continuous vertical indent line beneath parent folders to mirror Windows Explorer's tree hierarchy.
        marginLeft: depth === 0 ? 0 : 11,
        borderLeftWidth: depth === 0 ? 0 : 1,
        borderColor: palette.border,
      }}
    >
      {directoryEntries.map((directoryEntry) => {
        const isFolderExpanded = Boolean(expandedFoldersByPath[directoryEntry.path]);
        const isSupportedDocument = isSupportedOfficeExtension(directoryEntry.ext);
        const isRowSelected = selectedPath
          ? isMatchingOfficePath(selectedPath, directoryEntry.path)
          : false;
        const reportedFileIssues = findIssuesForFilePath(issuesByFile, directoryEntry.path);

        return (
          <View key={directoryEntry.path}>
            <Pressable
              onPress={() => {
                if (directoryEntry.isDirectory) {
                  setExpandedFoldersByPath((previousExpandedState) => ({
                    ...previousExpandedState,
                    [directoryEntry.path]: !previousExpandedState[directoryEntry.path],
                  }));
                } else if (isSupportedDocument) {
                  onSelectFile(directoryEntry.path, directoryEntry.ext);
                }
              }}
              style={({ pressed }) => ({
                flexDirection: "row",
                alignItems: "center",
                paddingVertical: 4,
                paddingLeft: 6,
                paddingRight: 8,
                // Use Windows-style translucent blue highlight when selected, or surface hover tint while pressed.
                backgroundColor: isRowSelected
                  ? "rgba(59, 130, 246, 0.2)"
                  : pressed
                    ? palette.surface
                    : "transparent",
                // Dim non-Office files so users can visually scan spreadsheets, contracts, and PDFs faster.
                opacity: directoryEntry.isDirectory || isSupportedDocument ? 1 : 0.45,
              })}
            >
              <Text
                style={{
                  width: 14,
                  fontSize: 10,
                  color: palette.textMuted,
                  textAlign: "center",
                }}
              >
                {directoryEntry.isDirectory ? (isFolderExpanded ? "▾" : "▸") : ""}
              </Text>

              <WindowsFileIcon
                isDirectory={directoryEntry.isDirectory}
                isOpen={isFolderExpanded}
                ext={directoryEntry.ext}
              />

              <Text
                style={{
                  flex: 1,
                  fontSize: 12,
                  color: palette.text,
                  fontWeight: isRowSelected || directoryEntry.isDirectory ? "600" : "normal",
                }}
                numberOfLines={1}
              >
                {directoryEntry.name}
              </Text>

              {/* Show soft warning pill beside the filename whenever the Agent has registered issues on this file */}
              {reportedFileIssues.length > 0 && (
                <View
                  style={{
                    backgroundColor: "rgba(245, 158, 11, 0.18)",
                    borderWidth: 1,
                    borderColor: "rgba(245, 158, 11, 0.45)",
                    paddingHorizontal: 5,
                    paddingVertical: 1,
                    borderRadius: 8,
                    marginLeft: 4,
                  }}
                >
                  <Text style={{ color: "#f59e0b", fontSize: 9, fontWeight: "bold" }}>
                    ⚠ {reportedFileIssues.length}
                  </Text>
                </View>
              )}
            </Pressable>

            {directoryEntry.isDirectory && isFolderExpanded && (
              <WindowsTreeNode
                dirPath={directoryEntry.path}
                depth={depth + 1}
                theme={theme}
                selectedPath={selectedPath}
                issuesByFile={issuesByFile}
                onSelectFile={onSelectFile}
              />
            )}
          </View>
        );
      })}
    </View>
  );
}
