import React, { useEffect, useRef, useState } from "react";
import { View, Text, Pressable, ScrollView } from "react-native";
import { formatColumnLabel, officeStore, parseCellRef } from "../shared/office-store";
import { readOfficeFileRpc } from "../shared/viewer";
import { resolveExplorerTheme, useResolvedWorkspacePath, useRpcMutation } from "./theme-colors";
import { WindowsTreeNode } from "./windows-tree-node";

export { WindowsTreeNode };

export const COL_WIDTH = 115;
export const ROW_HEIGHT = 32;

export interface OfficeExplorerSidebarProps {
  theme: unknown;
  workspaceId?: string;
  cwd?: string;
}

// Right-sidebar panel combining the Windows Explorer file tree with an auto-scrolling Office viewer.
export function OfficeExplorerSidebar({ theme, workspaceId, cwd: explicitCwd }: OfficeExplorerSidebarProps) {
  const palette = resolveExplorerTheme(theme);
  const resolvedWorkspaceDirectory = useResolvedWorkspacePath(workspaceId, explicitCwd);
  const readFileController = useRpcMutation(readOfficeFileRpc);

  const [storeSnapshot, setStoreSnapshot] = useState(() => officeStore.getState());
  const [activeSheetName, setActiveSheetName] = useState<string>("");
  const [isTreeCollapsed, setIsTreeCollapsed] = useState<boolean>(false);

  const verticalGridScrollRef = useRef<ScrollView>(null);
  const horizontalGridScrollRef = useRef<ScrollView>(null);

  const activeTarget = storeSnapshot.activeTarget;

  // Load any pre-existing target on mount and subscribe to subsequent chat/tree jump events.
  useEffect(() => {
    const initialTarget = officeStore.getState().activeTarget;
    if (initialTarget) {
      if (initialTarget.sheet) {
        setActiveSheetName(initialTarget.sheet);
      }
      readFileController.mutate({ filePath: initialTarget.filePath });
    }

    return officeStore.subscribe((nextStoreState) => {
      setStoreSnapshot(nextStoreState);
      if (nextStoreState.activeTarget) {
        if (nextStoreState.activeTarget.sheet) {
          setActiveSheetName(nextStoreState.activeTarget.sheet);
        }
        readFileController.mutate({ filePath: nextStoreState.activeTarget.filePath });
      }
    });
  }, [readFileController.mutate]);

  // Delay scroll slightly after data arrives so React Native ScrollViews have measured their content layout.
  useEffect(() => {
    if (!activeTarget?.cell || !readFileController.data?.sheets) return;
    const targetCoordinates = parseCellRef(activeTarget.cell);
    if (!targetCoordinates) return;

    const scrollTimer = setTimeout(() => {
      verticalGridScrollRef.current?.scrollTo({
        y: Math.max(0, (targetCoordinates.row - 2) * ROW_HEIGHT),
        animated: true,
      });
      horizontalGridScrollRef.current?.scrollTo({
        x: Math.max(0, (targetCoordinates.col - 1) * COL_WIDTH),
        animated: true,
      });
    }, 120);

    return () => clearTimeout(scrollTimer);
  }, [activeTarget, readFileController.data]);

  const currentWorksheet =
    readFileController.data?.sheets?.find((sheetItem) => sheetItem.name === activeSheetName) ??
    readFileController.data?.sheets?.[0];

  const highlightedCellCoords = parseCellRef(activeTarget?.cell);
  const columnHeaderCount = currentWorksheet?.rows?.[0]?.length ?? 0;
  const workspaceFolderLabel = resolvedWorkspaceDirectory?.split(/[\\/]/).pop() ?? "Workspace";

  return (
    <View style={{ flex: 1, backgroundColor: palette.background }}>
      {/* Top header with workspace title and quick toggle to collapse/expand the directory tree */}
      <View
        style={{
          flexDirection: "row",
          justifyContent: "space-between",
          alignItems: "center",
          paddingHorizontal: 10,
          paddingVertical: 6,
          backgroundColor: palette.surface,
          borderBottomWidth: 1,
          borderColor: palette.border,
        }}
      >
        <Text style={{ color: palette.text, fontSize: 11, fontWeight: "bold" }}>
          EXPLORER ({workspaceFolderLabel})
        </Text>
        {activeTarget && (
          <Pressable onPress={() => setIsTreeCollapsed(!isTreeCollapsed)}>
            <Text style={{ color: palette.primary, fontSize: 11 }}>
              {isTreeCollapsed ? "Hiện cây thư mục ▾" : "Thu gọn cây ▴"}
            </Text>
          </Pressable>
        )}
      </View>

      {/* Quick-switch tab bar matching the Office Explorer header design when a file is active */}
      {activeTarget && (
        <View
          style={{
            flexDirection: "row",
            alignItems: "center",
            gap: 6,
            paddingHorizontal: 8,
            paddingVertical: 4,
            backgroundColor: palette.surface,
            borderBottomWidth: 1,
            borderColor: palette.border,
          }}
        >
          <Pressable
            onPress={() => setIsTreeCollapsed(false)}
            style={{
              paddingHorizontal: 8,
              paddingVertical: 3,
              borderRadius: 4,
              backgroundColor: !isTreeCollapsed ? "rgba(59, 130, 246, 0.18)" : "transparent",
            }}
          >
            <Text
              style={{
                fontSize: 11,
                fontWeight: !isTreeCollapsed ? "600" : "normal",
                color: !isTreeCollapsed ? palette.primary : palette.textMuted,
              }}
            >
              📁 Cây thư mục
            </Text>
          </Pressable>

          <Pressable
            onPress={() => setIsTreeCollapsed(true)}
            style={{
              flexShrink: 1,
              paddingHorizontal: 8,
              paddingVertical: 3,
              borderRadius: 4,
              backgroundColor: isTreeCollapsed ? "rgba(16, 124, 65, 0.2)" : "transparent",
            }}
          >
            <Text
              style={{
                fontSize: 11,
                fontWeight: isTreeCollapsed ? "600" : "normal",
                color: isTreeCollapsed ? "#107c41" : palette.text,
              }}
              numberOfLines={1}
            >
              📊 {activeTarget.filePath.split(/[\\/]/).pop()}
              {activeTarget.cell
                ? ` (${currentWorksheet?.name ?? activeTarget.sheet ?? "Sheet1"}!${activeTarget.cell})`
                : ""}
            </Text>
          </Pressable>
        </View>
      )}

      {/* Upper pane: Windows Explorer file tree (collapses to 38% height when a document is open below) */}
      {!isTreeCollapsed && resolvedWorkspaceDirectory && (
        <ScrollView
          style={{
            maxHeight: activeTarget ? "38%" : "100%",
            borderBottomWidth: activeTarget ? 1 : 0,
            borderColor: palette.border,
            paddingVertical: 4,
          }}
        >
          <WindowsTreeNode
            dirPath={resolvedWorkspaceDirectory}
            depth={0}
            theme={theme}
            selectedPath={activeTarget?.filePath}
            issuesByFile={storeSnapshot.issuesByFile}
            onSelectFile={(selectedFilePath, selectedExt) => {
              // When clicking a file in the tree, jump directly to its first reported issue if one exists.
              const firstFileIssue = storeSnapshot.issuesByFile[selectedFilePath]?.[0];
              if (firstFileIssue) {
                officeStore.openAndJump({
                  filePath: selectedFilePath,
                  ext: selectedExt,
                  sheet: firstFileIssue.sheet,
                  cell: firstFileIssue.cell ?? "A1",
                  page: firstFileIssue.page,
                  reason: firstFileIssue.reason,
                });
              } else {
                officeStore.openAndJump({
                  filePath: selectedFilePath,
                  ext: selectedExt,
                  cell: "A1",
                });
              }
            }}
          />
        </ScrollView>
      )}

      {/* Lower pane: Integrated Office Document Viewer with cell-level error highlighting */}
      {activeTarget && (
        <View style={{ flex: 1 }}>
          {/* Active document banner showing coordinates and the Agent's diagnostic message */}
          <View
            style={{
              padding: 8,
              backgroundColor: activeTarget.reason ? "#fef2f2" : palette.surface,
              borderBottomWidth: 1,
              borderColor: activeTarget.reason ? "#ef4444" : palette.border,
            }}
          >
            <Text
              style={{
                color: activeTarget.reason ? "#7f1d1d" : palette.text,
                fontWeight: "bold",
                fontSize: 12,
              }}
            >
              {activeTarget.filePath.split(/[\\/]/).pop()}{" "}
              {activeTarget.cell
                ? `• [${currentWorksheet?.name ?? activeTarget.sheet ?? ""}!${activeTarget.cell}]`
                : activeTarget.page
                  ? `• [Trang ${activeTarget.page}]`
                  : ""}
            </Text>
            {activeTarget.reason && (
              <Text style={{ color: "#dc2626", fontSize: 11, marginTop: 2 }}>
                ⚠ Agent: {activeTarget.reason}
              </Text>
            )}
          </View>

          {/* Worksheet tab bar for multi-sheet Excel workbooks */}
          {readFileController.data?.fileType === "xlsx" && readFileController.data.sheets && (
            <View
              style={{
                flexDirection: "row",
                backgroundColor: palette.surface,
                borderBottomWidth: 1,
                borderColor: palette.border,
              }}
            >
              {readFileController.data.sheets.map((worksheetItem) => {
                const isSheetSelected = worksheetItem.name === currentWorksheet?.name;
                return (
                  <Pressable
                    key={worksheetItem.name}
                    onPress={() => setActiveSheetName(worksheetItem.name)}
                    style={{
                      paddingHorizontal: 10,
                      paddingVertical: 5,
                      borderBottomWidth: 2,
                      borderColor: isSheetSelected ? "#107c41" : "transparent",
                    }}
                  >
                    <Text
                      style={{
                        fontSize: 11,
                        fontWeight: isSheetSelected ? "bold" : "normal",
                        color: isSheetSelected ? "#107c41" : palette.textMuted,
                      }}
                    >
                      {worksheetItem.name}
                    </Text>
                  </Pressable>
                );
              })}
            </View>
          )}

          {/* 2D Excel grid with column letter headers (A, B, C...) and row numbers (1, 2, 3...) */}
          {readFileController.data?.fileType === "xlsx" && currentWorksheet && (
            <ScrollView ref={verticalGridScrollRef} style={{ flex: 1 }}>
              <ScrollView ref={horizontalGridScrollRef} horizontal>
                <View>
                  {/* Column header bar (A, B, C...) aligned with data cells below */}
                  {columnHeaderCount > 0 && (
                    <View style={{ flexDirection: "row", height: 24 }}>
                      <View
                        style={{
                          width: 36,
                          backgroundColor: palette.surface,
                          borderWidth: 0.5,
                          borderColor: palette.border,
                        }}
                      />
                      {Array.from({ length: columnHeaderCount }, (_, colIndex) => {
                        const isTargetCol = highlightedCellCoords?.col === colIndex;
                        return (
                          <View
                            key={colIndex}
                            style={{
                              width: COL_WIDTH,
                              justifyContent: "center",
                              alignItems: "center",
                              backgroundColor: isTargetCol ? "rgba(239, 68, 68, 0.12)" : palette.surface,
                              borderWidth: 0.5,
                              borderColor: palette.border,
                            }}
                          >
                            <Text
                              style={{
                                fontSize: 10,
                                fontWeight: "bold",
                                color: isTargetCol ? "#ef4444" : palette.textMuted,
                              }}
                            >
                              {formatColumnLabel(colIndex)}
                            </Text>
                          </View>
                        );
                      })}
                    </View>
                  )}

                  {/* Spreadsheet rows with 1-based row numbers and red-framed error cells */}
                  {currentWorksheet.rows.map((rowCells, rowIndex) => (
                    <View key={rowIndex} style={{ flexDirection: "row", height: ROW_HEIGHT }}>
                      <View
                        style={{
                          width: 36,
                          justifyContent: "center",
                          alignItems: "center",
                          backgroundColor: palette.surface,
                          borderWidth: 0.5,
                          borderColor: palette.border,
                        }}
                      >
                        <Text style={{ fontSize: 10, color: palette.textMuted }}>
                          {rowIndex + 1}
                        </Text>
                      </View>

                      {rowCells.map((cellTextValue, colIndex) => {
                        const isTargetCell =
                          highlightedCellCoords?.row === rowIndex &&
                          highlightedCellCoords?.col === colIndex;

                        return (
                          <View
                            key={colIndex}
                            style={{
                              width: COL_WIDTH,
                              height: ROW_HEIGHT,
                              justifyContent: "center",
                              paddingHorizontal: 6,
                              borderWidth: isTargetCell ? 2 : 0.5,
                              borderColor: isTargetCell ? "#ef4444" : palette.border,
                              backgroundColor: isTargetCell
                                ? "rgba(239, 68, 68, 0.18)"
                                : "transparent",
                            }}
                          >
                            <Text
                              style={{
                                fontSize: 11,
                                color: palette.text,
                                fontWeight: isTargetCell ? "bold" : "normal",
                              }}
                              numberOfLines={1}
                            >
                              {cellTextValue}
                            </Text>
                          </View>
                        );
                      })}
                    </View>
                  ))}
                </View>
              </ScrollView>
            </ScrollView>
          )}

          {/* Document reader for Word (.docx/.doc) and Acrobat (.pdf) files */}
          {(readFileController.data?.fileType === "docx" ||
            readFileController.data?.fileType === "pdf") && (
            <ScrollView style={{ flex: 1, padding: 12 }}>
              <Text style={{ color: palette.text, fontSize: 12, lineHeight: 20 }}>
                {readFileController.data.textContent}
              </Text>
            </ScrollView>
          )}
        </View>
      )}
    </View>
  );
}
