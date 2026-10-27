import React, { useEffect, useRef, useState } from "react";
import { View, Text, Pressable, ScrollView } from "react-native";
import type { PluginFile, PluginFilePanelProps } from "@getpaseo/plugin/client";
import {
  findIssuesForFilePath,
  formatColumnLabel,
  isCellInSelections,
  isMatchingOfficePath,
  officeStore,
  type OfficeTarget,
  parseGridSelection,
  parseGridSelections,
} from "../shared/office-store";
import { readOfficeFileRpc } from "../shared/viewer";
import { WindowsFileIcon } from "./file-icons";
import { resolveExplorerTheme, useResolvedWorkspacePath, useRpcMutation } from "./theme-colors";

export const COL_WIDTH = 115;
export const ROW_HEIGHT = 32;

// Soft warning amber palette replaces harsh bright-red highlights so spreadsheet cells remain comfortable to read.
const SOFT_WARNING_ACCENT = "#f59e0b";
const SOFT_WARNING_BORDER = "rgba(245, 158, 11, 0.65)";
const SOFT_WARNING_INACTIVE_BORDER = "rgba(245, 158, 11, 0.30)";
const SOFT_WARNING_CELL_BG = "rgba(245, 158, 11, 0.14)";
const SOFT_WARNING_HEADER_ACTIVE_BG = "rgba(245, 158, 11, 0.24)";
const SOFT_WARNING_HEADER_SUBTLE_BG = "rgba(245, 158, 11, 0.10)";

export interface OfficeViewerPanelProps {
  theme: unknown;
  workspaceId?: string;
  lockedFilePath?: string;
  lockedExt?: string;
  watchedFile?: PluginFile;
}

// Build a stable key for each reported issue toast so individual toasts can be selected or dismissed independently.
function getIssueToastKey(issueTarget: Pick<OfficeTarget, "sheet" | "cell" | "page" | "reason">): string {
  const normalizedSheet = (issueTarget.sheet ?? "").trim().toLowerCase();
  const normalizedCoord =
    parseGridSelection(issueTarget.cell)?.label ?? (issueTarget.cell ?? "").trim().toUpperCase();
  return `${normalizedSheet}!${normalizedCoord}#${issueTarget.page ?? ""}#${issueTarget.reason ?? ""}`;
}

// Check whether two issue targets refer to the same worksheet coordinate so the active toast is highlighted.
function isSameIssueCoordinate(
  firstTarget: Pick<OfficeTarget, "sheet" | "cell" | "page">,
  secondTarget: Pick<OfficeTarget, "sheet" | "cell" | "page">,
  fallbackSheetName?: string,
): boolean {
  const firstSheet = (firstTarget.sheet ?? fallbackSheetName ?? "").trim().toLowerCase();
  const secondSheet = (secondTarget.sheet ?? fallbackSheetName ?? "").trim().toLowerCase();
  if (firstSheet && secondSheet && firstSheet !== secondSheet) {
    return false;
  }
  const firstCoord =
    parseGridSelection(firstTarget.cell)?.label ?? (firstTarget.cell ?? "").trim().toUpperCase();
  const secondCoord =
    parseGridSelection(secondTarget.cell)?.label ?? (secondTarget.cell ?? "").trim().toUpperCase();
  return firstCoord === secondCoord && (firstTarget.page ?? 0) === (secondTarget.page ?? 0);
}

// Main-pane tab viewer: renders the full-screen Excel/Word/PDF viewer with a right-side stack of individual issue toasts.
export function OfficeViewerPanel({
  theme,
  workspaceId,
  lockedFilePath,
  lockedExt,
  watchedFile,
}: OfficeViewerPanelProps) {
  const palette = resolveExplorerTheme(theme);
  const resolvedWorkspaceDirectory = useResolvedWorkspacePath(workspaceId);
  const readFileController = useRpcMutation(readOfficeFileRpc);

  const [storeSnapshot, setStoreSnapshot] = useState(() => officeStore.getState());
  const [activeSheetName, setActiveSheetName] = useState<string>("");
  const [dismissedToastKeys, setDismissedToastKeys] = useState<Record<string, boolean>>({});

  const verticalGridScrollRef = useRef<ScrollView>(null);
  const horizontalGridScrollRef = useRef<ScrollView>(null);

  useEffect(() => {
    return officeStore.subscribe((nextStoreState) => {
      setStoreSnapshot(nextStoreState);
    });
  }, []);

  // Match lockedFilePath against activeTarget even when activeTarget used a relative path from Chat.
  const lockedFileIssues = lockedFilePath
    ? findIssuesForFilePath(storeSnapshot.issuesByFile, lockedFilePath)
    : [];

  const activeTarget = lockedFilePath
    ? storeSnapshot.activeTarget &&
      isMatchingOfficePath(storeSnapshot.activeTarget.filePath, lockedFilePath)
      ? { ...storeSnapshot.activeTarget, filePath: lockedFilePath }
      : {
          filePath: lockedFilePath,
          ext: lockedExt ?? ".xlsx",
          sheet: lockedFileIssues[0]?.sheet,
          cell: lockedFileIssues[0]?.cell ?? "A1",
          page: lockedFileIssues[0]?.page,
          reason: lockedFileIssues[0]?.reason,
          timestamp: 0,
        }
    : storeSnapshot.activeTarget;

  // Whenever the user clicks a link in Chat, ensure that link's toast is visible again even if previously dismissed.
  useEffect(() => {
    if (!activeTarget) return;
    const clickedKey = getIssueToastKey(activeTarget);
    setDismissedToastKeys((previousDismissed) => {
      if (!previousDismissed[clickedKey]) return previousDismissed;
      const updatedDismissed = { ...previousDismissed };
      delete updatedDismissed[clickedKey];
      return updatedDismissed;
    });
  }, [activeTarget?.timestamp]);

  // Fetch parsed sheets/text from the server whenever the active file path or target sheet changes.
  useEffect(() => {
    if (!activeTarget?.filePath) return;
    if (activeTarget.sheet) {
      setActiveSheetName(activeTarget.sheet);
    }
    readFileController.mutate({
      filePath: activeTarget.filePath,
      cwd: resolvedWorkspaceDirectory ?? undefined,
    });
  }, [
    activeTarget?.filePath,
    activeTarget?.sheet,
    resolvedWorkspaceDirectory,
    readFileController.mutate,
  ]);

  // Reload file content automatically when Paseo detects an on-disk edit to a native File Tab.
  useEffect(() => {
    if (!watchedFile) return;
    return watchedFile.watch((fileEvent) => {
      if (fileEvent.type === "changed") {
        readFileController.mutate({
          filePath: watchedFile.path,
          cwd: resolvedWorkspaceDirectory ?? undefined,
        });
      }
    });
  }, [watchedFile, resolvedWorkspaceDirectory, readFileController.mutate]);

  const currentWorksheet =
    readFileController.data?.sheets?.find((sheetItem) => sheetItem.name === activeSheetName) ??
    readFileController.data?.sheets?.[0];

  // Only highlight selections when the currently viewed worksheet matches the clicked link's target worksheet.
  const isTargetOnActiveWorksheet =
    !activeTarget?.sheet ||
    !currentWorksheet ||
    activeTarget.sheet.trim().toLowerCase() === currentWorksheet.name.trim().toLowerCase();

  // Parse ONLY the cell/row/col/range coordinates described in the currently clicked link.
  const activeClickedSelections = isTargetOnActiveWorksheet
    ? parseGridSelections(activeTarget?.cell)
    : [];
  const primarySelection = activeClickedSelections[0] ?? null;

  // Auto-scroll vertically and/or horizontally to bring the clicked link's primary target into view.
  useEffect(() => {
    if (!primarySelection || !readFileController.data?.sheets) return;

    const scrollTimer = setTimeout(() => {
      if (primarySelection.startRow !== undefined) {
        verticalGridScrollRef.current?.scrollTo({
          y: Math.max(0, (primarySelection.startRow - 2) * ROW_HEIGHT),
          animated: true,
        });
      }
      if (primarySelection.startCol !== undefined) {
        horizontalGridScrollRef.current?.scrollTo({
          x: Math.max(0, (primarySelection.startCol - 1) * COL_WIDTH),
          animated: true,
        });
      }
    }, 120);

    return () => clearTimeout(scrollTimer);
  }, [primarySelection?.label, activeTarget?.timestamp, readFileController.data]);

  if (!activeTarget) {
    return (
      <View
        style={{
          flex: 1,
          backgroundColor: palette.background,
          justifyContent: "center",
          alignItems: "center",
          padding: 24,
        }}
      >
        <Text style={{ color: palette.textMuted, fontSize: 13 }}>
          Chọn một file Office (.xlsx, .docx, .pdf) từ cây thư mục bên phải hoặc bấm vào lỗi của Agent trong khung Chat để xem nội dung.
        </Text>
      </View>
    );
  }

  const columnHeaderCount = currentWorksheet?.rows?.[0]?.length ?? 0;
  const activeFileIssues = findIssuesForFilePath(
    storeSnapshot.issuesByFile,
    activeTarget.filePath,
  );

  // Build the list of individual toasts: each reported [[...]] issue is its own separate toast card.
  const allToastTargets: OfficeTarget[] =
    activeFileIssues.length > 0
      ? activeTarget.reason &&
        !activeFileIssues.some(
          (existingIssue) => getIssueToastKey(existingIssue) === getIssueToastKey(activeTarget),
        )
        ? [activeTarget, ...activeFileIssues]
        : activeFileIssues
      : activeTarget.reason
        ? [activeTarget]
        : [];

  const visibleToastTargets = allToastTargets.filter(
    (toastItem) => !dismissedToastKeys[getIssueToastKey(toastItem)],
  );

  return (
    <View style={{ flex: 1, backgroundColor: palette.background, position: "relative" }}>
      {/* Multi-file tab bar when opened via the Office Viewer workspace tab */}
      {!lockedFilePath && storeSnapshot.openTargets.length > 0 && (
        <View
          style={{
            flexDirection: "row",
            flexWrap: "wrap",
            backgroundColor: palette.surface,
            borderBottomWidth: 1,
            borderColor: palette.border,
          }}
        >
          {storeSnapshot.openTargets.map((openTabItem) => {
            const isTabActive = isMatchingOfficePath(
              openTabItem.filePath,
              activeTarget.filePath,
            );
            const tabFileName = openTabItem.filePath.split(/[\\/]/).pop() ?? openTabItem.filePath;
            return (
              <View
                key={openTabItem.filePath}
                style={{
                  flexDirection: "row",
                  alignItems: "center",
                  paddingHorizontal: 10,
                  paddingVertical: 6,
                  borderRightWidth: 1,
                  borderBottomWidth: 2,
                  borderRightColor: palette.border,
                  borderBottomColor: isTabActive ? "#107c41" : "transparent",
                  backgroundColor: isTabActive ? palette.background : palette.surface,
                }}
              >
                <Pressable
                  onPress={() =>
                    officeStore.openAndJump(openTabItem, { openMainTab: false })
                  }
                  style={{ flexDirection: "row", alignItems: "center" }}
                >
                  <WindowsFileIcon isDirectory={false} ext={openTabItem.ext} />
                  <Text
                    style={{
                      fontSize: 12,
                      color: palette.text,
                      fontWeight: isTabActive ? "bold" : "normal",
                    }}
                  >
                    {tabFileName}
                  </Text>
                </Pressable>
                <Pressable
                  onPress={() => officeStore.closeFileTab(openTabItem.filePath)}
                  style={{ marginLeft: 8, paddingHorizontal: 2 }}
                >
                  <Text style={{ fontSize: 12, color: palette.textMuted }}>×</Text>
                </Pressable>
              </View>
            );
          })}
        </View>
      )}

      {/* Worksheet tab selector for Excel workbooks */}
      {readFileController.data?.fileType === "xlsx" && readFileController.data.sheets && (
        <View
          style={{
            flexDirection: "row",
            flexWrap: "wrap",
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
                  paddingHorizontal: 12,
                  paddingVertical: 6,
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

      {/* Right-Side Toast Stack: each [[file#Sheet!Coord | Reason]] item renders as its own separate Toast card */}
      {visibleToastTargets.length > 0 && (
        <View
          style={{
            position: "absolute",
            top: 42,
            right: 16,
            zIndex: 30,
            width: 310,
            gap: 8,
          }}
        >
          {visibleToastTargets.map((toastIssue) => {
            const toastKey = getIssueToastKey(toastIssue);
            const formattedToastCoord =
              parseGridSelection(toastIssue.cell)?.label ?? toastIssue.cell ?? "A1";
            const toastSheetLabel =
              toastIssue.sheet ?? currentWorksheet?.name ?? "Sheet1";
            const isToastActive = isSameIssueCoordinate(
              toastIssue,
              activeTarget,
              currentWorksheet?.name,
            );

            return (
              <Pressable
                key={toastKey}
                onPress={() =>
                  officeStore.openAndJump(
                    {
                      filePath: activeTarget.filePath,
                      ext: activeTarget.ext,
                      sheet: toastIssue.sheet,
                      cell: toastIssue.cell,
                      page: toastIssue.page,
                      reason: toastIssue.reason,
                    },
                    { openMainTab: false },
                  )
                }
                style={{
                  paddingHorizontal: 11,
                  paddingVertical: 8,
                  borderRadius: 8,
                  backgroundColor: palette.surface,
                  borderWidth: 1,
                  borderColor: isToastActive
                    ? SOFT_WARNING_BORDER
                    : SOFT_WARNING_INACTIVE_BORDER,
                  borderLeftWidth: isToastActive ? 3.5 : 2,
                  borderLeftColor: isToastActive
                    ? SOFT_WARNING_ACCENT
                    : "rgba(245, 158, 11, 0.45)",
                }}
              >
                <View
                  style={{
                    flexDirection: "row",
                    justifyContent: "space-between",
                    alignItems: "center",
                    gap: 6,
                  }}
                >
                  <Text
                    style={{
                      flex: 1,
                      color: SOFT_WARNING_ACCENT,
                      fontWeight: isToastActive ? "bold" : "600",
                      fontSize: 11,
                    }}
                    numberOfLines={1}
                  >
                    ⚠{" "}
                    {toastIssue.page
                      ? `[Trang ${toastIssue.page}]`
                      : `[${toastSheetLabel}!${formattedToastCoord}]`}{" "}
                    <Text style={{ color: palette.textMuted, fontWeight: "normal", fontSize: 10 }}>
                      • {activeTarget.filePath.split(/[\\/]/).pop()}
                    </Text>
                  </Text>
                  <Pressable
                    onPress={(pressEvent) => {
                      pressEvent.stopPropagation?.();
                      setDismissedToastKeys((previousDismissed) => ({
                        ...previousDismissed,
                        [toastKey]: true,
                      }));
                    }}
                    style={{ paddingHorizontal: 4, paddingVertical: 1 }}
                  >
                    <Text style={{ color: palette.textMuted, fontSize: 12, fontWeight: "bold" }}>
                      ×
                    </Text>
                  </Pressable>
                </View>

                {toastIssue.reason && (
                  <Text
                    style={{
                      color: isToastActive ? palette.text : palette.textMuted,
                      fontSize: 11,
                      marginTop: 3,
                      lineHeight: 15,
                    }}
                  >
                    {toastIssue.reason}
                  </Text>
                )}
              </Pressable>
            );
          })}
        </View>
      )}

      {/* Full-pane 2D Excel grid highlighting ONLY the cells described in the currently clicked link */}
      {readFileController.data?.fileType === "xlsx" && currentWorksheet && (
        <ScrollView ref={verticalGridScrollRef} style={{ flex: 1 }}>
          <ScrollView ref={horizontalGridScrollRef} horizontal>
            <View>
              {/* Clickable Column Header Bar (A, B, C...) */}
              {columnHeaderCount > 0 && (
                <View style={{ flexDirection: "row", height: 26 }}>
                  <View
                    style={{
                      width: 40,
                      backgroundColor: palette.surface,
                      borderWidth: 0.5,
                      borderColor: palette.border,
                    }}
                  />
                  {Array.from({ length: columnHeaderCount }, (_, colIndex) => {
                    const colLetter = formatColumnLabel(colIndex);
                    const isWholeColSelected = activeClickedSelections.some(
                      (selectionItem) =>
                        selectionItem.kind === "col" &&
                        selectionItem.startCol !== undefined &&
                        selectionItem.endCol !== undefined &&
                        colIndex >= selectionItem.startCol &&
                        colIndex <= selectionItem.endCol,
                    );
                    const isColInSelection =
                      isWholeColSelected ||
                      activeClickedSelections.some(
                        (selectionItem) =>
                          selectionItem.startCol !== undefined &&
                          selectionItem.endCol !== undefined &&
                          colIndex >= selectionItem.startCol &&
                          colIndex <= selectionItem.endCol,
                      );

                    return (
                      <Pressable
                        key={colIndex}
                        onPress={() =>
                          officeStore.openAndJump(
                            {
                              filePath: activeTarget.filePath,
                              ext: activeTarget.ext,
                              sheet: currentWorksheet.name,
                              cell: `Col:${colLetter}`,
                            },
                            { openMainTab: false },
                          )
                        }
                        style={{
                          width: COL_WIDTH,
                          justifyContent: "center",
                          alignItems: "center",
                          backgroundColor: isWholeColSelected
                            ? SOFT_WARNING_HEADER_ACTIVE_BG
                            : isColInSelection
                              ? SOFT_WARNING_HEADER_SUBTLE_BG
                              : palette.surface,
                          borderWidth: 0.5,
                          borderColor: isColInSelection ? SOFT_WARNING_BORDER : palette.border,
                        }}
                      >
                        <Text
                          style={{
                            fontSize: 10,
                            fontWeight: "bold",
                            color: isColInSelection ? SOFT_WARNING_ACCENT : palette.textMuted,
                          }}
                        >
                          {colLetter}
                        </Text>
                      </Pressable>
                    );
                  })}
                </View>
              )}

              {/* Spreadsheet rows with clickable 1-based row headers */}
              {currentWorksheet.rows.map((rowCells, rowIndex) => {
                const isWholeRowSelected = activeClickedSelections.some(
                  (selectionItem) =>
                    selectionItem.kind === "row" &&
                    selectionItem.startRow !== undefined &&
                    selectionItem.endRow !== undefined &&
                    rowIndex >= selectionItem.startRow &&
                    rowIndex <= selectionItem.endRow,
                );
                const isRowInSelection =
                  isWholeRowSelected ||
                  activeClickedSelections.some(
                    (selectionItem) =>
                      selectionItem.startRow !== undefined &&
                      selectionItem.endRow !== undefined &&
                      rowIndex >= selectionItem.startRow &&
                      rowIndex <= selectionItem.endRow,
                  );

                return (
                  <View key={rowIndex} style={{ flexDirection: "row", height: ROW_HEIGHT }}>
                    <Pressable
                      onPress={() =>
                        officeStore.openAndJump(
                          {
                            filePath: activeTarget.filePath,
                            ext: activeTarget.ext,
                            sheet: currentWorksheet.name,
                            cell: `Row:${rowIndex + 1}`,
                          },
                          { openMainTab: false },
                        )
                      }
                      style={{
                        width: 40,
                        justifyContent: "center",
                        alignItems: "center",
                        backgroundColor: isWholeRowSelected
                          ? SOFT_WARNING_HEADER_ACTIVE_BG
                          : isRowInSelection
                            ? SOFT_WARNING_HEADER_SUBTLE_BG
                            : palette.surface,
                        borderWidth: 0.5,
                        borderColor: isRowInSelection ? SOFT_WARNING_BORDER : palette.border,
                      }}
                    >
                      <Text
                        style={{
                          fontSize: 10,
                          fontWeight: isRowInSelection ? "bold" : "normal",
                          color: isRowInSelection ? SOFT_WARNING_ACCENT : palette.textMuted,
                        }}
                      >
                        {rowIndex + 1}
                      </Text>
                    </Pressable>

                    {rowCells.map((cellTextValue, colIndex) => {
                      // Highlight ONLY if this cell belongs to the currently clicked link's selection(s).
                      const isTargetCell = isCellInSelections(
                        activeClickedSelections,
                        rowIndex,
                        colIndex,
                      );
                      const cellAddressLabel = `${formatColumnLabel(colIndex)}${rowIndex + 1}`;
                      const matchingClickedCellIssue = activeFileIssues.find((issueItem) => {
                        if (issueItem.sheet && issueItem.sheet !== currentWorksheet.name) {
                          return false;
                        }
                        return (
                          (issueItem.cell ?? "").toUpperCase() === cellAddressLabel.toUpperCase()
                        );
                      });

                      return (
                        <Pressable
                          key={colIndex}
                          onPress={() =>
                            officeStore.openAndJump(
                              {
                                filePath: activeTarget.filePath,
                                ext: activeTarget.ext,
                                sheet: currentWorksheet.name,
                                cell: cellAddressLabel,
                                reason: matchingClickedCellIssue?.reason,
                              },
                              { openMainTab: false },
                            )
                          }
                          style={{
                            width: COL_WIDTH,
                            height: ROW_HEIGHT,
                            justifyContent: "center",
                            paddingHorizontal: 6,
                            borderWidth: isTargetCell ? 1.5 : 0.5,
                            borderColor: isTargetCell ? SOFT_WARNING_BORDER : palette.border,
                            backgroundColor: isTargetCell ? SOFT_WARNING_CELL_BG : "transparent",
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
                        </Pressable>
                      );
                    })}
                  </View>
                );
              })}
            </View>
          </ScrollView>
        </ScrollView>
      )}

      {/* Full-pane Word (.docx/.doc) and Acrobat (.pdf) document reader */}
      {(readFileController.data?.fileType === "docx" ||
        readFileController.data?.fileType === "pdf") && (
        <ScrollView style={{ flex: 1, padding: 16 }}>
          <Text style={{ color: palette.text, fontSize: 12, lineHeight: 20 }}>
            {readFileController.data.textContent}
          </Text>
        </ScrollView>
      )}
    </View>
  );
}

// Native Paseo File Tab component (opened when clicking .xlsx/.docx/.pdf in Paseo's Files tab).
export function OfficeFilePanel({ file, workspaceId, theme }: PluginFilePanelProps) {
  const normalizedExtension = file.extension ? `.${file.extension.toLowerCase()}` : ".xlsx";
  return (
    <OfficeViewerPanel
      theme={theme}
      workspaceId={workspaceId}
      lockedFilePath={file.path}
      lockedExt={normalizedExtension}
      watchedFile={file}
    />
  );
}
