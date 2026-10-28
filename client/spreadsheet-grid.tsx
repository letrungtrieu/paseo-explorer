import React, { useMemo } from "react";
import { Text, View } from "react-native";
import { officeStore, parseGridSelections, type OfficeTarget } from "../shared/office-store";
import type { WorksheetSheet } from "../shared/viewer";
import type { SpreadsheetWindow } from "../shared/spreadsheet-window";
import type { ExplorerThemePalette } from "./theme-colors";
import { COLUMN_HEADING_HEIGHT, ROW_HEADING_WIDTH, createSpreadsheetGeometry } from "./spreadsheet-geometry";
import { SpreadsheetCanvas } from "./spreadsheet-canvas";
import { SpreadsheetColumnHeadings, SpreadsheetRowHeadings } from "./spreadsheet-headings";
import { SpreadsheetViewport } from "./spreadsheet-viewport";
import { useSpreadsheetScroll } from "./use-spreadsheet-scroll";

// Preserve existing consumers while geometry lives with the shared pixel-layout calculation.
export { COL_WIDTH, ROW_HEIGHT } from "./spreadsheet-geometry";

interface SpreadsheetGridProps {
  worksheet: WorksheetSheet;
  target: OfficeTarget;
  issues: OfficeTarget[];
  palette: ExplorerThemePalette;
  loadingWindow: boolean;
  loadWindow: (window: SpreadsheetWindow) => void;
}

export function SpreadsheetGrid({ worksheet, target, issues, palette, loadingWindow, loadWindow }: SpreadsheetGridProps) {
  const selections = !target.sheet || target.sheet.toLowerCase() === worksheet.name.toLowerCase()
    ? parseGridSelections(target.cell) : [];
  const geometry = useMemo(() => createSpreadsheetGeometry(worksheet), [worksheet]);
  const scrolling = useSpreadsheetScroll(worksheet, geometry, target, selections[0], loadingWindow, loadWindow);
  // Index reasons once per loaded window instead of searching every reported issue for every rendered cell.
  const issueReasons = useMemo(() => new Map(issues.filter(
    (issue) => !issue.sheet || issue.sheet.toLowerCase() === worksheet.name.toLowerCase(),
  ).map((issue) => [(issue.cell ?? "").toUpperCase(), issue.reason])), [issues, worksheet.name]);

  const jumpToCoordinate = (coordinate: string) => {
    // A local click highlights in place; incoming chat jumps retain their separate scroll-to-target behavior.
    scrolling.markLocalSelection(coordinate);
    officeStore.openAndJump({ filePath: target.filePath, ext: target.ext, sheet: worksheet.name,
      cell: coordinate, reason: issueReasons.get(coordinate.toUpperCase()) }, { openMainTab: false });
  };
  const headingProps = { worksheet, geometry, selections, palette, jumpToCoordinate, scrollbarInsets: scrolling.scrollbarInsets };

  return (
    // Flex minimums prevent the canvas's intrinsic size from expanding the workspace or its parent page.
    <View testID="office-sheet-viewport" style={{ flex: 1, minWidth: 0, minHeight: 0, overflow: "hidden" }}>
      {worksheet.rowCount === 0 && <Text style={{ color: palette.textMuted, padding: 16 }}>Sheet trống.</Text>}
      <View style={{ flexDirection: "row", height: COLUMN_HEADING_HEIGHT, flexShrink: 0, minWidth: 0 }}>
        <View style={{ width: ROW_HEADING_WIDTH, flexShrink: 0, backgroundColor: palette.surface,
          borderWidth: 0.5, borderColor: palette.border }} />
        {/* The header is outside the scrolling data area and follows only its horizontal offset. */}
        <SpreadsheetColumnHeadings ref={scrolling.columnHeadings} {...headingProps} />
      </View>
      <View style={{ flex: 1, flexDirection: "row", minWidth: 0, minHeight: 0, overflow: "hidden" }}>
        <SpreadsheetRowHeadings ref={scrolling.rowHeadings} {...headingProps} />
        {/* Only this bounded canvas viewport handles scrolling and near-edge data loading. */}
        <SpreadsheetViewport geometry={geometry} scrolling={scrolling}>
          <SpreadsheetCanvas worksheet={worksheet} geometry={geometry} selections={selections}
            jumpToCoordinate={jumpToCoordinate} />
        </SpreadsheetViewport>
      </View>
    </View>
  );
}
