import React, { forwardRef } from "react";
import { Pressable, ScrollView, Text, View } from "react-native";
import { formatColumnLabel, type ParsedGridSelection } from "../shared/office-store";
import type { WorksheetSheet } from "../shared/viewer";
import type { ExplorerThemePalette } from "./theme-colors";
import { COLUMN_HEADING_HEIGHT, ROW_HEADING_WIDTH, type SpreadsheetGeometry } from "./spreadsheet-geometry";
import { officeWarningColors as warning } from "./office-warning-colors";

interface SpreadsheetHeadingProps {
  worksheet: WorksheetSheet;
  geometry: SpreadsheetGeometry;
  selections: ParsedGridSelection[];
  palette: ExplorerThemePalette;
  jumpToCoordinate: (coordinate: string) => void;
  scrollbarInsets: { right: number; bottom: number };
}

function headingStyle(selected: boolean, palette: ExplorerThemePalette) {
  // App-themed headings remain distinct from workbook colors, including when the worksheet itself is white.
  return { justifyContent: "center" as const, alignItems: "center" as const,
    backgroundColor: selected ? warning.activeHeader : palette.surface,
    borderWidth: 0.5, borderColor: selected ? warning.border : palette.border };
}

export const SpreadsheetColumnHeadings = forwardRef<ScrollView, SpreadsheetHeadingProps>(
  function SpreadsheetColumnHeadings({ worksheet, geometry, selections, palette, jumpToCoordinate, scrollbarInsets }, ref) {
    // This gutter only follows data.scrollLeft; wheel/touch input belongs to the data viewport below it.
    return <ScrollView ref={ref} horizontal scrollEnabled={false} showsHorizontalScrollIndicator={false}
      testID="office-sheet-column-headings" style={{ flex: 1, minWidth: 0, height: COLUMN_HEADING_HEIGHT,
        marginRight: scrollbarInsets.right }}>
      <View style={{ position: "relative", width: geometry.width, height: COLUMN_HEADING_HEIGHT }}>
        {geometry.columnWidths.map((columnWidth, visibleColumn) => {
          // A zero-width button still has borders/minimum sizing in some hosts; hidden columns must draw nothing.
          if (columnWidth === 0) return null;
          const columnIndex = worksheet.columnStart + visibleColumn;
          const columnLabel = formatColumnLabel(columnIndex);
          // Headers use absolute workbook coordinates as buffered columns are added or removed.
          const selected = selections.some((selection) => selection.startCol !== undefined &&
            columnIndex >= selection.startCol && columnIndex <= selection.endCol!);
          return <Pressable key={columnIndex} accessibilityLabel={`Cột ${columnLabel}`}
            onPress={() => jumpToCoordinate(`Col:${columnLabel}`)}
            style={{ position: "absolute", left: geometry.columnOffsets[visibleColumn], width: columnWidth,
              height: COLUMN_HEADING_HEIGHT, overflow: "hidden", ...headingStyle(selected, palette) }}>
            <Text numberOfLines={1} style={{ color: selected ? warning.accent : palette.textMuted, fontSize: 10 }}>{columnLabel}</Text>
          </Pressable>;
        })}
      </View>
    </ScrollView>;
  },
);

export const SpreadsheetRowHeadings = forwardRef<ScrollView, SpreadsheetHeadingProps>(
  function SpreadsheetRowHeadings({ worksheet, geometry, selections, palette, jumpToCoordinate, scrollbarInsets }, ref) {
    // Keep row numbers inside a fixed-width clipped gutter while synchronizing only their vertical position.
    return <ScrollView ref={ref} scrollEnabled={false} showsVerticalScrollIndicator={false}
      testID="office-sheet-row-headings" style={{ width: ROW_HEADING_WIDTH, flexGrow: 0, flexShrink: 0, minHeight: 0,
        marginBottom: scrollbarInsets.bottom }}>
      <View style={{ position: "relative", width: ROW_HEADING_WIDTH, height: geometry.height }}>
        {geometry.rowHeights.map((rowHeight, visibleRow) => {
          // Hidden rows must not accumulate border pixels in the fixed row gutter either.
          if (rowHeight === 0) return null;
          const rowIndex = worksheet.rowStart + visibleRow;
          const selected = selections.some((selection) => selection.startRow !== undefined &&
            rowIndex >= selection.startRow && rowIndex <= selection.endRow!);
          // Saved heights, including hidden rows, align the gutter with merged cells and image anchors in the canvas.
          return <Pressable key={rowIndex} accessibilityLabel={`Dòng ${rowIndex + 1}`}
            onPress={() => jumpToCoordinate(`Row:${rowIndex + 1}`)}
            style={{ position: "absolute", top: geometry.rowOffsets[visibleRow], width: ROW_HEADING_WIDTH,
              height: rowHeight, overflow: "hidden", ...headingStyle(selected, palette) }}>
            <Text style={{ color: selected ? warning.accent : palette.textMuted, fontSize: 10 }}>{rowIndex + 1}</Text>
          </Pressable>;
        })}
      </View>
    </ScrollView>;
  },
);
