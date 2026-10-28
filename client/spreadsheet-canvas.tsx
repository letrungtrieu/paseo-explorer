import React, { useMemo } from "react";
import { Image, View } from "react-native";
import { formatColumnLabel, isCellInSelections, type ParsedGridSelection } from "../shared/office-store";
import type { WorksheetSheet } from "../shared/viewer";
import { SpreadsheetCell } from "./spreadsheet-cell";
import type { SpreadsheetGeometry } from "./spreadsheet-geometry";

interface SpreadsheetCanvasProps {
  worksheet: WorksheetSheet;
  geometry: SpreadsheetGeometry;
  selections: ParsedGridSelection[];
  jumpToCoordinate: (coordinate: string) => void;
}

export function SpreadsheetCanvas({ worksheet, geometry, selections, jumpToCoordinate }: SpreadsheetCanvasProps) {
  const presentation = worksheet.presentation;
  // A window contains all intersecting merges, so covered cells can be suppressed without loading neighboring pages.
  const merges = presentation?.merges ?? [];
  const coveredCells = useMemo(() => {
    const covered = new Set<number>();
    const columnCount = geometry.columnWidths.length;
    // Index only this clipped window; large merged ranges must not cause full-sheet iteration or per-cell merge scans.
    for (const merge of merges) {
      const firstRow = Math.max(0, merge.startRow - worksheet.rowStart);
      const lastRow = Math.min(geometry.rowHeights.length - 1, merge.endRow - worksheet.rowStart);
      const firstColumn = Math.max(0, merge.startColumn - worksheet.columnStart);
      const lastColumn = Math.min(columnCount - 1, merge.endColumn - worksheet.columnStart);
      for (let visibleRow = firstRow; visibleRow <= lastRow; visibleRow++) {
        for (let visibleColumn = firstColumn; visibleColumn <= lastColumn; visibleColumn++) {
          covered.add(visibleRow * columnCount + visibleColumn);
        }
      }
    }
    return covered;
  }, [merges, worksheet.rowStart, worksheet.columnStart, geometry]);
  const commonCellProps = { showGridLines: presentation?.showGridLines !== false };
  return <View testID="office-spreadsheet-canvas" style={{
    width: geometry.width, height: geometry.height, position: "relative", overflow: "hidden", backgroundColor: "#FFFFFF",
  }}>
    {worksheet.rows.map((rowCells, visibleRow) => rowCells.map((cellText, visibleColumn) => {
      // Hidden coordinates still exist in the workbook, but need neither a React cell nor any layout pixels.
      if (geometry.rowHeights[visibleRow] === 0 || geometry.columnWidths[visibleColumn] === 0) return null;
      const rowIndex = worksheet.rowStart + visibleRow;
      const columnIndex = worksheet.columnStart + visibleColumn;
      // Covered cells must disappear rather than drawing gridlines and duplicate text through a merged region.
      if (coveredCells.has(visibleRow * geometry.columnWidths.length + visibleColumn)) return null;
      const coordinate = `${formatColumnLabel(columnIndex)}${rowIndex + 1}`;
      // Resolve a shared style at render time rather than copying identical style objects onto every cell.
      const styleId = presentation?.cellStyleIds?.[visibleRow]?.[visibleColumn] ?? 0;
      return <SpreadsheetCell key={coordinate} coordinate={coordinate} text={cellText}
        left={geometry.columnOffsets[visibleColumn]} top={geometry.rowOffsets[visibleRow]}
        width={geometry.columnWidths[visibleColumn]} height={geometry.rowHeights[visibleRow]}
        cellStyle={presentation?.styles?.[styleId]} selected={isCellInSelections(selections, rowIndex, columnIndex)}
        onPress={() => jumpToCoordinate(coordinate)} {...commonCellProps} />;
    }))}
    {merges.map((merge) => {
      // Only the visible portion of a merge belongs to this canvas; its original anchor still owns navigation.
      const firstRow = Math.max(merge.startRow - worksheet.rowStart, 0);
      const lastRow = Math.min(merge.endRow - worksheet.rowStart + 1, geometry.rowHeights.length);
      const firstColumn = Math.max(merge.startColumn - worksheet.columnStart, 0);
      const lastColumn = Math.min(merge.endColumn - worksheet.columnStart + 1, geometry.columnWidths.length);
      const coordinate = `${formatColumnLabel(merge.startColumn)}${merge.startRow + 1}`;
      // A jump to any covered cell selects the merged region, including an anchor outside this preview window.
      const selected = selections.some((selection) => (selection.startRow === undefined ||
        (selection.startRow <= merge.endRow && selection.endRow! >= merge.startRow)) &&
        (selection.startCol === undefined || (selection.startCol <= merge.endColumn && selection.endCol! >= merge.startColumn)));
      return <SpreadsheetCell key={coordinate} coordinate={coordinate} text={merge.text}
        left={geometry.columnOffsets[firstColumn]} top={geometry.rowOffsets[firstRow]}
        width={geometry.columnOffsets[lastColumn] - geometry.columnOffsets[firstColumn]}
        height={geometry.rowOffsets[lastRow] - geometry.rowOffsets[firstRow]}
        cellStyle={presentation?.styles?.[merge.styleId]} selected={selected}
        onPress={() => jumpToCoordinate(coordinate)} {...commonCellProps} />;
    })}
    {/* Drawing coordinates share the cells' pixel geometry; clipping also preserves images crossing preview pages. */}
    {presentation?.images?.map((image, imageIndex) => <View key={imageIndex} pointerEvents="none"
      style={{ position: "absolute", left: image.left, top: image.top, width: image.width, height: image.height }}>
      <Image accessibilityLabel={`Ảnh nhúng ${imageIndex + 1}`} resizeMode="stretch"
        source={{ uri: image.uri }} style={{ width: image.width, height: image.height }} />
    </View>)}
  </View>;
}
