import { useEffect, useLayoutEffect, useRef, useState } from "react";
import type { NativeScrollEvent, NativeSyntheticEvent, LayoutChangeEvent, ScrollView } from "react-native";
import type { OfficeTarget, ParsedGridSelection } from "../shared/office-store";
import type { WorksheetSheet } from "../shared/viewer";
import type { SpreadsheetWindow } from "../shared/spreadsheet-window";
import { visibleSpreadsheetIndex, type SpreadsheetGeometry } from "./spreadsheet-geometry";
import { nextSpreadsheetWindow, type SpreadsheetScrollAxis } from "./spreadsheet-scroll-window";
import { readSpreadsheetViewport } from "./web";

interface ScrollAnchor {
  row: number;
  column: number;
  rowOffset: number;
  columnOffset: number;
  requested: SpreadsheetWindow;
}

export function useSpreadsheetScroll(worksheet: WorksheetSheet, geometry: SpreadsheetGeometry,
  target: OfficeTarget, selection: ParsedGridSelection | undefined, loading: boolean,
  loadWindow: (window: SpreadsheetWindow) => void,
) {
  const verticalScroll = useRef<ScrollView>(null);
  const horizontalScroll = useRef<ScrollView>(null);
  const columnHeadings = useRef<ScrollView>(null);
  const rowHeadings = useRef<ScrollView>(null);
  const [scrollbarInsets, setScrollbarInsets] = useState({ right: 0, bottom: 0 });
  const position = useRef({ rows: 0, columns: 0 });
  const viewport = useRef({ rows: 0, columns: 0 });
  const pendingAnchor = useRef<ScrollAnchor | null>(null);
  const localSelection = useRef<string | null>(null);
  const acceptedTarget = useRef("");
  const targetKey = JSON.stringify([target.filePath, target.sheet, target.cell, target.timestamp]);

  useLayoutEffect(() => {
    // Gutters may have clamped their offset before scrollbar strips were excluded; realign before the next paint.
    rowHeadings.current?.scrollTo({ y: position.current.rows, animated: false });
    columnHeadings.current?.scrollTo({ x: position.current.columns, animated: false });
  }, [scrollbarInsets.right, scrollbarInsets.bottom]);

  const restorePosition = (verticalOffset: number, horizontalOffset: number) => {
    // Update before scrollTo: its native scroll event must not be mistaken for another user request.
    const clampedOffset = (offset: number, extent: number, viewportSize: number) => Math.max(0,
      viewportSize ? Math.min(offset, Math.max(0, extent - viewportSize)) : offset);
    position.current = { rows: clampedOffset(verticalOffset, geometry.height, viewport.current.rows),
      columns: clampedOffset(horizontalOffset, geometry.width, viewport.current.columns) };
    // Web shares one two-axis scroller; separate scrollTo calls would reset the other axis to zero there.
    if (verticalScroll.current === horizontalScroll.current) {
      verticalScroll.current?.scrollTo({ x: position.current.columns, y: position.current.rows, animated: false });
    } else {
      verticalScroll.current?.scrollTo({ y: position.current.rows, animated: false });
      horizontalScroll.current?.scrollTo({ x: position.current.columns, animated: false });
    }
    rowHeadings.current?.scrollTo({ y: position.current.rows, animated: false });
    columnHeadings.current?.scrollTo({ x: position.current.columns, animated: false });
  };

  useLayoutEffect(() => {
    if (acceptedTarget.current !== targetKey) {
      acceptedTarget.current = targetKey;
      // A direct cell click highlights in place; only an external jump should reposition the viewport.
      if (localSelection.current === target.cell) { localSelection.current = null; return; }
      pendingAnchor.current = null;
      const visibleRow = Math.min(Math.max(0, (selection?.startRow ?? worksheet.rowStart) - worksheet.rowStart - 2), geometry.rowHeights.length);
      const visibleColumn = Math.min(Math.max(0, (selection?.startCol ?? worksheet.columnStart) - worksheet.columnStart - 1), geometry.columnWidths.length);
      restorePosition(geometry.rowOffsets[visibleRow], geometry.columnOffsets[visibleColumn]);
      return;
    }
    const anchor = pendingAnchor.current;
    if (!anchor || anchor.requested.rowStart !== worksheet.rowStart ||
      anchor.requested.columnStart !== worksheet.columnStart ||
      geometry.rowHeights.length !== Math.min(anchor.requested.rowLimit, worksheet.rowCount - worksheet.rowStart) ||
      geometry.columnWidths.length !== Math.min(anchor.requested.columnLimit, worksheet.columnCount - worksheet.columnStart)) return;
    // Pixel offsets within the visible cell preserve position with custom heights, hidden rows, and prepended data.
    const visibleRow = anchor.row - worksheet.rowStart;
    const visibleColumn = anchor.column - worksheet.columnStart;
    restorePosition(geometry.rowOffsets[visibleRow] + anchor.rowOffset,
      geometry.columnOffsets[visibleColumn] + anchor.columnOffset);
    pendingAnchor.current = null;
  }, [worksheet, geometry, targetKey, loading]);

  const captureAnchor = (requested: SpreadsheetWindow): ScrollAnchor => {
    const firstRow = visibleSpreadsheetIndex(geometry.rowOffsets, position.current.rows);
    const firstColumn = visibleSpreadsheetIndex(geometry.columnOffsets, position.current.columns);
    // Retain an absolute cell anchor instead of assuming every preceding row is the default height.
    return { row: worksheet.rowStart + firstRow, column: worksheet.columnStart + firstColumn,
      rowOffset: position.current.rows - geometry.rowOffsets[firstRow],
      columnOffset: position.current.columns - geometry.columnOffsets[firstColumn], requested };
  };
  const requestMore = (axis: SpreadsheetScrollAxis, direction: number) => {
    if (loading || pendingAnchor.current || !viewport.current[axis]) return false;
    const requested = nextSpreadsheetWindow(worksheet, geometry, axis,
      position.current[axis], viewport.current[axis], direction);
    if (!requested) return false;
    pendingAnchor.current = captureAnchor(requested);
    loadWindow(requested);
    return true;
  };
  const measureDataViewport = () => {
    const measured = readSpreadsheetViewport(verticalScroll.current);
    if (measured) {
      // Scrollbar space varies by OS/theme; preserve alignment at the far right and bottom without hardcoded gutters.
      setScrollbarInsets((previous) => previous.right === measured.right && previous.bottom === measured.bottom
        ? previous : { right: measured.right, bottom: measured.bottom });
    }
    return measured;
  };
  const onScroll = (axis: SpreadsheetScrollAxis) => (event: NativeSyntheticEvent<NativeScrollEvent>) => {
    const scrollOffset = axis === "rows" ? event.nativeEvent.contentOffset.y : event.nativeEvent.contentOffset.x;
    const direction = Math.sign(scrollOffset - position.current[axis]);
    const measured = measureDataViewport();
    viewport.current[axis] = axis === "rows" ? measured?.height ?? event.nativeEvent.layoutMeasurement.height
      : measured?.width ?? event.nativeEvent.layoutMeasurement.width;
    position.current[axis] = scrollOffset;
    // Heading mirrors are clipped and noninteractive; they never change the data viewport's scroll ownership.
    if (axis === "rows") rowHeadings.current?.scrollTo({ y: scrollOffset, animated: false });
    else columnHeadings.current?.scrollTo({ x: scrollOffset, animated: false });
    if (pendingAnchor.current) {
      // Momentum may continue during a read; restore the latest visible cell rather than where loading started.
      const anchor = captureAnchor(pendingAnchor.current.requested);
      if (anchor.row >= anchor.requested.rowStart && anchor.row < anchor.requested.rowStart + anchor.requested.rowLimit &&
        anchor.column >= anchor.requested.columnStart && anchor.column < anchor.requested.columnStart + anchor.requested.columnLimit) {
        pendingAnchor.current = anchor;
      }
      return;
    }
    // No movement includes our programmatic anchor correction, which must not cascade into another fetch.
    if (direction !== 0) requestMore(axis, direction);
  };
  const checkViewportBuffers = () => {
    // Layout changes and chat jumps can reach a loaded edge without generating a user scroll event.
    for (const axis of ["rows", "columns"] as const) {
      if (viewport.current[axis] && !requestMore(axis, 1)) requestMore(axis, -1);
    }
  };
  useEffect(checkViewportBuffers, [worksheet, geometry, loading, targetKey]);
  const onLayout = (axis: SpreadsheetScrollAxis) => (event: LayoutChangeEvent) => {
    const measured = measureDataViewport();
    viewport.current[axis] = axis === "rows" ? measured?.height ?? event.nativeEvent.layout.height
      : measured?.width ?? event.nativeEvent.layout.width;
    checkViewportBuffers();
  };

  return { verticalScroll, horizontalScroll, columnHeadings, rowHeadings, scrollbarInsets,
    onVerticalScroll: onScroll("rows"), onHorizontalScroll: onScroll("columns"),
    onVerticalLayout: onLayout("rows"), onHorizontalLayout: onLayout("columns"),
    markLocalSelection: (coordinate: string) => { localSelection.current = coordinate; },
  };
}
