import React from "react";
import { act, create, type ReactTestRenderer } from "react-test-renderer";
import { afterEach, describe, expect, it, vi } from "vitest";
import { Pressable, View } from "react-native";
import { SpreadsheetColumnHeadings, SpreadsheetRowHeadings } from "../client/spreadsheet-headings";
import { SpreadsheetCanvas } from "../client/spreadsheet-canvas";
import { createSpreadsheetGeometry } from "../client/spreadsheet-geometry";
import { nextSpreadsheetWindow } from "../client/spreadsheet-scroll-window";
import { formatColumnLabel } from "../shared/office-store";
import type { WorksheetSheet } from "../shared/viewer";

// Render production headings and cells together, without substituting the geometry used by either component.
vi.mock("react-native", () => ({ View: "View", Text: "Text", Pressable: "Pressable", ScrollView: "ScrollView" }));
const palette = { background: "#111", surface: "#222", border: "#444", text: "#fff", textMuted: "#aaa", primary: "#38f" };
let preview: ReactTestRenderer | undefined;
afterEach(async () => { await act(async () => preview?.unmount()); preview = undefined; });

function hiddenColumnSheet(): WorksheetSheet {
  // Match the important pattern in BT2026: hidden A:C, a very wide G, and many hidden columns between narrow visible ones.
  const columnWidths = [0, 0, 0, 14, 35, 89, 441, 0, 77, 0, 45, 0, 0, 57, 0, 0,
    50, 0, 0, 40, 0, 0, 43, 0, 0, 42, 0, 0, 42, 0, 0, 42, 0, 0, 45, 0];
  return { name: "BT2026", rowStart: 0, columnStart: 0, rowCount: 1000, columnCount: 106,
    rows: Array.from({ length: 3 }, () => columnWidths.map(() => "Nội dung")),
    presentation: { columnWidths, rowHeights: [24, 0, 36] } };
}

describe("hidden worksheet coordinate alignment", () => {
  it("positions each visible label at exactly the cell boundary and omits hidden label borders", async () => {
    (globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
    const worksheet = hiddenColumnSheet();
    const geometry = createSpreadsheetGeometry(worksheet);
    const headingProps = { worksheet, geometry, selections: [], palette, jumpToCoordinate: vi.fn(),
      scrollbarInsets: { right: 15, bottom: 15 } };
    await act(async () => { preview = create(<View>
      <SpreadsheetColumnHeadings {...headingProps} />
      <SpreadsheetRowHeadings {...headingProps} />
      <SpreadsheetCanvas worksheet={worksheet} geometry={geometry} selections={[]} jumpToCoordinate={vi.fn()} />
    </View>); });
    const buttons = preview!.root.findAllByType(Pressable);
    for (const [columnIndex, columnWidth] of geometry.columnWidths.entries()) {
      const columnLabel = formatColumnLabel(columnIndex);
      const heading = buttons.find((button) => button.props.accessibilityLabel === `Cột ${columnLabel}`);
      if (columnWidth === 0) {
        // A zero-width label must not leave even its one-pixel border in the header layout.
        expect(heading).toBeUndefined();
        continue;
      }
      const cell = buttons.find((button) => button.props.accessibilityLabel === `${columnLabel}1`)!;
      expect(heading!.props.style).toMatchObject({ position: "absolute", left: cell.props.style.left, width: cell.props.style.width });
    }
    // Hidden row borders used to accumulate in the vertical gutter for the same reason.
    expect(buttons.some((button) => button.props.accessibilityLabel === "Dòng 2")).toBe(false);
    const thirdRow = buttons.find((button) => button.props.accessibilityLabel === "Dòng 3")!;
    const thirdRowCell = buttons.find((button) => button.props.accessibilityLabel === "D3")!;
    expect(thirdRow.props.style.top).toBe(thirdRowCell.props.style.top);
  });

  it("grows beyond 36 physical columns when hidden columns leave the viewport underfilled", () => {
    const worksheet = hiddenColumnSheet();
    const nextWindow = nextSpreadsheetWindow(worksheet, createSpreadsheetGeometry(worksheet), "columns", 0, 1600, 1)!;
    // Growing in place keeps every visible column while making the not-yet-readable right side reachable automatically.
    expect(nextWindow.columnStart).toBe(0);
    expect(nextWindow.columnLimit).toBeGreaterThan(36);
    expect(nextWindow.columnLimit).toBeLessThanOrEqual(106);
  });

  it("passes an entirely hidden initial prefix without requiring a click or scrollbar movement", () => {
    const worksheet = hiddenColumnSheet();
    worksheet.rows = worksheet.rows.map((row) => row.slice(0, 12));
    worksheet.presentation!.columnWidths = Array(12).fill(0);
    // There are no scrollable pixels yet; the next request must retain the prefix and discover later visible columns.
    const nextWindow = nextSpreadsheetWindow(worksheet, createSpreadsheetGeometry(worksheet), "columns", 0, 1600, 1)!;
    expect(nextWindow).toMatchObject({ columnStart: 0, columnLimit: 106 });
  });
});
