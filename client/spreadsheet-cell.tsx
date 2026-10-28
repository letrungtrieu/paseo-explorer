import React from "react";
import { Pressable, Text, View, type TextStyle, type ViewStyle } from "react-native";
import type { SpreadsheetCellStyle } from "../shared/spreadsheet-presentation";
import { officeWarningColors as warning } from "./office-warning-colors";

interface SpreadsheetCellProps {
  coordinate: string;
  text: string;
  left: number;
  top: number;
  width: number;
  height: number;
  selected: boolean;
  showGridLines: boolean;
  cellStyle?: SpreadsheetCellStyle;
  onPress: () => void;
}

function SavedCellBorders({ cellStyle }: { cellStyle?: SpreadsheetCellStyle }) {
  return <>{(["top", "right", "bottom", "left"] as const).map((side) => {
    const edge = cellStyle?.[side];
    if (!edge) return null;
    const horizontal = side === "top" || side === "bottom";
    // Separate edges preserve mixed borders (e.g. a dashed bottom and solid sides) on React Native too.
    const edgeStyle: ViewStyle = { position: "absolute", [side]: 0,
      ...(horizontal ? { left: 0, right: 0, borderTopWidth: edge.style === "double" ? 1 : edge.width }
        : { top: 0, bottom: 0, borderLeftWidth: edge.style === "double" ? 1 : edge.width }),
      borderColor: edge.color, borderStyle: edge.style === "double" ? "solid" : edge.style };
    return <React.Fragment key={side}>
      <View pointerEvents="none" style={edgeStyle} />
      {/* Two thin rules express a double edge without relying on web-only CSS border styles. */}
      {edge.style === "double" && <View pointerEvents="none" style={{ ...edgeStyle, [side]: 3 }} />}
    </React.Fragment>;
  })}</>;
}

export function SpreadsheetCell({ coordinate, text, left, top, width, height, selected,
  showGridLines, cellStyle, onPress }: SpreadsheetCellProps) {
  // Hidden rows/columns occupy no pixels and must not leave a clickable or visible sliver behind.
  if (width === 0 || height === 0) return null;
  const decoration = [cellStyle?.underline ? "underline" : "", cellStyle?.strike ? "line-through" : ""]
    .filter(Boolean).join(" ") as TextStyle["textDecorationLine"];
  const vertical = cellStyle?.vertical ?? "bottom";
  const containerStyle: ViewStyle = {
    position: "absolute", left, top, width, height, overflow: "hidden",
    justifyContent: vertical === "center" ? "center" : vertical === "bottom" ? "flex-end" : "flex-start",
    paddingHorizontal: 3, paddingLeft: 3 + (cellStyle?.indent ?? 0) * 9,
    backgroundColor: cellStyle?.background ?? "#FFFFFF",
    borderWidth: showGridLines ? 0.5 : 0, borderColor: "#D4D4D4",
  };
  // Workbook colors stay readable in a dark app: a white worksheet still needs its original black text.
  const textStyle: TextStyle = { fontFamily: cellStyle?.fontFamily ?? "Calibri",
    fontSize: cellStyle?.fontSize ?? 11 * 96 / 72, color: cellStyle?.color ?? "#000000",
    fontWeight: cellStyle?.bold ? "bold" : "normal", fontStyle: cellStyle?.italic ? "italic" : "normal",
    textAlign: cellStyle?.horizontal ?? "left", textDecorationLine: decoration || "none" };
  return (
    <Pressable accessibilityRole="button" accessibilityLabel={coordinate} onPress={onPress} style={containerStyle}>
      <Text numberOfLines={cellStyle?.wrap ? undefined : 1} style={textStyle}>{text}</Text>
      <SavedCellBorders cellStyle={cellStyle} />
      {/* Selection is an overlay so jumping from chat never destroys the saved font, fill, or borders. */}
      {selected && <View pointerEvents="none" testID={`office-selection-${coordinate}`} style={{
        position: "absolute", top: 0, right: 0, bottom: 0, left: 0,
        borderWidth: 2, borderColor: warning.accent,
      }} />}
    </Pressable>
  );
}
