import React from "react";
import { Pressable, ScrollView, Text, View } from "react-native";
import { isMatchingOfficePath, officeStore, type OfficeTarget } from "../shared/office-store";
import type { WorksheetSheet } from "../shared/viewer";
import { WindowsFileIcon } from "./file-icons";
import type { ExplorerThemePalette } from "./theme-colors";

export function OfficeFileTabs({ openTargets, activeTarget, palette }: {
  openTargets: OfficeTarget[];
  activeTarget: OfficeTarget;
  palette: ExplorerThemePalette;
}) {
  // Horizontal tabs stay a single row on narrow panels, leaving room for the document itself.
  return (
    <ScrollView horizontal style={{ flexGrow: 0, flexShrink: 0, height: 38, minWidth: 0,
      backgroundColor: palette.surface }}>
      {openTargets.map((openTarget) => {
        const isActive = isMatchingOfficePath(openTarget.filePath, activeTarget.filePath);
        return (
          <View key={openTarget.filePath} style={{ flexDirection: "row", alignItems: "center",
            padding: 8, borderBottomWidth: 2, borderColor: isActive ? palette.primary : "transparent" }}>
            <Pressable onPress={() => officeStore.openAndJump(openTarget, { openMainTab: false })}
              style={{ flexDirection: "row", alignItems: "center" }}>
              <WindowsFileIcon isDirectory={false} ext={openTarget.ext} />
              <Text numberOfLines={1} style={{ color: palette.text, fontSize: 12 }}>{openTarget.filePath.split(/[\\/]/).pop()}</Text>
            </Pressable>
            {/* Closing uses the store so the fallback panel focuses the next remaining document. */}
            <Pressable onPress={() => officeStore.closeFileTab(openTarget.filePath)} style={{ paddingHorizontal: 8 }}>
              <Text style={{ color: palette.textMuted }}>×</Text>
            </Pressable>
          </View>
        );
      })}
    </ScrollView>
  );
}

export function WorksheetTabs({ sheets, activeSheet, palette, selectSheet }: {
  sheets: WorksheetSheet[];
  activeSheet?: string;
  palette: ExplorerThemePalette;
  selectSheet: (sheet: string) => void;
}) {
  // A fixed tab-strip height leaves the remaining viewport to the grid even when there are many sheet names.
  return (
    <ScrollView horizontal style={{ flexGrow: 0, flexShrink: 0, height: 34, minWidth: 0,
      backgroundColor: palette.surface }}>
      {sheets.map((sheet) => (
        <Pressable key={sheet.name} onPress={() => selectSheet(sheet.name)}
          style={{ paddingHorizontal: 12, paddingVertical: 8, borderBottomWidth: 2,
            borderColor: sheet.name === activeSheet ? palette.primary : "transparent" }}>
          <Text style={{ color: sheet.name === activeSheet ? palette.primary : palette.textMuted, fontSize: 11 }}>
            {sheet.name}
          </Text>
        </Pressable>
      ))}
    </ScrollView>
  );
}
