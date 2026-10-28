import React, { useEffect, useState } from "react";
import { Pressable, ScrollView, Text, View } from "react-native";
import { officeStore, parseGridSelection, type OfficeTarget } from "../shared/office-store";
import type { ExplorerThemePalette } from "./theme-colors";
import { officeWarningColors as warning } from "./office-warning-colors";

function issueKey(issue: Pick<OfficeTarget, "sheet" | "cell" | "page" | "reason">): string {
  // Normalize coordinates so dismissing an issue also hides equivalent Vietnamese/English row references.
  const coordinate = parseGridSelection(issue.cell)?.label ?? issue.cell?.trim().toUpperCase() ?? "";
  return `${issue.sheet?.trim().toLowerCase() ?? ""}!${coordinate}#${issue.page ?? ""}#${issue.reason ?? ""}`;
}

export function OfficeIssueToasts({ target, issues, sheetName, palette }: {
  target: OfficeTarget;
  issues: OfficeTarget[];
  sheetName?: string;
  palette: ExplorerThemePalette;
}) {
  const [dismissedIssues, setDismissedIssues] = useState<Record<string, boolean>>({});
  const clickedIssueKey = issueKey(target);
  useEffect(() => {
    // A chat click deliberately reopens its toast even if the user dismissed it earlier.
    setDismissedIssues((previousDismissed) => {
      if (!previousDismissed[clickedIssueKey]) return previousDismissed;
      const nextDismissed = { ...previousDismissed };
      delete nextDismissed[clickedIssueKey];
      return nextDismissed;
    });
  }, [clickedIssueKey, target.timestamp]);

  const toastTargets = target.reason && !issues.some((issue) => issueKey(issue) === clickedIssueKey)
    ? [target, ...issues] : issues;
  const visibleIssues = toastTargets.filter((issue) => !dismissedIssues[issueKey(issue)]);
  if (!visibleIssues.length) return null;

  // Constrain the overlay to the panel and let long issue lists scroll instead of covering every worksheet row.
  return (
    <View pointerEvents="box-none" style={{ position: "absolute", top: 42, right: 12,
      zIndex: 30, maxHeight: "65%", maxWidth: "90%", width: 310 }}>
      <ScrollView contentContainerStyle={{ gap: 8 }}>
        {visibleIssues.map((issue) => {
          const toastKey = issueKey(issue);
          const coordinate = parseGridSelection(issue.cell)?.label ?? issue.cell ?? "A1";
          // Reasons are excluded from the active-coordinate check so duplicate reports still focus together.
          const isActive = (issue.sheet ?? sheetName)?.toLowerCase() === (target.sheet ?? sheetName)?.toLowerCase()
            && (parseGridSelection(issue.cell)?.label ?? issue.cell ?? "") ===
              (parseGridSelection(target.cell)?.label ?? target.cell ?? "")
            && (issue.page ?? 0) === (target.page ?? 0);
          return (
            <Pressable key={toastKey} onPress={() => officeStore.openAndJump({ ...issue,
              filePath: target.filePath, ext: target.ext }, { openMainTab: false })}
              style={{ padding: 10, borderRadius: 8, backgroundColor: palette.surface,
                borderWidth: 1, borderColor: isActive ? warning.border : warning.inactiveBorder,
                borderLeftWidth: isActive ? 3 : 2 }}>
              <View style={{ flexDirection: "row", alignItems: "center", gap: 6 }}>
                <Text numberOfLines={1} style={{ flex: 1, color: warning.accent, fontSize: 11,
                  fontWeight: isActive ? "bold" : "600" }}>
                  ⚠ {issue.page ? `[Trang ${issue.page}]` : `[${issue.sheet ?? sheetName ?? "Sheet1"}!${coordinate}]`}
                </Text>
                <Pressable accessibilityLabel="Đóng thông báo" onPress={(pressEvent) => {
                  // Closing a toast must not bubble into the parent and immediately reopen its target.
                  pressEvent.stopPropagation();
                  setDismissedIssues((previousDismissed) => ({ ...previousDismissed, [toastKey]: true }));
                }} style={{ paddingHorizontal: 4 }}>
                  <Text style={{ color: palette.textMuted }}>×</Text>
                </Pressable>
              </View>
              {issue.reason && <Text style={{ color: isActive ? palette.text : palette.textMuted,
                fontSize: 11, marginTop: 4 }}>{issue.reason}</Text>}
            </Pressable>
          );
        })}
      </ScrollView>
    </View>
  );
}
