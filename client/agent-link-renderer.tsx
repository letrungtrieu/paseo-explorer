import React, { useEffect } from "react";
import { View, Text, Pressable } from "react-native";
import type {
  PluginTimelineItemProps,
  PluginTimelineTransformerContribution,
} from "@getpaseo/plugin/client";
import { officeStore } from "../shared/office-store";
import {
  OFFICE_ISSUES_TIMELINE_KIND,
  OFFICE_ISSUES_TIMELINE_VERSION,
  OfficeIssueReportPayloadSchema,
  type OfficeIssueReportPayload,
} from "../shared/viewer";
import { WindowsFileIcon } from "./file-icons";
import { resolveExplorerTheme } from "./theme-colors";

export interface ToolCallRendererProps {
  input: unknown;
  theme: unknown;
}

// Render interactive issue cards in the Agent chat that sync badges to the file tree and jump to cells on click.
export function AgentOfficeLinkRenderer({ input, theme }: ToolCallRendererProps) {
  const palette = resolveExplorerTheme(theme);
  const reportPayload = input as OfficeIssueReportPayload | undefined;

  // Push reported issues into the shared store immediately so the file tree displays "⚠ N" badges right away.
  useEffect(() => {
    if (reportPayload?.filePath && Array.isArray(reportPayload?.issues)) {
      officeStore.registerIssues(
        reportPayload.filePath,
        reportPayload.issues.map((issueItem) => ({
          ...issueItem,
          ext: reportPayload.ext,
        })),
      );
    }
  }, [reportPayload]);

  if (!reportPayload?.issues || reportPayload.issues.length === 0) {
    return null;
  }

  const shortFileName = reportPayload.filePath.split(/[\\/]/).pop() ?? reportPayload.filePath;

  return (
    <View style={{ gap: 6, marginVertical: 6 }}>
      {reportPayload.issues.map((issueItem, issueIndex) => (
        <Pressable
          key={issueIndex}
          onPress={() =>
            officeStore.openAndJump({
              filePath: reportPayload.filePath,
              ext: reportPayload.ext,
              sheet: issueItem.sheet,
              cell: issueItem.cell,
              page: issueItem.page,
              reason: issueItem.reason,
            })
          }
          style={({ pressed }) => ({
            flexDirection: "row",
            alignItems: "center",
            padding: 8,
            borderRadius: 6,
            borderWidth: 1,
            borderColor: "#ef4444",
            backgroundColor: pressed ? "rgba(239, 68, 68, 0.1)" : palette.surface,
          })}
        >
          <WindowsFileIcon isDirectory={false} ext={reportPayload.ext} />
          <View style={{ flex: 1 }}>
            <Text style={{ color: palette.text, fontWeight: "bold", fontSize: 12 }}>
              {shortFileName}{" "}
              <Text style={{ color: "#ef4444" }}>
                {issueItem.sheet
                  ? `[${issueItem.sheet}!${issueItem.cell ?? "A1"}]`
                  : `[Trang ${issueItem.page ?? 1}]`}
              </Text>
            </Text>
            <Text style={{ color: palette.textMuted, fontSize: 11, marginTop: 2 }}>
              {issueItem.reason}
            </Text>
          </View>
          <Text style={{ color: palette.primary, fontSize: 11, fontWeight: "600" }}>
            Mở ô →
          </Text>
        </Pressable>
      ))}
    </View>
  );
}

// Adapter component connecting Paseo's PluginTimelineRenderer contract to AgentOfficeLinkRenderer.
export function OfficeIssuesTimelineCard({
  item,
  theme,
}: PluginTimelineItemProps<OfficeIssueReportPayload>) {
  return <AgentOfficeLinkRenderer input={item.data} theme={theme} />;
}

// Extract validated office issue payloads from tool_call timeline items (e.g., report_office_issues).
function extractIssuePayloadFromToolCall(toolItem: {
  name: string;
  detail: { type: string; input?: unknown; output?: unknown };
  metadata?: Record<string, unknown>;
}): OfficeIssueReportPayload | null {
  const candidates: unknown[] = [];
  if (toolItem.detail.type === "unknown") {
    candidates.push(toolItem.detail.input, toolItem.detail.output);
  }
  if (toolItem.metadata) {
    candidates.push(toolItem.metadata.input, toolItem.metadata);
  }

  for (const rawCandidate of candidates) {
    const validationResult = OfficeIssueReportPayloadSchema.safeParse(rawCandidate);
    if (validationResult.success) {
      return validationResult.data;
    }
  }
  return null;
}

// Transform Agent tool calls carrying Office validation issues into interactive timeline cards.
export const officeIssuesTimelineTransformer: PluginTimelineTransformerContribution<"tool_call"> = {
  id: "office-explorer-tool-transformer",
  query: { itemType: "tool_call" },
  transform({ item }) {
    const parsedPayload = extractIssuePayloadFromToolCall(item);
    if (!parsedPayload) return undefined;

    // Pre-register issues during timeline transformation so the file tree badge updates even if the chat card is offscreen.
    officeStore.registerIssues(
      parsedPayload.filePath,
      parsedPayload.issues.map((issueItem) => ({
        ...issueItem,
        ext: parsedPayload.ext,
      })),
    );

    return {
      items: [
        {
          type: "plugin",
          id: `office-issues-${item.callId}`,
          kind: OFFICE_ISSUES_TIMELINE_KIND,
          version: OFFICE_ISSUES_TIMELINE_VERSION,
          data: parsedPayload,
        },
      ],
    };
  },
};
