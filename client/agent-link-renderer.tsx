import React, { useEffect } from "react";
import { View, Text, Pressable } from "react-native";
import type { PluginTimelineData } from "@getpaseo/plugin";
import type {
  PluginTimelineItemProps,
  PluginTimelineTransformerContribution,
} from "@getpaseo/plugin/client";
import { Markdown } from "@getpaseo/plugin/client/ui";
import { parseOfficeIssuesFromMessageText } from "../shared/office-parser";
import { officeStore, parseGridSelection } from "../shared/office-store";
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

// Render interactive issue cards in the Agent chat that sync badges to the file tree and open the Main Viewer Tab on click.
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
      {reportPayload.issues.map((issueItem, issueIndex) => {
        const formattedTargetLabel =
          parseGridSelection(issueItem.cell)?.label ?? issueItem.cell ?? "A1";
        const bracketLabel = issueItem.sheet
          ? `[${issueItem.sheet}!${formattedTargetLabel}]`
          : issueItem.cell
            ? `[${formattedTargetLabel}]`
            : `[Trang ${issueItem.page ?? 1}]`;

        return (
          <Pressable
            key={issueIndex}
            onPress={() =>
              officeStore.openAndJump(
                {
                  filePath: reportPayload.filePath,
                  ext: reportPayload.ext,
                  sheet: issueItem.sheet,
                  cell: issueItem.cell,
                  page: issueItem.page,
                  reason: issueItem.reason,
                },
                { openMainTab: true },
              )
            }
            style={({ pressed }) => ({
              flexDirection: "row",
              alignItems: "center",
              padding: 8,
              borderRadius: 6,
              borderWidth: 1,
              borderColor: "rgba(245, 158, 11, 0.55)",
              backgroundColor: pressed ? "rgba(245, 158, 11, 0.12)" : palette.surface,
            })}
          >
            <WindowsFileIcon isDirectory={false} ext={reportPayload.ext} />
            <View style={{ flex: 1 }}>
              <Text style={{ color: palette.text, fontWeight: "bold", fontSize: 12 }}>
                {shortFileName}{" "}
                <Text style={{ color: "#f59e0b" }}>{bracketLabel}</Text>
              </Text>
              <Text style={{ color: palette.textMuted, fontSize: 11, marginTop: 2 }}>
                {issueItem.reason}
              </Text>
            </View>
            <Text style={{ color: palette.primary, fontSize: 11, fontWeight: "600" }}>
              Mở ô →
            </Text>
          </Pressable>
        );
      })}
    </View>
  );
}

// Adapter component rendering the original Markdown message (if transformed from chat text) plus the interactive issue cards.
export function OfficeIssuesTimelineCard({
  item,
  theme,
}: PluginTimelineItemProps<OfficeIssueReportPayload>) {
  return (
    <View>
      {item.data.originalMessageText ? (
        <View style={{ marginBottom: 6 }}>
          <Markdown>{item.data.originalMessageText}</Markdown>
        </View>
      ) : null}
      <AgentOfficeLinkRenderer input={item.data} theme={theme} />
    </View>
  );
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

// Transform natural-language or inline-tag Office references in assistant messages into clickable link cards.
export const officeAssistantMessageTransformer: PluginTimelineTransformerContribution<"assistant_message"> =
  {
    id: "office-explorer-assistant-msg-transformer",
    query: { itemType: "assistant_message" },
    transform({ item, phase }) {
      if (phase !== "complete") return undefined;
      const parsedPayloads = parseOfficeIssuesFromMessageText(item.text);
      if (parsedPayloads.length === 0) return undefined;

      for (const reportGroup of parsedPayloads) {
        officeStore.registerIssues(
          reportGroup.filePath,
          reportGroup.issues.map((issueItem) => ({
            ...issueItem,
            ext: reportGroup.ext,
          })),
        );
      }

      return {
        items: parsedPayloads.map((reportGroup, groupIndex) => ({
          type: "plugin" as const,
          id: `office-msg-${item.messageId ?? "assistant"}-${groupIndex}`,
          kind: OFFICE_ISSUES_TIMELINE_KIND,
          version: OFFICE_ISSUES_TIMELINE_VERSION,
          data: {
            ...reportGroup,
            ...(groupIndex === 0 ? { originalMessageText: item.text } : {}),
          } as unknown as PluginTimelineData,
        })),
      };
    },
  };

// Transform user messages mentioning Office files + cells/rows/cols so users can also test links directly by typing in Chat.
export const officeUserMessageTransformer: PluginTimelineTransformerContribution<"user_message"> = {
  id: "office-explorer-user-msg-transformer",
  query: { itemType: "user_message" },
  transform({ item }) {
    const parsedPayloads = parseOfficeIssuesFromMessageText(item.text);
    if (parsedPayloads.length === 0) return undefined;

    for (const reportGroup of parsedPayloads) {
      officeStore.registerIssues(
        reportGroup.filePath,
        reportGroup.issues.map((issueItem) => ({
          ...issueItem,
          ext: reportGroup.ext,
        })),
      );
    }

    return {
      items: parsedPayloads.map((reportGroup, groupIndex) => ({
        type: "plugin" as const,
        id: `office-user-msg-${item.messageId ?? item.clientMessageId ?? "user"}-${groupIndex}`,
        kind: OFFICE_ISSUES_TIMELINE_KIND,
        version: OFFICE_ISSUES_TIMELINE_VERSION,
        data: {
          ...reportGroup,
          ...(groupIndex === 0 ? { originalMessageText: item.text } : {}),
        } as unknown as PluginTimelineData,
      })),
    };
  },
};
