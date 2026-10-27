import React, { useEffect } from "react";
import { View, Text, Pressable } from "react-native";
import type { PluginTimelineData } from "@getpaseo/plugin";
import type {
  PluginTimelineItemProps,
  PluginTimelineTransformerContribution,
} from "@getpaseo/plugin/client";
import { Markdown } from "@getpaseo/plugin/client/ui";
import {
  convertOfficeTagsToMarkdownLinks,
  parseOfficeIssuesFromMessageText,
} from "../shared/office-parser";
import { officeStore, parseGridSelection } from "../shared/office-store";
import {
  OFFICE_ISSUES_TIMELINE_KIND,
  OFFICE_ISSUES_TIMELINE_VERSION,
  OfficeIssueReportPayloadSchema,
  type OfficeIssueReportPayload,
} from "../shared/viewer";
import { resolveExplorerTheme } from "./theme-colors";

export interface ToolCallRendererProps {
  input: unknown;
  theme: unknown;
}

// Render interactive Office issue references as clean URL-style hyperlinks instead of boxed bottom cards.
export function AgentOfficeLinkRenderer({ input, theme }: ToolCallRendererProps) {
  const palette = resolveExplorerTheme(theme);
  const reportPayload = input as OfficeIssueReportPayload | undefined;

  // Push reported issues into the shared store immediately so the native Files tree displays "⚠ N" badges right away.
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
    <View style={{ gap: 4, marginVertical: 2 }}>
      {reportPayload.issues.map((issueItem, issueIndex) => {
        const formattedTargetLabel =
          parseGridSelection(issueItem.cell)?.label ?? issueItem.cell ?? "A1";
        const coordinateLabel = issueItem.sheet
          ? `#${issueItem.sheet}!${formattedTargetLabel}`
          : issueItem.cell
            ? `#${formattedTargetLabel}`
            : `#Trang ${issueItem.page ?? 1}`;
        const urlDisplayLabel = issueItem.reason
          ? `${shortFileName}${coordinateLabel} — ${issueItem.reason}`
          : `${shortFileName}${coordinateLabel}`;

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
            style={{ alignSelf: "flex-start", paddingVertical: 1 }}
          >
            {({ pressed }) => (
              <Text
                style={{
                  color: palette.primary,
                  fontSize: 13,
                  textDecorationLine: "underline",
                  opacity: pressed ? 0.7 : 1,
                }}
              >
                {urlDisplayLabel}
              </Text>
            )}
          </Pressable>
        );
      })}
    </View>
  );
}

// Render the message with inline URL hyperlinks for `[[...]]` tags, or URL-style links for natural-language/tool reports.
export function OfficeIssuesTimelineCard({
  item,
  theme,
}: PluginTimelineItemProps<OfficeIssueReportPayload>) {
  useEffect(() => {
    if (item.data.filePath && Array.isArray(item.data.issues)) {
      officeStore.registerIssues(
        item.data.filePath,
        item.data.issues.map((issueItem) => ({
          ...issueItem,
          ext: item.data.ext,
        })),
      );
    }
  }, [item.data]);

  if (item.data.originalMessageText) {
    const { markdownText, hasConvertedInlineTags } = convertOfficeTagsToMarkdownLinks(
      item.data.originalMessageText,
    );

    // When `[[...]]` tags were converted directly into inline Markdown URLs, omit any extra bottom link row.
    if (hasConvertedInlineTags) {
      return (
        <View>
          <Markdown>{markdownText}</Markdown>
        </View>
      );
    }

    return (
      <View>
        <View style={{ marginBottom: 4 }}>
          <Markdown>{markdownText}</Markdown>
        </View>
        <AgentOfficeLinkRenderer input={item.data} theme={theme} />
      </View>
    );
  }

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

// Transform Agent tool calls carrying Office validation issues into interactive URL-style timeline links.
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

// Transform natural-language or inline-tag Office references in assistant messages into clickable URL hyperlinks.
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

      const { hasConvertedInlineTags } = convertOfficeTagsToMarkdownLinks(item.text);
      const effectivePayloads = hasConvertedInlineTags ? [parsedPayloads[0]] : parsedPayloads;

      return {
        items: effectivePayloads.map((reportGroup, groupIndex) => ({
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

// Transform user messages mentioning Office files + cells/rows/cols into clickable URL hyperlinks.
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

    const { hasConvertedInlineTags } = convertOfficeTagsToMarkdownLinks(item.text);
    const effectivePayloads = hasConvertedInlineTags ? [parsedPayloads[0]] : parsedPayloads;

    return {
      items: effectivePayloads.map((reportGroup, groupIndex) => ({
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
