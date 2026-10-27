import type { PluginClientContext } from "@getpaseo/plugin/client";
import {
  OFFICE_ISSUES_TIMELINE_KIND,
  OFFICE_ISSUES_TIMELINE_VERSION,
  OfficeIssueReportPayloadSchema,
} from "./shared/viewer";
import {
  OfficeIssuesTimelineCard,
  officeIssuesTimelineTransformer,
} from "./client/agent-link-renderer";
import { OfficeExplorerSidebar } from "./client/office-explorer-sidebar";

export const OFFICE_EXPLORER_PANEL_ID = "office-explorer";

// Register the right-sidebar Office Explorer panel and chat timeline link cards, returning a unified cleanup.
export default function contribute(client: PluginClientContext) {
  const cleanupCallbacks = [
    client.addWorkspacePanel({
      id: OFFICE_EXPLORER_PANEL_ID,
      title: "Office Explorer",
      icon: "FolderTree",
      // Mount in both the right-hand explorer dock and workspace panel switcher.
      locations: ["explorer", "workspace"],
      Component: OfficeExplorerSidebar,
    }),
    client.addCommandCenterItem({
      id: "open-office-explorer",
      title: "Open Office Explorer Sidebar",
      icon: "FolderTree",
      context: "workspace",
      keywords: ["office", "explorer", "excel", "xlsx", "docx", "pdf", "tree"],
      onSelect({ openPanel }) {
        openPanel(OFFICE_EXPLORER_PANEL_ID, { location: "explorer" });
      },
    }),
    client.addTimelineTransformer(officeIssuesTimelineTransformer),
    client.addTimelineRenderer({
      kind: OFFICE_ISSUES_TIMELINE_KIND,
      version: OFFICE_ISSUES_TIMELINE_VERSION,
      schema: OfficeIssueReportPayloadSchema,
      Component: OfficeIssuesTimelineCard,
    }),
  ];

  return () => {
    for (const unregisterContribution of cleanupCallbacks) {
      unregisterContribution();
    }
  };
}
