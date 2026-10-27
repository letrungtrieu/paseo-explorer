import type { PluginClientContext } from "@getpaseo/plugin/client";
import { parseOfficeIssuesFromMessageText } from "./shared/office-parser";
import { officeStore } from "./shared/office-store";
import {
  OFFICE_ISSUES_TIMELINE_KIND,
  OFFICE_ISSUES_TIMELINE_VERSION,
  OfficeIssueReportPayloadSchema,
} from "./shared/viewer";
import {
  officeAssistantMessageTransformer,
  OfficeIssuesTimelineCard,
  officeIssuesTimelineTransformer,
  officeUserMessageTransformer,
} from "./client/agent-link-renderer";
import {
  OfficeExplorerSidebar,
  OfficeFilePanel,
  OfficeViewerPanel,
} from "./client/office-explorer-sidebar";

export const OFFICE_EXPLORER_PANEL_ID = "office-explorer";
export const OFFICE_VIEWER_PANEL_ID = "office-viewer";
export const OFFICE_FILE_PANEL_ID = "office-file-viewer";

// Register the Right Sidebar Tree ("explorer"), Main Tab Viewer ("workspace" + addFilePanel), and Chat Link Transformers.
export default function contribute(client: PluginClientContext) {
  let latestActiveWorkspaceId: string | null = null;

  // Whenever a file is clicked in the right-sidebar tree or an issue is clicked in Chat, open the Viewer as a Main Tab.
  officeStore.setWorkspaceTabOpener((requestedWorkspaceId) => {
    const targetWorkspaceId = requestedWorkspaceId ?? latestActiveWorkspaceId;
    if (!targetWorkspaceId) return;
    try {
      client.openPanel(OFFICE_VIEWER_PANEL_ID, {
        workspaceId: targetWorkspaceId,
        location: "workspace",
      });
    } catch {
      // Ignore if the workspace window is not currently focused.
    }
  });

  const cleanupCallbacks = [
    // 1. Native File Tab viewer: opens a new tab in the main pane when clicking .xlsx/.docx/.pdf in Paseo's Files tree.
    client.addFilePanel({
      id: OFFICE_FILE_PANEL_ID,
      title: "Office Viewer",
      icon: "FileSpreadsheet",
      files: {
        extensions: ["xlsx", "xls", "csv", "docx", "doc", "pdf"],
      },
      Component: OfficeFilePanel,
    }),
    // 2. Main Workspace Tab viewer: opened as a main tab when clicking files in the custom Windows Tree or Agent Chat links.
    client.addWorkspacePanel({
      id: OFFICE_VIEWER_PANEL_ID,
      title: "Office Viewer",
      icon: "FileSpreadsheet",
      locations: ["workspace"],
      Component: OfficeViewerPanel,
    }),
    // 3. Right Sidebar File Tree: mounts strictly in the right-hand Explorer dock (tree only).
    client.addWorkspacePanel({
      id: OFFICE_EXPLORER_PANEL_ID,
      title: "Office Explorer",
      icon: "FolderTree",
      locations: ["explorer"],
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
    // 4. Slash command `/office-jump` to quickly jump to any file + cell/row/col from the chat composer.
    client.addSlashCommand({
      name: "office-jump",
      description: "Mở file Office và nhảy tới đúng ô/dòng/cột (VD: E3 trong file \"examples/Ngân sách hàng tháng.xlsx\" lỗi để trống)",
      argumentHint: "<mô tả lỗi hoặc [[file.xlsx#E3 | Lý do]]>",
      context: "workspace",
      onSubmit({ args, workspace, openPanel }) {
        const parsedGroups = parseOfficeIssuesFromMessageText(args);
        const firstGroup = parsedGroups[0];
        const firstIssue = firstGroup?.issues[0];
        if (!firstGroup || !firstIssue) return;

        officeStore.registerIssues(
          firstGroup.filePath,
          firstGroup.issues.map((issueItem) => ({ ...issueItem, ext: firstGroup.ext })),
        );
        officeStore.openAndJump(
          {
            filePath: firstGroup.filePath,
            ext: firstGroup.ext,
            sheet: firstIssue.sheet,
            cell: firstIssue.cell,
            page: firstIssue.page,
            reason: firstIssue.reason,
          },
          { workspaceId: workspace.id, openMainTab: false },
        );
        openPanel(OFFICE_VIEWER_PANEL_ID, { location: "workspace" });
      },
    }),
    client.addTimelineTransformer(officeIssuesTimelineTransformer),
    client.addTimelineTransformer(officeAssistantMessageTransformer),
    client.addTimelineTransformer(officeUserMessageTransformer),
    client.addTimelineRenderer({
      kind: OFFICE_ISSUES_TIMELINE_KIND,
      version: OFFICE_ISSUES_TIMELINE_VERSION,
      schema: OfficeIssueReportPayloadSchema,
      Component: OfficeIssuesTimelineCard,
    }),
  ];

  // Auto-mount the Windows Explorer tree tab in the right sidebar for open workspaces.
  setTimeout(() => {
    client.paseo?.workspaces
      ?.list()
      .then((workspaceList) => {
        for (const workspaceEntry of workspaceList.entries) {
          latestActiveWorkspaceId = workspaceEntry.id;
          try {
            client.openPanel(OFFICE_EXPLORER_PANEL_ID, {
              workspaceId: workspaceEntry.id,
              location: "explorer",
            });
          } catch {
            // Ignore workspaces not mounted in the current window.
          }
        }
      })
      .catch(() => undefined);
  }, 50);

  const unsubscribeWorkspaces = client.paseo?.workspaces?.subscribe((workspaceUpdate) => {
    if (workspaceUpdate?.kind === "upsert" && workspaceUpdate.workspace.id) {
      latestActiveWorkspaceId = workspaceUpdate.workspace.id;
    }
  });

  return () => {
    officeStore.setWorkspaceTabOpener(null);
    unsubscribeWorkspaces?.();
    for (const unregisterContribution of cleanupCallbacks) {
      unregisterContribution();
    }
  };
}
