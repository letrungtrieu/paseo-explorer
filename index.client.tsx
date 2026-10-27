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
  OfficeFilePanel,
  OfficeViewerPanel,
} from "./client/office-explorer-sidebar";
import {
  startNativeFilesTreeEnhancer,
  tryOpenNativeFileTabFromDom,
} from "./client/web";

export const OFFICE_VIEWER_PANEL_ID = "office-viewer";
export const OFFICE_FILE_PANEL_ID = "office-file-viewer";

// Register native File Tab viewer, native Files tree DOM enhancer, and Chat Link Transformers.
export default function contribute(client: PluginClientContext) {
  let latestActiveWorkspaceId: string | null = null;

  // Prefer focusing/opening Paseo's native file tab from the Files tree; fall back to the workspace viewer tab.
  officeStore.setWorkspaceTabOpener((requestedWorkspaceId) => {
    const activeTargetPath = officeStore.getState().activeTarget?.filePath;
    if (activeTargetPath && tryOpenNativeFileTabFromDom(activeTargetPath)) {
      return;
    }

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

  // Decorate Paseo's built-in Files sidebar tab with XLS/CSV/DOC/PDF icons and `⚠` issue count pills.
  const stopNativeFilesEnhancer = startNativeFilesTreeEnhancer();

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
    // 2. Fallback Workspace Tab viewer: opened when jumping from Chat links to a file whose folder is collapsed in Files.
    client.addWorkspacePanel({
      id: OFFICE_VIEWER_PANEL_ID,
      title: "Office Viewer",
      icon: "FileSpreadsheet",
      locations: ["workspace"],
      Component: OfficeViewerPanel,
    }),
    // 3. Slash command `/office-jump` to quickly jump to any file + cell/row/col from the chat composer.
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
        if (!tryOpenNativeFileTabFromDom(firstGroup.filePath)) {
          openPanel(OFFICE_VIEWER_PANEL_ID, { location: "workspace" });
        }
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

  // Track the active workspace ID so Chat links can open the fallback viewer tab even before a turn starts.
  client.paseo?.workspaces
    ?.list()
    .then((workspaceList) => {
      for (const workspaceEntry of workspaceList.entries) {
        latestActiveWorkspaceId = workspaceEntry.id;
      }
    })
    .catch(() => undefined);

  const unsubscribeWorkspaces = client.paseo?.workspaces?.subscribe((workspaceUpdate) => {
    if (workspaceUpdate?.kind === "upsert" && workspaceUpdate.workspace.id) {
      latestActiveWorkspaceId = workspaceUpdate.workspace.id;
    }
  });

  return () => {
    stopNativeFilesEnhancer();
    officeStore.setWorkspaceTabOpener(null);
    unsubscribeWorkspaces?.();
    for (const unregisterContribution of cleanupCallbacks) {
      unregisterContribution();
    }
  };
}
