import type { PluginBeforeRequests, PluginServerContext } from "@getpaseo/plugin/server";

// Unique marker prevents duplicate prompt injection when multiple hooks or session clones reuse config.
export const OFFICE_EXPLORER_PROMPT_MARKER = "[PASEO_OFFICE_EXPLORER_PROMPT]";

// Concise, token-efficient system instruction teaching agents (including Antigravity CLI) how to emit clickable Office issue links.
export const OFFICE_EXPLORER_SYSTEM_PROMPT = `${OFFICE_EXPLORER_PROMPT_MARKER}
## Office Explorer Integration (Excel / CSV / Word / PDF Issue Linking)

When reporting errors, anomalies, empty cells, or highlights in Office documents (\`.xlsx\`, \`.xls\`, \`.csv\`, \`.docx\`, \`.doc\`, \`.pdf\`), format references so Paseo's Office Explorer can automatically highlight the file in the Right Sidebar tree and jump to the exact cell, row, column, range, or page in a new Workspace Tab when clicked:

1. **Inline Interactive Link Tag (Preferred)**:
   - **Single Cell**: \`[[path/to/file.xlsx#SheetName!E3 | Reason why cell E3 has an issue]]\` (or \`[[path/to/file.xlsx#E3 | Reason]]\` for the first sheet)
   - **Entire Row**: \`[[path/to/file.xlsx#SheetName!Row:5 | Reason for row 5]]\` (or \`Dòng 5\`)
   - **Entire Column**: \`[[path/to/file.xlsx#SheetName!Col:E | Reason for column E]]\` (or \`Cột E\`)
   - **Cell Range**: \`[[path/to/file.xlsx#SheetName!B4:D8 | Reason for range B4:D8]]\`
   - **Word / PDF Document**: \`[[path/to/document.pdf#Trang 2 | Reason on page 2]]\` or \`[[path/to/contract.docx | Reason]]\`

2. **Natural Language Reference (Also Auto-Detected)**:
   - \`E3 trong file "examples/Ngân sách hàng tháng.xlsx" lỗi để trống\`
   - \`Dòng 5 trong file "examples/Ngân sách hàng tháng.xlsx" thiếu thông tin\`
   - \`Cột E trong file "examples/Ngân sách hàng tháng.xlsx" sai định dạng\`
   - \`Vùng B4:D8 trong file "examples/Ngân sách hàng tháng.xlsx" lệch tổng\`

Always include the workspace-relative file path (with extension) and the specific coordinate (\`E3\`, \`Row:5\`, \`Col:E\`, \`B4:D8\`, or \`Sheet!Cell\`) whenever pointing out spreadsheet or document issues.`;

// Merges the Office Explorer instructions into an existing system prompt without duplicating them.
export function injectOfficeExplorerSystemPrompt(existingSystemPrompt?: string): string {
  const normalizedExistingPrompt = existingSystemPrompt?.trim() ?? "";
  if (normalizedExistingPrompt.includes(OFFICE_EXPLORER_PROMPT_MARKER)) {
    return normalizedExistingPrompt;
  }

  // Append after existing instructions so primary role prompts keep top precedence.
  return normalizedExistingPrompt.length > 0
    ? `${normalizedExistingPrompt}\n\n${OFFICE_EXPLORER_SYSTEM_PROMPT}`
    : OFFICE_EXPLORER_SYSTEM_PROMPT;
}

// Transforms an agent.create request by injecting Office Explorer instructions for interactive user-facing sessions.
export function augmentAgentCreateRequestWithOfficePrompt(
  agentCreateRequest: PluginBeforeRequests["agent.create"],
): PluginBeforeRequests["agent.create"] {
  // Internal ephemeral agents (e.g. commit message or title generators) do not inspect Office files.
  if (agentCreateRequest.config.internal === true) {
    return agentCreateRequest;
  }

  const augmentedSystemPrompt = injectOfficeExplorerSystemPrompt(
    agentCreateRequest.config.systemPrompt,
  );

  return {
    ...agentCreateRequest,
    config: {
      ...agentCreateRequest.config,
      systemPrompt: augmentedSystemPrompt,
    },
  };
}

// Registers the Paseo server lifecycle hook so every new agent automatically receives Office Explorer reporting instructions.
export function registerOfficePromptInjector(server: PluginServerContext): () => void {
  if (typeof server.before !== "function") {
    return () => {};
  }

  return server.before("agent.create", ({ request: agentCreateRequest }) => {
    return augmentAgentCreateRequestWithOfficePrompt(agentCreateRequest);
  });
}
