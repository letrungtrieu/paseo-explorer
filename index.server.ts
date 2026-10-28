import type { PluginServerContext } from "@getpaseo/plugin/server";
import { listDirRpc } from "./shared/explorer";
import { readOfficeFileRpc } from "./shared/viewer";
import { listDirectoryEntries } from "./server/explorer";
import { readOfficeFile } from "./server/viewer";
import { clearWorkbookCache } from "./server/spreadsheet";
import { registerOfficePromptInjector } from "./server/prompt-injector";

// Register RPC endpoints and the agent.create prompt-injection hook so Office Explorer works out-of-the-box.
export default function contribute(server: PluginServerContext) {
  server.handle(listDirRpc, listDirectoryEntries);
  server.handle(readOfficeFileRpc, readOfficeFile);
  const unregisterPromptInjector = registerOfficePromptInjector(server);

  return () => {
    // Release parsed documents when reloading the plugin during an active Paseo session.
    clearWorkbookCache();
    unregisterPromptInjector();
  };
}
