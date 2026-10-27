import type { PluginServerContext } from "@getpaseo/plugin/server";
import { listDirRpc } from "./shared/explorer";
import { readOfficeFileRpc } from "./shared/viewer";
import { listDirectoryEntries } from "./server/explorer";
import { readOfficeFile } from "./server/viewer";

// Register RPC endpoints synchronously so the client sidebar can browse directories and read Office files immediately.
export default function contribute(server: PluginServerContext) {
  server.handle(listDirRpc, listDirectoryEntries);
  server.handle(readOfficeFileRpc, readOfficeFile);
  return () => {};
}
