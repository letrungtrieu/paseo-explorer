import { defineRpc } from "@getpaseo/plugin";
import { z } from "zod";

// Extensions that can be opened directly inside the integrated Office Explorer viewer.
export const SUPPORTED_EXTS = [".xlsx", ".xls", ".csv", ".docx", ".doc", ".pdf"] as const;

export const DirectoryEntrySchema = z.object({
  name: z.string(),
  path: z.string(),
  isDirectory: z.boolean(),
  // Normalized lowercase file extension including the leading dot (e.g. ".xlsx"), or "" for directories.
  ext: z.string(),
  size: z.number().optional(),
});

export type DirectoryEntry = z.infer<typeof DirectoryEntrySchema>;

// RPC contract used by each expanded tree node to lazily fetch its immediate children from the daemon.
// Lowercase snake_case is required by @getpaseo/plugin's RPC_NAME validation regex (/^[a-z][a-z0-9._-]*$/).
export const listDirRpc = defineRpc({
  name: "office.explorer.list_dir",
  input: z.object({
    dirPath: z.string().min(1),
  }),
  output: z.object({
    entries: z.array(DirectoryEntrySchema),
  }),
});
