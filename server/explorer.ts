import { readdir, stat } from "node:fs/promises";
import { extname, join } from "node:path";
import type { RpcInput, RpcOutput } from "@getpaseo/plugin";
import { listDirRpc, type DirectoryEntry } from "../shared/explorer";

// Omit heavy tool/version-control directories so expanding a workspace root stays fast and focused on documents.
const IGNORED_DIRECTORY_NAMES = new Set([".git", "node_modules", ".paseo", ".slp"]);

export async function listDirectoryEntries(
  input: RpcInput<typeof listDirRpc>,
): Promise<RpcOutput<typeof listDirRpc>> {
  const targetDirectoryPath = input.dirPath;
  try {
    const rawDirEntries = await readdir(targetDirectoryPath, { withFileTypes: true });
    const discoveredEntries: DirectoryEntry[] = [];

    for (const direntEntry of rawDirEntries) {
      // Skip ignored internal directories so users only navigate relevant workspace folders.
      if (direntEntry.isDirectory() && IGNORED_DIRECTORY_NAMES.has(direntEntry.name)) {
        continue;
      }

      const fullEntryPath = join(targetDirectoryPath, direntEntry.name);
      let isDirectoryFlag = direntEntry.isDirectory();
      let fileByteSize: number | undefined;

      // Follow symlinks via stat so linked document folders behave like regular directories in the tree.
      if (direntEntry.isSymbolicLink() || !direntEntry.isDirectory()) {
        try {
          const entryStats = await stat(fullEntryPath);
          isDirectoryFlag = entryStats.isDirectory();
          if (!isDirectoryFlag) {
            fileByteSize = entryStats.size;
          }
        } catch {
          // Ignore broken symlinks or transiently locked files rather than failing the whole folder listing.
          continue;
        }
      }

      const normalizedExtension = isDirectoryFlag ? "" : extname(direntEntry.name).toLowerCase();
      discoveredEntries.push({
        name: direntEntry.name,
        path: fullEntryPath,
        isDirectory: isDirectoryFlag,
        ext: normalizedExtension,
        size: fileByteSize,
      });
    }

    // Sort folders above files like Windows Explorer, using numeric collation for natural invoice numbering.
    discoveredEntries.sort((firstEntry, secondEntry) => {
      if (firstEntry.isDirectory !== secondEntry.isDirectory) {
        return firstEntry.isDirectory ? -1 : 1;
      }
      return firstEntry.name.localeCompare(secondEntry.name, undefined, {
        numeric: true,
        sensitivity: "base",
      });
    });

    return { entries: discoveredEntries };
  } catch {
    // Return an empty listing when a path disappears or lacks permission so the sidebar tree does not crash.
    return { entries: [] };
  }
}
