import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { readOfficeFile } from "../server/viewer";
import { clearWorkbookCache } from "../server/spreadsheet";

const workbookReads = vi.hoisted(() => vi.fn());
vi.mock("node:fs/promises", async (importOriginal) => {
  const fileSystem = await importOriginal<typeof import("node:fs/promises")>();
  // Count actual file reads rather than relying on flaky parse-duration assertions.
  return { ...fileSystem, readFile: (...args: Parameters<typeof fileSystem.readFile>) => {
    workbookReads(args[0]);
    return fileSystem.readFile(...args);
  } };
});

let workspaceDirectory = "";
afterEach(async () => {
  clearWorkbookCache();
  if (workspaceDirectory) await rm(workspaceDirectory, { recursive: true, force: true });
});

describe("active large workbook cache", () => {
  it("reuses a workbook exceeding the normal cell budget for subsequent scroll windows", async () => {
    workspaceDirectory = await mkdtemp(join(tmpdir(), "large-workbook-cache-"));
    const filePath = join(workspaceDirectory, "large.csv");
    // 510,000 populated cells exceed the real 500,000-cell budget without a multi-million-cell test artifact.
    await writeFile(filePath, "A,B,C,D,E,F,G,H,I,J\n".repeat(51_000));
    workbookReads.mockClear();
    const initialPreview = await readOfficeFile({ filePath });
    const laterPreview = await readOfficeFile({ filePath, rowStart: 50_000 });
    expect(initialPreview.sheets![0].rowCount).toBe(51_000);
    expect(laterPreview.sheets![0].rows[0][0]).toBe("A");
    expect(workbookReads).toHaveBeenCalledTimes(1);
  });
});
