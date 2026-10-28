import { useEffect, useRef, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { useRpc, type PluginFile } from "@getpaseo/plugin/client";
import { parseGridSelections, type OfficeTarget } from "../shared/office-store";
import { readOfficeFileRpc, type OfficeFileData } from "../shared/viewer";
import {
  WORKSHEET_ROW_BUFFER, WORKSHEET_WINDOW_COLUMNS, WORKSHEET_WINDOW_ROWS, type SpreadsheetWindow,
} from "../shared/spreadsheet-window";

interface BrowsingWindow extends SpreadsheetWindow {
  documentKey: string;
  targetKey: string;
  retainPrevious: boolean;
}

export function useOfficeDocument(target: OfficeTarget | null, workspaceDirectory: string | null, watchedFile?: PluginFile) {
  const readOfficeFile = useRpc(readOfficeFileRpc);
  const [fileRevision, setFileRevision] = useState(0);
  const [browsingWindow, setBrowsingWindow] = useState<BrowsingWindow | null>(null);
  const loadedSheet = useRef<{ documentKey: string; name: string } | null>(null);
  const lastLoadedDocument = useRef<{ documentKey: string; sheet?: string; contents: OfficeFileData } | null>(null);
  const documentKey = JSON.stringify([workspaceDirectory, target?.filePath, fileRevision]);
  const targetKey = JSON.stringify([target?.sheet, target?.cell, target?.timestamp]);
  const selection = parseGridSelections(target?.cell)[0];
  const selectedRow = selection?.startRow;
  const selectedColumn = selection?.startCol;
  const matchingSheet = !target?.sheet || browsingWindow?.sheet === target.sheet ||
    (!browsingWindow?.sheet && loadedSheet.current?.documentKey === documentKey && loadedSheet.current.name === target.sheet);
  // Selecting a loaded cell should only highlight it; a new distant chat jump owns a fresh centered window.
  const targetInsideWindow = browsingWindow && matchingSheet &&
    (selectedRow === undefined || selectedRow >= browsingWindow.rowStart && selectedRow < browsingWindow.rowStart + browsingWindow.rowLimit) &&
    (selectedColumn === undefined || selectedColumn >= browsingWindow.columnStart && selectedColumn < browsingWindow.columnStart + browsingWindow.columnLimit);
  const requestedWindow: BrowsingWindow = browsingWindow?.documentKey === documentKey &&
    (browsingWindow.targetKey === targetKey || targetInsideWindow)
    ? { ...browsingWindow, targetKey } : {
      documentKey, targetKey, sheet: target?.sheet,
      retainPrevious: false,
      rowStart: Math.max(0, (selectedRow ?? 0) - WORKSHEET_ROW_BUFFER),
      columnStart: Math.max(0, (selectedColumn ?? 0) - Math.floor(WORKSHEET_WINDOW_COLUMNS / 2)),
      rowLimit: WORKSHEET_WINDOW_ROWS, columnLimit: WORKSHEET_WINDOW_COLUMNS,
    };
  const { sheet, rowStart, columnStart, rowLimit, columnLimit, retainPrevious } = requestedWindow;
  // Accept new props before child effects run; a later synchronization effect could overwrite their scroll prefetch.
  if (browsingWindow?.documentKey !== documentKey || browsingWindow.targetKey !== targetKey) {
    setBrowsingWindow(requestedWindow);
  }

  const isNativePdf = watchedFile?.extension.toLowerCase() === "pdf";
  const canRetainWindow = retainPrevious && browsingWindow?.documentKey === documentKey && browsingWindow.targetKey === targetKey;
  const documentQuery = useQuery({
    queryKey: ["office-document", documentKey, sheet, rowStart, columnStart, rowLimit, columnLimit],
    queryFn: () => readOfficeFile({ filePath: target!.filePath, cwd: workspaceDirectory ?? undefined,
      sheet, rowStart, columnStart, rowLimit, columnLimit }),
    enabled: Boolean(target?.filePath) && !isNativePdf,
    // Keep existing cells mounted during scroll prefetch, while isolating file, sheet and revision changes.
    placeholderData: (previousDocument, previousQuery) => canRetainWindow &&
      previousQuery?.queryKey[1] === documentKey && (previousQuery.queryKey[2] === sheet ||
        previousQuery.queryKey[2] === undefined && loadedSheet.current?.name === sheet) ? previousDocument : undefined,
    retry: false,
    refetchOnWindowFocus: false,
    gcTime: 60_000,
  });
  useEffect(() => {
    if (documentQuery.data && !documentQuery.isPlaceholderData) {
      lastLoadedDocument.current = { documentKey, sheet, contents: documentQuery.data };
    }
  }, [documentKey, sheet, documentQuery.data, documentQuery.isPlaceholderData]);
  // Failed prefetch retains readable cells and the scroll anchor; file revisions and explicit jumps still clear stale data.
  const retainedDocument = retainPrevious && lastLoadedDocument.current?.documentKey === documentKey &&
    (lastLoadedDocument.current.sheet === sheet || lastLoadedDocument.current.sheet === undefined && loadedSheet.current?.name === sheet)
    ? lastLoadedDocument.current.contents : undefined;
  const displayedDocument = documentQuery.data ?? retainedDocument;
  // Re-read on edits and deletions rather than retaining a preview of a superseded file revision.
  useEffect(() => watchedFile?.watch(() => setFileRevision((revision) => revision + 1)), [watchedFile]);

  const currentWorksheet = displayedDocument?.sheets?.find((worksheet) =>
    worksheet.name.toLowerCase() === sheet?.toLowerCase()) ?? displayedDocument?.sheets?.[0];
  useEffect(() => {
    // The first-sheet request omits a name; remember its resolved name without a redundant second read.
    if (currentWorksheet) loadedSheet.current = { documentKey, name: currentWorksheet.name };
  }, [documentKey, currentWorksheet?.name]);

  return { documentQuery: { ...documentQuery, data: displayedDocument }, currentWorksheet, isNativePdf,
    navigate: (sheetName: string, window?: SpreadsheetWindow) => {
      // Sheet tabs reset to the beginning; scroll requests retain overlapping row and column boundaries.
      setBrowsingWindow({ documentKey, targetKey, sheet: sheetName,
        retainPrevious: Boolean(window),
        rowStart: 0, columnStart: 0, rowLimit: WORKSHEET_WINDOW_ROWS, columnLimit: WORKSHEET_WINDOW_COLUMNS, ...window });
    },
  };
}
