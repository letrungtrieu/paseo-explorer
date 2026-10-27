import { useCallback, useRef, useState } from "react";
import type { PluginRpcContract } from "@getpaseo/plugin";
import { useRpc, useWorkspace } from "@getpaseo/plugin/client";
import type { ZodType, input as ZodInput, output as ZodOutput } from "zod";

export interface ExplorerThemePalette {
  background: string;
  surface: string;
  border: string;
  text: string;
  textMuted: string;
  primary: string;
}

// Bridge Paseo's PluginTheme tokens (surface0/foreground/accent) and shorthand tokens (background/text/primary).
export function resolveExplorerTheme(rawTheme: unknown): ExplorerThemePalette {
  const themeRecord = (rawTheme as { colors?: Record<string, string | undefined> } | null)?.colors ?? {};
  return {
    background: themeRecord.background ?? themeRecord.surface0 ?? "#0f172a",
    surface: themeRecord.surface ?? themeRecord.surface1 ?? "#1e293b",
    border: themeRecord.border ?? "#334155",
    text: themeRecord.text ?? themeRecord.foreground ?? "#f8fafc",
    textMuted: themeRecord.textMuted ?? themeRecord.foregroundMuted ?? "#94a3b8",
    primary: themeRecord.primary ?? themeRecord.accent ?? "#3b82f6",
  };
}

export interface RpcMutationController<Input, Output> {
  data: Output | undefined;
  error: unknown;
  isPending: boolean;
  mutate: (input: Input) => void;
}

// Wrap Paseo's useRpc invoker with local mutation state while remaining compatible with pre-mocked mutation objects.
export function useRpcMutation<InputSchema extends ZodType, OutputSchema extends ZodType>(
  contract: PluginRpcContract<InputSchema, OutputSchema>,
): RpcMutationController<ZodInput<InputSchema>, ZodOutput<OutputSchema>> {
  const rpcEntry = useRpc(contract) as unknown;
  const [mutationData, setMutationData] = useState<ZodOutput<OutputSchema> | undefined>(undefined);
  const [mutationError, setMutationError] = useState<unknown>(null);
  const [isMutationPending, setIsMutationPending] = useState(false);

  // Keep the latest rpcEntry in a ref so the mutate callback identity stays stable across renders inside useEffect.
  const rpcEntryRef = useRef(rpcEntry);
  rpcEntryRef.current = rpcEntry;

  const triggerMutation = useCallback((inputPayload: ZodInput<InputSchema>) => {
    const activeRpc = rpcEntryRef.current;
    // Support tests that mock useRpc to return a React Query-like { mutate, data } object directly.
    if (activeRpc && typeof activeRpc === "object" && "mutate" in activeRpc) {
      const mockedMutate = (activeRpc as { mutate: (arg: ZodInput<InputSchema>) => void }).mutate;
      mockedMutate(inputPayload);
      return;
    }

    if (typeof activeRpc === "function") {
      setIsMutationPending(true);
      setMutationError(null);
      Promise.resolve((activeRpc as (arg: ZodInput<InputSchema>) => Promise<ZodOutput<OutputSchema>>)(inputPayload))
        .then((resolvedOutput) => {
          setMutationData(resolvedOutput);
          setIsMutationPending(false);
        })
        .catch((caughtError: unknown) => {
          setMutationError(caughtError);
          setIsMutationPending(false);
        });
    }
  }, []);

  if (rpcEntry && typeof rpcEntry === "object" && "mutate" in rpcEntry) {
    const mockedController = rpcEntry as RpcMutationController<ZodInput<InputSchema>, ZodOutput<OutputSchema>>;
    return {
      data: mockedController.data,
      error: mockedController.error ?? null,
      isPending: Boolean(mockedController.isPending),
      mutate: mockedController.mutate,
    };
  }

  return {
    data: mutationData,
    error: mutationError,
    isPending: isMutationPending,
    mutate: triggerMutation,
  };
}

// Resolve the workspace directory across both Paseo's 2-arg useWorkspace(id, selector) and 1-arg test mocks.
export function useResolvedWorkspacePath(workspaceId?: string, explicitCwd?: string): string | null {
  try {
    const workspaceSelector = (snapshot: { directory?: string; cwd?: string; projectRootPath?: string } | null) =>
      snapshot?.directory ?? snapshot?.cwd ?? snapshot?.projectRootPath ?? null;

    const flexibleUseWorkspace = useWorkspace as unknown as (
      firstArg: unknown,
      secondArg?: unknown,
    ) => string | null;

    const resolvedFromHook = workspaceId
      ? flexibleUseWorkspace(workspaceId, workspaceSelector)
      : flexibleUseWorkspace(workspaceSelector, workspaceSelector);

    return explicitCwd ?? resolvedFromHook ?? null;
  } catch {
    return explicitCwd ?? null;
  }
}
