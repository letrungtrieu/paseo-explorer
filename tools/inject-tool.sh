#!/usr/bin/env bash
# Discover executable project tools in tools/bin/ and print their concise one-line descriptions.
set -euo pipefail

tools_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
bin_dir="$tools_dir/bin"

if [[ ! -d "$bin_dir" ]]; then
  exit 0
fi

# Iterate in deterministic sorted order and continue if any individual tool fails its --desc call.
for tool_executable in "$bin_dir"/*; do
  if [[ -f "$tool_executable" && -x "$tool_executable" ]]; then
    tool_name="$(basename "$tool_executable")"
    tool_description="$("$tool_executable" --desc 2>/dev/null || true)"
    if [[ -n "$tool_description" ]]; then
      printf "%s: %s\n" "$tool_name" "$tool_description"
    fi
  fi
done
