import React from "react";
import { View, Text } from "react-native";

export interface FileIconProps {
  isDirectory: boolean;
  isOpen?: boolean;
  ext: string;
}

export interface FileBadgePalette {
  label: string;
  bg: string;
  fg: string;
  border: string;
}

// Brand-accurate Office and data badge colors so users can distinguish file types at a glance in dense trees.
export const ICON_CONFIG: Record<string, FileBadgePalette> = {
  ".xlsx": { label: "XLS", bg: "#107c41", fg: "#ffffff", border: "#0b5c2f" },
  ".xls": { label: "XLS", bg: "#107c41", fg: "#ffffff", border: "#0b5c2f" },
  ".csv": { label: "CSV", bg: "#21a366", fg: "#ffffff", border: "#107c41" },
  ".docx": { label: "DOC", bg: "#185abd", fg: "#ffffff", border: "#103f87" },
  ".doc": { label: "DOC", bg: "#185abd", fg: "#ffffff", border: "#103f87" },
  ".pdf": { label: "PDF", bg: "#d93831", fg: "#ffffff", border: "#a8241e" },
  ".json": { label: "{ }", bg: "#f59e0b", fg: "#1e293b", border: "#d97706" },
};

// Normalize extension strings so both ".XLSX" and "xlsx" resolve to the same badge palette.
function normalizeExtensionKey(rawExtension: string): string {
  const trimmedExt = rawExtension.trim().toLowerCase();
  if (!trimmedExt) return "";
  return trimmedExt.startsWith(".") ? trimmedExt : `.${trimmedExt}`;
}

export function WindowsFileIcon({ isDirectory, isOpen, ext }: FileIconProps) {
  if (isDirectory) {
    // Render a two-part yellow Windows Explorer folder using pure React Native Views (no SVG dependency required).
    return (
      <View
        style={{
          width: 22,
          height: 18,
          justifyContent: "center",
          alignItems: "center",
          marginRight: 6,
        }}
      >
        {/* Upper folder tab positioned on the top-left edge */}
        <View
          style={{
            position: "absolute",
            top: 1,
            left: 1,
            width: 8,
            height: 4,
            backgroundColor: "#d97706",
            borderTopLeftRadius: 2,
            borderTopRightRadius: 2,
          }}
        />
        {/* Lighter gold fill when expanded so open folders visually stand out from collapsed ones */}
        <View
          style={{
            marginTop: 3,
            width: 18,
            height: 13,
            backgroundColor: isOpen ? "#fbbf24" : "#f59e0b",
            borderRadius: 2,
            borderWidth: 1,
            borderColor: "#b45309",
          }}
        />
      </View>
    );
  }

  const normalizedExtKey = normalizeExtensionKey(ext);
  const badgePalette = ICON_CONFIG[normalizedExtKey];
  if (badgePalette) {
    return (
      <View
        style={{
          width: 26,
          height: 16,
          backgroundColor: badgePalette.bg,
          borderColor: badgePalette.border,
          borderWidth: 1,
          borderRadius: 3,
          justifyContent: "center",
          alignItems: "center",
          marginRight: 6,
        }}
      >
        <Text style={{ color: badgePalette.fg, fontSize: 8, fontWeight: "800", letterSpacing: 0.3 }}>
          {badgePalette.label}
        </Text>
      </View>
    );
  }

  // Render a muted document rectangle for non-Office files so they remain aligned with badged siblings.
  return (
    <View
      style={{
        width: 18,
        height: 16,
        backgroundColor: "#e2e8f0",
        borderColor: "#94a3b8",
        borderWidth: 1,
        borderRadius: 2,
        marginRight: 8,
        marginLeft: 4,
      }}
    />
  );
}
