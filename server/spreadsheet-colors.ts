import type { Color } from "exceljs";

// Office's indexed palette predates theme colors and still appears in files exported by other spreadsheet apps.
const INDEXED_COLORS = ("000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF " +
  "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF " +
  "800000 008000 000080 808000 800080 008080 C0C0C0 808080 " +
  "9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCCC " +
  "000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF " +
  "00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 " +
  "3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 " +
  "003366 339966 003300 333300 993300 993366 333399 333333").split(" ");
const DEFAULT_THEME = ["FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD", "C0504D",
  "9BBB59", "8064A2", "4BACC6", "F79646", "0000FF", "800080"];

export function readSpreadsheetTheme(themeXml?: string): string[] {
  if (!themeXml) return DEFAULT_THEME;
  // Theme indices use light/dark order, which differs from the element order inside clrScheme.
  const themeKeys = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
    "accent4", "accent5", "accent6", "hlink", "folHlink"];
  return themeKeys.map((themeKey, themeIndex) => {
    const colorElement = themeXml.match(new RegExp(`<[^<>]*:?${themeKey}\\b[^>]*>([\\s\\S]*?)</[^<>]*:?${themeKey}>`))?.[1];
    // System colors need lastClr; a host OS color name such as windowText is not a portable CSS color.
    return colorElement?.match(/(?:lastClr|val)=["']([\dA-Fa-f]{6})["']/)?.[1] ?? DEFAULT_THEME[themeIndex];
  });
}

function applyColorTint(rgbHex: string, tint: number): string {
  const channels = [0, 2, 4].map((offset) => Number.parseInt(rgbHex.slice(offset, offset + 2), 16) / 255);
  const brightest = Math.max(...channels);
  const darkest = Math.min(...channels);
  const originalLightness = (brightest + darkest) / 2;
  const chroma = brightest - darkest;
  // Excel tints change HSL luminance; interpolating RGB channels gives visibly wrong accent shades.
  const saturation = chroma === 0 ? 0 : chroma / (1 - Math.abs(2 * originalLightness - 1));
  let hue = 0;
  if (chroma > 0) {
    const brightestChannel = channels.indexOf(brightest);
    hue = brightestChannel === 0 ? (channels[1] - channels[2]) / chroma
      : brightestChannel === 1 ? (channels[2] - channels[0]) / chroma + 2
        : (channels[0] - channels[1]) / chroma + 4;
    hue = ((hue / 6) + 1) % 1;
  }
  const lightness = tint < 0 ? originalLightness * (1 + tint)
    : originalLightness * (1 - tint) + tint;
  const tintedChroma = (1 - Math.abs(2 * lightness - 1)) * saturation;
  const hueSector = hue * 6;
  const intermediate = tintedChroma * (1 - Math.abs(hueSector % 2 - 1));
  const sectorChannels = [[tintedChroma, intermediate, 0], [intermediate, tintedChroma, 0],
    [0, tintedChroma, intermediate], [0, intermediate, tintedChroma],
    [intermediate, 0, tintedChroma], [tintedChroma, 0, intermediate]][Math.floor(hueSector) % 6];
  // Rounding only after the color conversion avoids cumulative tint errors on light theme fills.
  return sectorChannels.map((channel) => Math.round((channel + lightness - tintedChroma / 2) * 255)
    .toString(16).padStart(2, "0")).join("");
}

export function resolveSpreadsheetColor(
  color: (Partial<Color> & { indexed?: number; tint?: number }) | undefined, theme: string[],
): string | undefined {
  if (!color) return undefined;
  // Excel commonly writes 00-prefixed ARGB for opaque cell colors; the RGB portion is authoritative here.
  const rgbHex = color.argb?.slice(-6) ?? (color.theme !== undefined ? theme[color.theme] : undefined)
    ?? (color.indexed !== undefined ? INDEXED_COLORS[color.indexed] : undefined);
  if (!rgbHex || !/^[\dA-Fa-f]{6}$/.test(rgbHex)) return undefined;
  return `#${color.tint ? applyColorTint(rgbHex, color.tint) : rgbHex}`.toUpperCase();
}
