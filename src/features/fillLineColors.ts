/**
 * Port of M_Fill_and_Line_Colors.bas.
 *
 * The module's five live subs (LineColor, FillColor, F22NoFill, L22NoLine,
 * TextColor) are all thin wrappers around ObjectFormat — the group/table
 * handling they used to do by hand was refactored into that shared engine
 * before this port. Everything else in the source file (LineColor_OLD,
 * FillColor_old, L22NoLine_OLD, F22NoFill_OLD, TextColor_OLD — ~800 lines)
 * is explicitly marked "DEPRECATED?" by the original author and, confirmed
 * by grep, is unreferenced anywhere else in the codebase. Not ported.
 *
 * Not ported: the custom-colour swatch palette / frmColourPicker UI those
 * ribbon buttons used to open — originally dropped per the finding that
 * the templates bake the same named brand palette into PowerPoint's own
 * custClrLst (see custClrLst audit). That assumption doesn't hold once
 * the picker moved off the native <input type="color"> (see
 * getThemeColors below) - custClrLst itself still has no Office.js
 * accessor at all (confirmed: nothing in @types/office-js references it),
 * so the closest available substitute is the 12-role ThemeColorScheme
 * (Accent1-6/Dark1-2/Light1-2/Hyperlink/FollowedHyperlink), which is a
 * real, different OOXML concept but the only one Office.js exposes.
 */

import { objectFormat, RGB } from "../core/objectFormat";

export const fillColor = (color: RGB) => objectFormat("Fill", { shape: { color } });
export const lineColor = (color: RGB) => objectFormat("Line", { shape: { color } });
export const textColor = (color: RGB) => objectFormat("Text", { text: { color } });
export const noFill = () => objectFormat("Fill", { shape: { transparent: true } });
export const noLine = () => objectFormat("Line", { shape: { transparent: true } });

/** Parses a "#RRGGBB" string from an <input type="color"> into an RGB triple. */
export function hexToRgb(hex: string): RGB {
  const clean = hex.replace("#", "");
  return {
    r: parseInt(clean.substring(0, 2), 16),
    g: parseInt(clean.substring(2, 4), 16),
    b: parseInt(clean.substring(4, 6), 16),
  };
}

/** Slide masters' ThemeColorScheme, like slide dimensions, needs PowerPointApi 1.10 (Mac support only since Dec 2025). */
export function isThemeColorsSupported(): boolean {
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
}

// String literals, not PowerPoint.ThemeColor enum members - the enum
// values are identical strings (confirmed against @types/office-js), but
// referencing the real enum object here would evaluate PowerPoint.ThemeColor
// at module load time, before Office.js has necessarily finished
// initializing that global - confirmed to break the whole bundle (a
// top-level ReferenceError aborts every subsequent script-level statement,
// including registering Office.onReady itself) when tested against a
// harness where the PowerPoint global isn't yet defined at parse time.
const THEME_COLOR_ROLES: { key: "Light1" | "Dark1" | "Light2" | "Dark2" | "Accent1" | "Accent2" | "Accent3" | "Accent4" | "Accent5" | "Accent6"; label: string }[] = [
  { key: "Light1", label: "Background 1" },
  { key: "Dark1", label: "Text 1" },
  { key: "Light2", label: "Background 2" },
  { key: "Dark2", label: "Text 2" },
  { key: "Accent1", label: "Accent 1" },
  { key: "Accent2", label: "Accent 2" },
  { key: "Accent3", label: "Accent 3" },
  { key: "Accent4", label: "Accent 4" },
  { key: "Accent5", label: "Accent 5" },
  { key: "Accent6", label: "Accent 6" },
];

/**
 * Reads the current presentation's actual theme colors (the closest
 * Office.js equivalent to "the custom colours built into the template" -
 * see this module's header comment) off the first slide master, in the
 * same order PowerPoint's own native color picker's "Theme Colors" row
 * uses. Returns [] (not a thrown error) when unsupported, so callers can
 * just render nothing extra rather than handle a failure case.
 */
export async function getThemeColors(): Promise<{ label: string; hex: string }[]> {
  if (!isThemeColorsSupported()) return [];
  return PowerPoint.run(async (context) => {
    const scheme = context.presentation.slideMasters.getItemAt(0).themeColorScheme;
    const results = THEME_COLOR_ROLES.map((role) => ({ role, result: scheme.getThemeColor(role.key) }));
    await context.sync();
    return results.map(({ role, result }) => ({ label: role.label, hex: `#${result.value}` }));
  });
}
