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
 * see this module's header comment), in the same order PowerPoint's own
 * native color picker's "Theme Colors" row uses. Returns [] (not a thrown
 * error) when unsupported, so callers can just render nothing extra
 * rather than handle a failure case.
 *
 * Reads off the currently selected slide's own themeColorScheme, not
 * slideMasters.getItemAt(0) - a presentation can have more than one slide
 * master (e.g. a multi-layout template like the ones under
 * Package Files/.../Templates), and master index 0 isn't guaranteed to be
 * the one actually applied to the slide the user is looking at. The
 * selected slide's own scheme is always the one that's actually
 * WYSIWYG-correct for what's on screen, regardless of how many masters
 * the file has.
 */
/** Normalizes whatever ThemeColorScheme.getThemeColor().value actually comes back as into a "#RRGGBB" string — defensively, since real Office hosts aren't guaranteed to match what the docs/samples imply (e.g. a value that already carries its own "#"). */
function normalizeHex(value: string): string {
  const trimmed = value.trim();
  return trimmed.startsWith("#") ? trimmed.toUpperCase() : `#${trimmed.toUpperCase()}`;
}

export async function getThemeColors(): Promise<{ label: string; hex: string }[]> {
  if (!isThemeColorsSupported()) return [];
  return PowerPoint.run(async (context) => {
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();
    const slide = selectedSlides.items.length > 0 ? selectedSlides.items[0] : context.presentation.slides.getItemAt(0);
    const scheme = slide.themeColorScheme;
    const results = THEME_COLOR_ROLES.map((role) => ({ role, result: scheme.getThemeColor(role.key) }));
    await context.sync();
    return results.map(({ role, result }) => {
      const hex = normalizeHex(result.value);
      if (!/^#[0-9A-F]{6}$/.test(hex)) {
        console.warn(`Unexpected theme color value for ${role.label} (${role.key}): "${result.value}" -> "${hex}"`);
      }
      return { label: role.label, hex };
    });
  });
}
