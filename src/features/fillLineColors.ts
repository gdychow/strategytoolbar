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
 * ribbon buttons used to open — dropped per the earlier finding that the
 * current templates already bake the same named brand palette into
 * PowerPoint's native color picker (see custClrLst audit). This module
 * exposes the action functions directly; pick a colour with a native
 * <input type="color"> or PowerPoint's own fill/line dropdown and call one
 * of these to apply it via the shared engine.
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
