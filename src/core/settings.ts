/**
 * Replaces two VBA mechanisms with Office's cross-platform document settings API
 * (Office.context.document.settings), which persists small values with the document:
 *  - Public glb_Left/glb_Top/glb_Width/glb_Height/glb_Bottom/glb_LeftBottom globals
 *    used by Get Position / Set Position (K_Layout.bas).
 *  - ActivePresentation.CustomDocumentProperties "TemplateLeftOffset" etc. used by
 *    the Centre-on-Slide margin system (K_Layout.bas).
 */

const POSITION_KEY = "strategyToolbar.savedPosition";
const MARGINS_KEY = "strategyToolbar.templateMargins";

export interface SavedPosition {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface SavedBottomPosition {
  leftBottom: number;
  bottom: number;
  width: number;
  height: number;
}

export interface TemplateMargins {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

const ZERO_MARGINS: TemplateMargins = { left: 0, right: 0, top: 0, bottom: 0 };

function save(): void {
  Office.context.document.settings.saveAsync();
}

export function getSavedPosition(): SavedPosition | null {
  return (Office.context.document.settings.get(POSITION_KEY) as SavedPosition) ?? null;
}

export function setSavedPosition(pos: SavedPosition): void {
  Office.context.document.settings.set(POSITION_KEY, pos);
  save();
}

export function getSavedBottomPosition(): SavedBottomPosition | null {
  return (Office.context.document.settings.get(POSITION_KEY + ".bottom") as SavedBottomPosition) ?? null;
}

export function setSavedBottomPosition(pos: SavedBottomPosition): void {
  Office.context.document.settings.set(POSITION_KEY + ".bottom", pos);
  save();
}

export function getTemplateMargins(): TemplateMargins {
  return (Office.context.document.settings.get(MARGINS_KEY) as TemplateMargins) ?? ZERO_MARGINS;
}

export function setTemplateMargins(margins: TemplateMargins): void {
  Office.context.document.settings.set(MARGINS_KEY, margins);
  save();
}
