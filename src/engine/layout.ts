/**
 * Page layout constants and helpers.
 */

import { type XNode, attr, setAttr, findOne, el } from "./xml-helpers.js";

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

export const TWIPS_PER_MM = 1440 / 25.4; // ≈56.692913…
export const TWIPS_PER_INCH = 1440;

export function twipsToMm(twips: number): number {
  return Math.round((twips / TWIPS_PER_MM) * 10) / 10;
}

export function mmToTwips(mm: number): number {
  return Math.round(mm * TWIPS_PER_MM);
}

/** Well-known page sizes in twips (width × height in portrait) */
export const PAGE_SIZE_PRESETS: Record<string, { w: number; h: number; label: string }> = {
  A3:     { w: 16838, h: 23811, label: "A3 (297×420 mm)" },
  A4:     { w: 11906, h: 16838, label: "A4 (210×297 mm)" },
  A5:     { w: 8391,  h: 11906, label: "A5 (148×210 mm)" },
  B4:     { w: 14570, h: 20636, label: "B4 JIS (257×364 mm)" },
  B5:     { w: 10318, h: 14570, label: "B5 JIS (182×257 mm)" },
  LETTER: { w: 12240, h: 15840, label: "Letter (8.5×11 in)" },
  LEGAL:  { w: 12240, h: 20160, label: "Legal (8.5×14 in)" },
};

/** Well-known margin presets in twips */
export const MARGIN_PRESETS: Record<string, { top: number; right: number; bottom: number; left: number; label: string }> = {
  NORMAL:  { top: 1440, right: 1440, bottom: 1440, left: 1440, label: "Normal (1 in / 25.4 mm all)" },
  NARROW:  { top: 720,  right: 720,  bottom: 720,  left: 720,  label: "Narrow (0.5 in / 12.7 mm all)" },
  WIDE:    { top: 1440, right: 2880, bottom: 1440, left: 2880, label: "Wide (1 in top/bottom, 2 in left/right)" },
  JP_COURT_25: { top: 1418, right: 1418, bottom: 1418, left: 1418, label: "JP Court 25 mm (25 mm all)" },
  JP_COURT_30_20: { top: 1701, right: 1134, bottom: 1701, left: 1134, label: "JP Court 30/20 mm (30 mm top/bottom, 20 mm left/right)" },
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

export function detectPageSizePreset(w: number, h: number): string | null {
  // Normalize to portrait for comparison
  const pw = Math.min(w, h);
  const ph = Math.max(w, h);
  for (const [key, preset] of Object.entries(PAGE_SIZE_PRESETS)) {
    if (Math.abs(pw - preset.w) <= 10 && Math.abs(ph - preset.h) <= 10) {
      return key;
    }
  }
  return null;
}

export function getSectPr(body: XNode[]): XNode | undefined {
  return body.find((n: XNode) => n["w:sectPr"] !== undefined);
}

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface PageLayoutOptions {
  pageSizePreset?: string;
  marginPreset?: string;
  widthMm?: number;
  heightMm?: number;
  orientation?: "portrait" | "landscape";
  topMm?: number;
  rightMm?: number;
  bottomMm?: number;
  leftMm?: number;
  headerMm?: number;
  footerMm?: number;
  gutterMm?: number;
}
