import { detectBaseDirection } from "../packages/core/src/utils/baseDirection";

import type { TextDirection } from "./types";

/** Preserve an undecided state around the renderer's canonical first-strong detector. */
export const firstStrongTextDirection = (text: string): TextDirection =>
  detectBaseDirection(text) ?? "unknown";
