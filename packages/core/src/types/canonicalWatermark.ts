import type { SetDocumentWatermarkOp } from "@stll/docx-core/ops";
import type { CANONICAL_GAP } from "./canonicalCapabilities";

export type CanonicalWatermarkRequest = SetDocumentWatermarkOp["change"];
export type CanonicalWatermarkResult =
  | { status: "applied"; version: number }
  | { status: "refused"; gap: typeof CANONICAL_GAP.watermark; message: string };
