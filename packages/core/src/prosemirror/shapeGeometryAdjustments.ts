import type { ShapeGeometryAdjustment } from "../types/document";

const MAX_ADJUSTMENT_COUNT = 32;
const MAX_ADJUSTMENT_NAME_LENGTH = 64;
const MAX_ADJUSTMENT_FORMULA_LENGTH = 128;

export const parseShapeGeometryAdjustments = (
  raw: string | undefined,
): ShapeGeometryAdjustment[] | undefined => {
  if (!raw) {
    return undefined;
  }
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed) || parsed.length > MAX_ADJUSTMENT_COUNT) {
      return undefined;
    }
    const adjustments: ShapeGeometryAdjustment[] = [];
    for (const item of parsed) {
      if (
        typeof item !== "object" ||
        item === null ||
        !("name" in item) ||
        typeof item.name !== "string" ||
        item.name.length === 0 ||
        item.name.length > MAX_ADJUSTMENT_NAME_LENGTH ||
        !("formula" in item) ||
        typeof item.formula !== "string" ||
        item.formula.length === 0 ||
        item.formula.length > MAX_ADJUSTMENT_FORMULA_LENGTH
      ) {
        return undefined;
      }
      adjustments.push({ name: item.name, formula: item.formula });
    }
    return adjustments.length === 0 ? undefined : adjustments;
  } catch {
    return undefined;
  }
};
