import type { ParagraphFormatting, TextFormatting } from "../types/document";
export const paragraphMarkRunPropertiesPatch = (
  originalFormatting: ParagraphFormatting | null | undefined,
  runProperties: TextFormatting | undefined,
): { _originalFormatting: ParagraphFormatting | null } => {
  const formatting: ParagraphFormatting = { ...(originalFormatting ?? {}) };
  if (runProperties === undefined) {
    Reflect.deleteProperty(formatting, "runProperties");
  } else {
    formatting.runProperties = structuredClone(runProperties);
  }
  return {
    _originalFormatting: Object.keys(formatting).length > 0 ? formatting : null,
  };
};
