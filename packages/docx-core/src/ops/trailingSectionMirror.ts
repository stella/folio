import type { DocumentBody } from "../model/document";

/** Match an exact cold section mirror after the caller rules out block section breaks. */
export const hasExactTrailingSection = (body: DocumentBody): boolean => {
  if (body.sections?.length !== 1 || body.finalSectionProperties === undefined)
    return false;
  const section = body.sections.at(0);
  return (
    section !== undefined &&
    section.properties === body.finalSectionProperties &&
    Object.keys(section).length === 2 &&
    Object.hasOwn(section, "properties") &&
    Object.hasOwn(section, "content") &&
    section.content.length === body.content.length &&
    body.content.every((block, index) => block === section.content[index])
  );
};
