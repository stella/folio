import type { Style } from "../types/document";

type StyleReferenceLookup = { get: (styleId: string) => Style | undefined };

/** Unknown direct references carry no definition to import or apply. */
export const classifyStyleReference = (styleId: string, definitions: StyleReferenceLookup) => {
  const definition = definitions.get(styleId);
  return definition === undefined
    ? { kind: "unknown" as const }
    : { kind: "defined" as const, definition };
};
