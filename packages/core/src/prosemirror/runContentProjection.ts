import type { RunContent } from "../types/document";

/** RunContent cases and the inline nodes their projection may emit. */
const RUN_CONTENT_PROJECTION = {
  text: ["text"],
  tab: ["tab"],
  break: ["hardBreak", "pageBreakRun"],
  symbol: ["symbol"],
  footnoteRef: ["text"],
  endnoteRef: ["text"],
  noteMarker: ["noteMarker"],
  fieldChar: [],
  instrText: [],
  softHyphen: ["text"],
  noBreakHyphen: ["text"],
  renderedPageBreak: ["renderedPageBreak"],
  preservedXml: ["preservedXml"],
  drawing: ["image"],
  shape: ["shape", "textBoxAnchor"],
} as const satisfies Record<RunContent["type"], readonly string[]>;

// Consumers need names, not a literal union whose declaration order can vary across builds.
export const RUN_CONTENT_INLINE_NODE_NAMES: readonly string[] = [
  ...new Set(Object.values(RUN_CONTENT_PROJECTION).flat()),
];

export const isRunContentInlineNodeName = (name: string): boolean =>
  RUN_CONTENT_INLINE_NODE_NAMES.some((candidate) => candidate === name);

export const isRunContentProjectionNode = (
  sourceType: RunContent["type"],
  nodeName: string,
): boolean =>
  RUN_CONTENT_PROJECTION[sourceType].some((name) => name === nodeName);
