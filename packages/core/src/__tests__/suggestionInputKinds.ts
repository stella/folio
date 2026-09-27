/** Declared editor inputs for the editing/suggesting equivalence oracle. */
export const SUGGESTION_INPUT_KINDS = [
  "typing",
  "enter",
  "backspace",
  "delete",
  "pastePlain",
  "pasteHtml",
  "pasteWordHtml",
  "pasteListHtml",
  "pasteTable",
  "pasteMultiBlock",
  "imeReplacement",
  "cut",
  "dragCellDelete",
] as const;

type SuggestionInputKind = (typeof SUGGESTION_INPUT_KINDS)[number];

export const SUGGESTION_INPUT_DRIVERS = {
  typing: { type: "conformance", operations: ["type:text", "type:text(mid)"] },
  enter: { type: "conformance", operations: ["key:Enter"] },
  backspace: { type: "conformance", operations: ["key:Backspace"] },
  delete: { type: "conformance", operations: ["key:Delete"] },
  pastePlain: { type: "conformance", operations: ["paste:plain"] },
  pasteHtml: { type: "browser" },
  pasteWordHtml: { type: "browser" },
  pasteListHtml: { type: "browser" },
  pasteTable: { type: "conformance", operations: ["paste:table"] },
  pasteMultiBlock: { type: "conformance", operations: ["paste:paragraphs", "paste:list"] },
  imeReplacement: { type: "browser" },
  cut: { type: "conformance", operations: ["host:cut"] },
  dragCellDelete: { type: "focused" },
} as const satisfies Record<
  SuggestionInputKind,
  { type: "conformance"; operations: readonly string[] } | { type: "browser" } | { type: "focused" }
>;
