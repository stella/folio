/** Seeded, shrinkable input traces shared by browser interaction specs. */

import fc from "fast-check";

import { SUGGESTION_INPUT_KINDS } from "../../packages/core/src/__tests__/suggestionInputKinds";

export const BROWSER_TRACE_FIXED_SEEDS = {
  pullRequest: [11, 29],
  nightly: [11, 29, 47, 83, 131, 197, 263, 347, 431, 557],
} as const;

export type BrowserDragTarget = "table" | "list" | "note" | "field" | "inlineObject";
export const BROWSER_SHAPES = {
  table: "tables",
  list: "mixed-lists",
  note: "notes",
  field: "fields-links-bookmarks",
  inlineObject: "image",
} as const satisfies Record<BrowserDragTarget, string>;
export type BrowserPasteKind = Extract<
  (typeof SUGGESTION_INPUT_KINDS)[number],
  "pastePlain" | "pasteHtml" | "pasteWordHtml" | "pasteListHtml" | "pasteTable" | "pasteMultiBlock"
>;

export type BrowserInputAction =
  | { kind: "typing"; text: string }
  | { kind: "enter" }
  | { kind: "backspace" }
  | { kind: "delete" }
  | { kind: BrowserPasteKind; plain: string; html: string }
  | { kind: "imeReplacement"; text: string }
  | { kind: "cut" }
  | { kind: "dragCellDelete" }
  | { kind: "undo" }
  | { kind: "redo" }
  | { kind: "selectionDrag"; target: BrowserDragTarget };

export type BrowserInputTrace = {
  shape: (typeof BROWSER_SHAPES)[BrowserDragTarget];
  actions: BrowserInputAction[];
};

const WORDS = ["alpha", "buyer", "shall", "valid", "invoice", "café", "東京"] as const;
const plainTextArbitrary = fc
  .array(fc.constantFrom(...WORDS), { minLength: 1, maxLength: 4 })
  .map((words) => words.join(" "));
const textArbitrary = fc
  .array(fc.constantFrom(...WORDS), { minLength: 1, maxLength: 3 })
  .map((words) => words.join(" "));
const pastePayload = {
  pasteHtml: {
    plain: "First bold\nSecond",
    html: "<p>First <strong>bold</strong></p><p>Second</p>",
  },
  pasteGoogleDocsHtml: {
    plain: "First bold\nSecond",
    html: '<b id="docs-internal-guid-folio"><p style="line-height:1.15"><span style="font-weight:700">First bold</span></p><p><span>Second</span></p></b>',
  },
  pasteWordHtml: {
    plain: "Opening\nClosing",
    html: '<p class="MsoNormal">Opening</p><p class="MsoNormal">Closing</p>',
  },
  pasteListHtml: {
    plain: "Alpha\nNested\nOmega",
    html: "<ul><li>Alpha<ul><li>Nested</li></ul></li><li>Omega</li></ul>",
  },
  pasteTable: {
    plain: "Left\tRight",
    html: "<table><tbody><tr><td>Left</td><td>Right</td></tr></tbody></table>",
  },
  pasteMultiBlock: { plain: "Title\nBody", html: "<p>Title</p><p>Body</p>" },
} as const;

const suggestionActionArbitraries = {
  typing: textArbitrary.map((text) => ({ kind: "typing", text }) as const),
  enter: fc.constant({ kind: "enter" } as const),
  backspace: fc.constant({ kind: "backspace" } as const),
  delete: fc.constant({ kind: "delete" } as const),
  pastePlain: plainTextArbitrary.map((plain) => ({ kind: "pastePlain", plain, html: "" }) as const),
  pasteHtml: fc.constantFrom(
    { kind: "pasteHtml", ...pastePayload.pasteHtml } as const,
    { kind: "pasteHtml", ...pastePayload.pasteGoogleDocsHtml } as const,
  ),
  pasteWordHtml: fc.constant({ kind: "pasteWordHtml", ...pastePayload.pasteWordHtml } as const),
  pasteListHtml: fc.constant({ kind: "pasteListHtml", ...pastePayload.pasteListHtml } as const),
  pasteTable: fc.constant({ kind: "pasteTable", ...pastePayload.pasteTable } as const),
  pasteMultiBlock: fc.constant({
    kind: "pasteMultiBlock",
    ...pastePayload.pasteMultiBlock,
  } as const),
  imeReplacement: textArbitrary.map((text) => ({ kind: "imeReplacement", text }) as const),
  cut: fc.constant({ kind: "cut" } as const),
  dragCellDelete: fc.constant({ kind: "dragCellDelete" } as const),
} satisfies Record<(typeof SUGGESTION_INPUT_KINDS)[number], fc.Arbitrary<BrowserInputAction>>;

export const browserSuggestionActionKinds = Object.keys(suggestionActionArbitraries);

const commonActionArbitraries: readonly fc.Arbitrary<BrowserInputAction>[] = [
  ...SUGGESTION_INPUT_KINDS.filter((kind) => kind !== "dragCellDelete").map(
    (kind) => suggestionActionArbitraries[kind],
  ),
  fc.constant({ kind: "undo" } as const),
  fc.constant({ kind: "redo" } as const),
];

const targetForShape = {
  tables: "table",
  "mixed-lists": "list",
  notes: "note",
  "fields-links-bookmarks": "field",
  image: "inlineObject",
} as const satisfies Record<(typeof BROWSER_SHAPES)[BrowserDragTarget], BrowserDragTarget>;

/** fast-check shrinks action sequences and their payloads to minimal failing traces. */
export const browserInputTraceArbitrary = (maxActions = 6): fc.Arbitrary<BrowserInputTrace> =>
  fc.constantFrom(...Object.values(BROWSER_SHAPES)).chain((shape) => {
    const actions = [
      ...commonActionArbitraries,
      fc.constant({ kind: "selectionDrag", target: targetForShape[shape] } as const),
    ];
    if (shape === BROWSER_SHAPES.table) actions.push(suggestionActionArbitraries.dragCellDelete);
    return fc.record({
      shape: fc.constant(shape),
      actions: fc.array(fc.oneof(...actions), { minLength: 1, maxLength: maxActions }),
    });
  });

export type BrowserInputTraceConfig = { seeds: readonly number[]; runs: number };

/** Read explicit PR/nightly overrides; defaults keep both lanes reproducible. */
export const parseBrowserInputTraceConfig = (
  env: Readonly<Record<string, string | undefined>>,
  lane: "pullRequest" | "nightly",
): BrowserInputTraceConfig => {
  const seedsText = env["FOLIO_FUZZ_SEEDS"];
  const runsText = env["FOLIO_FUZZ_RUNS"];
  const seeds =
    seedsText === undefined ? BROWSER_TRACE_FIXED_SEEDS[lane] : seedsText.split(",").map(Number);
  const defaultRuns = lane === "nightly" ? 20 : 2;
  const runs = runsText === undefined ? defaultRuns : Number(runsText);
  if (seeds.length === 0 || seeds.some((seed) => !Number.isSafeInteger(seed)))
    throw new Error("FOLIO_FUZZ_SEEDS must be comma-separated safe integers");
  if (!Number.isSafeInteger(runs) || runs < 1)
    throw new Error("FOLIO_FUZZ_RUNS must be a positive safe integer");
  return { seeds, runs };
};
