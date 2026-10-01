/** Seeded, shrinkable input traces shared by browser interaction specs. */

import fc from "fast-check";

import { SUGGESTION_INPUT_KINDS } from "../../packages/core/src/__tests__/suggestionInputKinds";

export const BROWSER_TRACE_FIXED_SEEDS = {
  pullRequest: [11, 29],
  nightly: [11, 29, 47, 83, 131, 197, 263, 347, 431, 557],
} as const;

export type BrowserDragTarget = "table" | "list" | "note" | "field" | "inlineObject";
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
  | { kind: "imeReplacement"; updates: string[]; completion: "commit" | "cancel" }
  | { kind: "cut" }
  | { kind: "dragCellDelete" }
  | { kind: "undo" }
  | { kind: "redo" }
  | { kind: "historyBurst"; keys: ("undo" | "redo")[] }
  | { kind: "selectionDrag"; target: BrowserDragTarget };

export type BrowserInputTrace = {
  shape: keyof typeof BROWSER_SHAPE_TARGETS;
  actions: BrowserInputAction[];
};

const WORDS = [
  "alpha",
  "buyer",
  "shall",
  "valid",
  "invoice",
  "café",
  "東京",
  "e\u0301",
  "👩🏽‍⚖️",
  "مرحبا",
] as const;
const plainTextArbitrary = fc
  .array(fc.constantFrom(...WORDS), { minLength: 1, maxLength: 4 })
  .map((words) => words.join(" "));
const textArbitrary = fc
  .array(fc.constantFrom(...WORDS), { minLength: 1, maxLength: 3 })
  .map((words) => words.join(" "));
export const BROWSER_PASTE_PAYLOADS = {
  pasteHtml: {
    plain: "First bold\nSecond",
    html: "<p>First <strong>bold</strong></p><p>Second</p>",
  },
  pasteWebAppHtml: {
    plain: "First bold\nSecond",
    html: '<meta charset="utf-8"><b id="docs-internal-guid-folio"><p dir="ltr" style="line-height:1.15;margin:0"><span style="font-size:11pt;font-weight:700;white-space:pre-wrap">First bold</span></p><p><span style="white-space:pre-wrap">Second</span></p></b>',
  },
  pasteWordHtml: {
    plain: "Opening\nClosing",
    html: '<html xmlns:o="urn:schemas-microsoft-com:office:office"><body><!--StartFragment--><p class="MsoNormal" style="margin:0;font-size:11pt">Opening<o:p></o:p></p><p class="MsoNormal" style="margin-top:0">Closing<o:p>&nbsp;</o:p></p><!--EndFragment--></body></html>',
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

/** Updates and completion shrink independently, preserving a valid lifecycle. */
export const browserImeActionArbitrary = fc.record({
  kind: fc.constant("imeReplacement"),
  updates: fc.array(textArbitrary, { minLength: 1, maxLength: 4 }),
  completion: fc.constantFrom("commit", "cancel"),
}) satisfies fc.Arbitrary<BrowserInputAction>;

const suggestionActionArbitraries = {
  typing: textArbitrary.map((text) => ({ kind: "typing", text }) as const),
  enter: fc.constant({ kind: "enter" } as const),
  backspace: fc.constant({ kind: "backspace" } as const),
  delete: fc.constant({ kind: "delete" } as const),
  pastePlain: plainTextArbitrary.map((plain) => ({ kind: "pastePlain", plain, html: "" }) as const),
  pasteHtml: fc.constantFrom(
    { kind: "pasteHtml", ...BROWSER_PASTE_PAYLOADS.pasteHtml } as const,
    { kind: "pasteHtml", ...BROWSER_PASTE_PAYLOADS.pasteWebAppHtml } as const,
  ),
  pasteWordHtml: fc.constant({
    kind: "pasteWordHtml",
    ...BROWSER_PASTE_PAYLOADS.pasteWordHtml,
  } as const),
  pasteListHtml: fc.constant({
    kind: "pasteListHtml",
    ...BROWSER_PASTE_PAYLOADS.pasteListHtml,
  } as const),
  pasteTable: fc.constant({ kind: "pasteTable", ...BROWSER_PASTE_PAYLOADS.pasteTable } as const),
  pasteMultiBlock: fc.constant({
    kind: "pasteMultiBlock",
    ...BROWSER_PASTE_PAYLOADS.pasteMultiBlock,
  } as const),
  imeReplacement: browserImeActionArbitrary,
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
  fc.record({
    kind: fc.constant("historyBurst"),
    keys: fc.array(fc.constantFrom("undo", "redo"), { minLength: 2, maxLength: 6 }),
  }),
];

/** A shape declares only targets it actually contains; text-only fixtures omit drags. */
export const BROWSER_SHAPE_TARGETS = {
  tables: "table",
  "mixed-lists": "list",
  "single-decimal-list": "list",
  "single-bullet-list": "list",
  notes: "note",
  "fields-links-bookmarks": "field",
  image: "inlineObject",
  "bare-package": null,
  "plain-markdown": null,
  "rtl-cjk": null,
  sections: null,
  "header-footer": null,
} as const satisfies Record<string, BrowserDragTarget | null>;

const isBrowserShape = (shape: string): shape is keyof typeof BROWSER_SHAPE_TARGETS =>
  Object.hasOwn(BROWSER_SHAPE_TARGETS, shape);
const shapeArbitrary = fc.constantFrom(
  ...Object.keys(BROWSER_SHAPE_TARGETS).filter(isBrowserShape),
);

/** fast-check shrinks action sequences and their payloads to minimal failing traces. */
export const browserInputTraceArbitrary = (maxActions = 6): fc.Arbitrary<BrowserInputTrace> =>
  shapeArbitrary.chain((shape) => {
    const actions = [...commonActionArbitraries];
    const target = BROWSER_SHAPE_TARGETS[shape];
    if (target !== null) actions.push(fc.constant({ kind: "selectionDrag", target } as const));
    if (target === "table") actions.push(suggestionActionArbitraries.dragCellDelete);
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
  if (
    seedsText?.split(",").some((seed) => seed.trim().length === 0) ||
    seeds.length === 0 ||
    seeds.some((seed) => !Number.isSafeInteger(seed))
  )
    throw new Error("FOLIO_FUZZ_SEEDS must be comma-separated safe integers");
  if (!Number.isSafeInteger(runs) || runs < 1)
    throw new Error("FOLIO_FUZZ_RUNS must be a positive safe integer");
  return { seeds, runs };
};
