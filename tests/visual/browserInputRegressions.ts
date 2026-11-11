import {
  BROWSER_PASTE_PAYLOADS,
  type BrowserInputTrace,
} from "./browserInputTrace";

export const PINNED_BROWSER_INPUT_TAG = "@pinned-browser-input";

/** Exact findings stay as standing traces even when generator choices change. */
export const BROWSER_NOTES_SEED_197 = {
  seed: 197,
  path: "1",
  fingerprint: "d859e264ce3610d3",
  trace: {
    shape: "notes",
    actions: [
      { kind: "pasteWordHtml", ...BROWSER_PASTE_PAYLOADS.pasteWordHtml },
      { kind: "cut" },
      { kind: "delete" },
      { kind: "backspace" },
      { kind: "pasteTable", ...BROWSER_PASTE_PAYLOADS.pasteTable },
      { kind: "selectionDrag", target: "note" },
    ],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

export const BROWSER_LIST_SEED_47 = {
  seed: 47,
  path: "0",
  fingerprint: "17a0f526509ab97b",
  trace: {
    shape: "single-bullet-list",
    actions: [
      { kind: "pastePlain", plain: "shall", html: "" },
      { kind: "cut" },
      {
        kind: "pasteMultiBlock",
        plain: "Title\nBody",
        html: "<p>Title</p><p>Body</p>",
      },
      { kind: "selectionDrag", target: "list" },
      { kind: "typing", text: "👩🏽‍⚖️ café" },
      { kind: "backspace" },
    ],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

export const BROWSER_TABLE_SEED_29 = {
  seed: 29,
  path: "1:0:3:1",
  fingerprint: "5f0168fbf69566cd",
  trace: {
    shape: "tables",
    actions: [{ kind: "dragCellDelete" }, { kind: "typing", text: "alpha" }],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** A composition after a tracked insertion leaves that insertion's marks alone. */
export const CANONICAL_INSERTION_SEED_557 = {
  authority: "canonical",
  seed: 557,
  path: "7:2:0:1:1",
  fingerprint: "b470db459c2eaf7d",
  trace: {
    shape: "header-footer",
    actions: [
      { kind: "pasteMultiBlock", ...BROWSER_PASTE_PAYLOADS.pasteMultiBlock },
      { kind: "imeReplacement", updates: ["alpha"], completion: "commit" },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** A composition over a selection across list items replaces it once, in the first item. */
export const CANONICAL_ACROSS_PARAGRAPHS_SEED_347 = {
  authority: "canonical",
  seed: 347,
  path: "17:2:1:1:1",
  fingerprint: "92b692379e90307a",
  trace: {
    shape: "mixed-lists",
    actions: [
      { kind: "selectionDrag", target: "list" },
      { kind: "imeReplacement", updates: ["alpha"], completion: "commit" },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** A composition in front of a tracked deletion leaves that deletion's marks alone. */
export const CANONICAL_DELETION_SEED_11 = {
  authority: "canonical",
  seed: 11,
  path: "13:5:3:3:3:3",
  fingerprint: "950f061a7bc5243d",
  trace: {
    shape: "rtl-cjk",
    actions: [
      { kind: "delete" },
      { kind: "imeReplacement", updates: ["alpha"], completion: "commit" },
      { kind: "pasteListHtml", ...BROWSER_PASTE_PAYLOADS.pasteListHtml },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** Pasting over selected cells refuses and leaves the document as it was. */
export const CANONICAL_CELL_PASTE_SEED_431 = {
  authority: "canonical",
  seed: 431,
  path: "9:0:1:0:0:1",
  fingerprint: "52b2a7dfe8a76372",
  trace: {
    shape: "tables",
    actions: [
      { kind: "selectionDrag", target: "table" },
      { kind: "pasteWordHtml", ...BROWSER_PASTE_PAYLOADS.pasteWordHtml },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** Cancelling a composition over selected cells keeps their text. */
export const CANONICAL_CELL_COMPOSITION_CANCEL = {
  authority: "canonical",
  seed: 431,
  path: "selected-cells",
  fingerprint: "composition-cancel",
  trace: {
    shape: "tables",
    actions: [
      { kind: "selectionDrag", target: "table" },
      { kind: "imeReplacement", updates: ["alpha"], completion: "cancel" },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** The browser restarts a composition beside a tracked insertion; its end names the text. */
export const CANONICAL_RESTARTED_COMPOSITION_SEED_47 = {
  authority: "canonical",
  seed: 47,
  path: "15:1:0:3:2:2:2:2:2",
  fingerprint: "8c764b826e4fb3b6",
  trace: {
    shape: "tables",
    actions: [
      { kind: "pasteWordHtml", ...BROWSER_PASTE_PAYLOADS.pasteWordHtml },
      {
        kind: "imeReplacement",
        updates: ["alpha", "alpha"],
        completion: "commit",
      },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** A restart that lands away from the caret still commits at the caret. */
export const CANONICAL_MISPLACED_COMPOSITION_SEED_29 = {
  authority: "canonical",
  seed: 29,
  path: "0",
  fingerprint: "b0ac2c02d75b820c",
  trace: {
    shape: "single-bullet-list",
    actions: [
      { kind: "undo" },
      { kind: "pasteWordHtml", ...BROWSER_PASTE_PAYLOADS.pasteWordHtml },
      { kind: "cut" },
      {
        kind: "imeReplacement",
        updates: [
          "buyer café",
          "valid shall",
          "é invoice café",
          "مرحبا buyer café",
        ],
        completion: "commit",
      },
      {
        kind: "imeReplacement",
        updates: ["valid alpha"],
        completion: "commit",
      },
      { kind: "cut" },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** A multi-update composition in front of a tracked deletion leaves it deleted. */
export const CANONICAL_DELETION_UPDATES_SEED_11 = {
  authority: "canonical",
  seed: 11,
  path: "13",
  fingerprint: "a70c021688cbc310",
  trace: {
    shape: "rtl-cjk",
    actions: [
      { kind: "delete" },
      { kind: "backspace" },
      {
        kind: "imeReplacement",
        updates: ["東京", "buyer", "invoice café", "buyer"],
        completion: "commit",
      },
      { kind: "undo" },
      { kind: "pasteListHtml", ...BROWSER_PASTE_PAYLOADS.pasteListHtml },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** A cancelled composition beside a bookmark keeps the bookmark. */
export const CANONICAL_BOOKMARK_SEED_263 = {
  authority: "canonical",
  seed: 263,
  path: "8:4:2:2:3:3:2",
  fingerprint: "2cf9ea677cc3c57c",
  trace: {
    shape: "fields-links-bookmarks",
    actions: [
      { kind: "imeReplacement", updates: ["alpha"], completion: "cancel" },
      { kind: "pastePlain", plain: "alpha", html: "" },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** A composition at a paragraph's start stays there, after an empty paragraph. */
export const CANONICAL_PARAGRAPH_START_SEED_263 = {
  authority: "canonical",
  seed: 263,
  path: "12:2:2:2:2",
  fingerprint: "156d4359b84edef4",
  trace: {
    shape: "image",
    actions: [
      { kind: "enter" },
      { kind: "imeReplacement", updates: ["alpha"], completion: "commit" },
      { kind: "undo" },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** A composition at a paragraph's start stays there, after a pasted list. */
export const CANONICAL_PARAGRAPH_START_SEED_431 = {
  authority: "canonical",
  seed: 431,
  path: "14:4:2:2:2",
  fingerprint: "76156fa33702d1c6",
  trace: {
    shape: "sections",
    actions: [
      { kind: "pasteListHtml", ...BROWSER_PASTE_PAYLOADS.pasteListHtml },
      { kind: "imeReplacement", updates: ["alpha"], completion: "commit" },
      { kind: "pasteWordHtml", ...BROWSER_PASTE_PAYLOADS.pasteWordHtml },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** A restarted composition after a tracked paste commits once, at the caret. */
export const CANONICAL_RESTARTED_AFTER_PASTE_SEED_197 = {
  authority: "canonical",
  seed: 197,
  path: "19",
  fingerprint: "855d1f9b6feaad45",
  trace: {
    shape: "mixed-lists",
    actions: [
      { kind: "typing", text: "valid invoice" },
      { kind: "pastePlain", plain: "مرحبا valid", html: "" },
      {
        kind: "imeReplacement",
        updates: ["invoice", "👩🏽‍⚖️ invoice", "café", "alpha buyer café"],
        completion: "commit",
      },
      { kind: "cut" },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/**
 * A composition that keeps going after the editor redrew its first text
 * beside a pasted run: the reread names the run's font for one script only.
 */
export const CANONICAL_COMPOSITION_REREAD_AFTER_PASTE = {
  authority: "canonical",
  seed: 197,
  path: "reread-after-paste",
  fingerprint: "composition-reread-after-paste",
  trace: {
    shape: "mixed-lists",
    actions: [
      { kind: "typing", text: "valid invoice" },
      { kind: "pastePlain", plain: "alpha valid", html: "" },
      {
        kind: "imeReplacement",
        updates: ["invoice", "👩🏽‍⚖️ invoice", "café", "alpha buyer café"],
        completion: "commit",
      },
      { kind: "cut" },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/**
 * A composition right after a pasted run, as an input method grows it: its
 * text in progress shows at the caret, and it commits what the default
 * editor commits.
 */
export const CANONICAL_COMPOSITION_BESIDE_PASTED_RUN = {
  authority: "canonical",
  seed: 197,
  path: "beside-pasted-run",
  fingerprint: "composition-beside-pasted-run",
  comparesWithDefaultEditor: true,
  trace: {
    shape: "mixed-lists",
    actions: [
      { kind: "typing", text: "valid invoice" },
      { kind: "pastePlain", plain: "alpha valid", html: "" },
      {
        kind: "imeReplacement",
        updates: ["in", "inv", "invoice"],
        completion: "commit",
      },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  comparesWithDefaultEditor: true;
  trace: BrowserInputTrace;
};

/** A restarted composition after tracked typing commits at the caret, not mid-word. */
export const CANONICAL_RESTARTED_AFTER_TYPING = {
  authority: "canonical",
  seed: 197,
  path: "typed-restart",
  fingerprint: "restarted-after-typing",
  trace: {
    shape: "mixed-lists",
    actions: [
      { kind: "typing", text: "valid invoice" },
      {
        kind: "imeReplacement",
        updates: ["invoice", "buyer", "alpha"],
        completion: "commit",
      },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** A composition inside a tracked run survives the DOM rebuilding its review mark. */
export const CANONICAL_TRACKED_RUN_REREAD_SEED_1793897521 = {
  authority: "canonical",
  seed: 1793897521,
  path: "12",
  fingerprint: "f9c052d5091f3dce",
  trace: {
    shape: "header-footer",
    actions: [
      { kind: "delete" },
      { kind: "pastePlain", plain: "東京", html: "" },
      {
        kind: "imeReplacement",
        updates: ["café alpha", "👩🏽‍⚖️ buyer"],
        completion: "cancel",
      },
      { kind: "historyBurst", keys: ["redo", "redo"] },
      { kind: "backspace" },
      { kind: "cut" },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

export const BROWSER_DELETE_BESIDE_BOOKMARK_SEED_83 = {
  seed: 83,
  path: "3",
  fingerprint: "0ae2a73b64015980",
  trace: {
    shape: "fields-links-bookmarks",
    actions: [{ kind: "delete" }, { kind: "enter" }],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

export const BROWSER_CANCELLED_COMPOSITION_SEED_431 = {
  seed: 431,
  path: "9:2:1:1:1:2:2",
  fingerprint: "16b45becab03610e",
  trace: {
    shape: "single-decimal-list",
    actions: [
      {
        kind: "imeReplacement",
        updates: ["alpha"],
        completion: "cancel",
      },
      {
        kind: "redo",
      },
    ],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

export const BROWSER_CANCELLED_COMPOSITION_SEED_263 = {
  seed: 263,
  path: "8",
  fingerprint: "f833d2a2c838868e",
  trace: {
    shape: "fields-links-bookmarks",
    actions: [
      {
        kind: "imeReplacement",
        updates: ["valid", "invoice alpha shall", "buyer مرحبا é"],
        completion: "cancel",
      },
      {
        kind: "delete",
      },
      {
        kind: "undo",
      },
      {
        kind: "pasteTable",
        plain: "Left\tRight",
        html: "<table><tbody><tr><td>Left</td><td>Right</td></tr></tbody></table>",
      },
      {
        kind: "redo",
      },
      {
        kind: "pastePlain",
        plain: "café",
        html: "",
      },
    ],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

export const BROWSER_CANCELLED_COMPOSITION_SEED_1793897521 = {
  seed: 1793897521,
  path: "11",
  fingerprint: "6b03aac751613037",
  trace: {
    shape: "notes",
    actions: [
      {
        kind: "pasteTable",
        plain: "Left\tRight",
        html: "<table><tbody><tr><td>Left</td><td>Right</td></tr></tbody></table>",
      },
      {
        kind: "imeReplacement",
        updates: ["é 東京", "buyer", "invoice alpha alpha"],
        completion: "cancel",
      },
      {
        kind: "imeReplacement",
        updates: ["valid buyer"],
        completion: "commit",
      },
      {
        kind: "pasteMultiBlock",
        plain: "Title\nBody",
        html: "<p>Title</p><p>Body</p>",
      },
    ],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

export const BROWSER_CANCELLED_COMPOSITION_SEED_131 = {
  seed: 131,
  path: "5",
  fingerprint: "4ae472e6389a2494",
  trace: {
    shape: "header-footer",
    actions: [
      { kind: "historyBurst", keys: ["undo", "redo", "redo", "undo"] },
      { kind: "redo" },
      { kind: "redo" },
      { kind: "backspace" },
      { kind: "typing", text: "café buyer" },
      {
        kind: "imeReplacement",
        updates: ["shall alpha", "buyer alpha 東京", "東京 مرحبا", "é buyer"],
        completion: "cancel",
      },
    ],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

export const BROWSER_PASTE_AFTER_TABLE_SEED_29 = {
  seed: 29,
  path: "1:3:3",
  fingerprint: "44acba1a39ae6967",
  trace: {
    shape: "single-decimal-list",
    actions: [
      { kind: "selectionDrag", target: "list" },
      {
        kind: "pasteTable",
        plain: "Left\tRight",
        html: "<table><tbody><tr><td>Left</td><td>Right</td></tr></tbody></table>",
      },
      {
        kind: "pasteHtml",
        plain: "First bold\nSecond",
        html: "<p>First <strong>bold</strong></p><p>Second</p>",
      },
    ],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

export const BROWSER_PASTE_BEFORE_STRUCK_SEED_431 = {
  seed: 431,
  path: "9",
  fingerprint: "01f3a41c8510892a",
  trace: {
    shape: "single-decimal-list",
    actions: [
      { kind: "selectionDrag", target: "list" },
      {
        kind: "pasteListHtml",
        plain: "Alpha\nNested\nOmega",
        html: "<ul><li>Alpha<ul><li>Nested</li></ul></li><li>Omega</li></ul>",
      },
      {
        kind: "imeReplacement",
        updates: ["buyer", "buyer shall"],
        completion: "cancel",
      },
      {
        kind: "pasteTable",
        plain: "Left\tRight",
        html: "<table><tbody><tr><td>Left</td><td>Right</td></tr></tbody></table>",
      },
      {
        kind: "pasteListHtml",
        plain: "Alpha\nNested\nOmega",
        html: "<ul><li>Alpha<ul><li>Nested</li></ul></li><li>Omega</li></ul>",
      },
      { kind: "redo" },
    ],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

export const BROWSER_SECOND_CELL_PASTE_SEED_47 = {
  seed: 47,
  path: "14",
  fingerprint: "4e71fdf9e51b2500",
  trace: {
    shape: "tables",
    actions: [
      { kind: "typing", text: "alpha" },
      { kind: "selectionDrag", target: "table" },
      { kind: "delete" },
      {
        kind: "pasteWordHtml",
        plain: "Opening\nClosing",
        html: '<html xmlns:o="urn:schemas-microsoft-com:office:office"><body><!--StartFragment--><p class="MsoNormal" style="margin:0;font-size:11pt">Opening<o:p></o:p></p><p class="MsoNormal" style="margin-top:0">Closing<o:p>&nbsp;</o:p></p><!--EndFragment--></body></html>',
      },
      {
        kind: "pasteHtml",
        plain: "First bold\nSecond",
        html: '<meta charset="utf-8"><b id="docs-internal-guid-folio"><p dir="ltr" style="line-height:1.15;margin:0"><span style="font-size:11pt;font-weight:700;white-space:pre-wrap">First bold</span></p><p><span style="white-space:pre-wrap">Second</span></p></b>',
      },
      {
        kind: "pasteTable",
        plain: "Left\tRight",
        html: "<table><tbody><tr><td>Left</td><td>Right</td></tr></tbody></table>",
      },
    ],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

export const BROWSER_RTL_COMPOSITION_SEED_1793914965 = {
  seed: 1793914965,
  path: "7:2:1:2:2",
  fingerprint: "5736502082eb3a9b",
  trace: {
    shape: "single-bullet-list",
    actions: [
      { kind: "imeReplacement", updates: ["مرحبا"], completion: "commit" },
      {
        kind: "pasteHtml",
        plain: "First bold\nSecond",
        html: "<p>First <strong>bold</strong></p><p>Second</p>",
      },
    ],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

export const BROWSER_PASTE_BEHIND_BOOKMARK_SEED_1793921366 = {
  seed: 1793921366,
  path: "11:4",
  fingerprint: "62177b2dda921e89",
  trace: {
    shape: "fields-links-bookmarks",
    actions: [
      { kind: "delete" },
      {
        kind: "pasteListHtml",
        plain: "Alpha\nNested\nOmega",
        html: "<ul><li>Alpha<ul><li>Nested</li></ul></li><li>Omega</li></ul>",
      },
    ],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/**
 * The marker kinds the bookmark traces do not reach, through the actions that
 * found a defect beside a bookmark: deleting, then pasting behind what went;
 * typing a break; cancelling a composition.
 */
const INVISIBLE_MARKER_ACTIONS: readonly BrowserInputTrace["actions"][] = [
  [
    { kind: "delete" },
    {
      kind: "pasteListHtml",
      plain: "Alpha\nNested\nOmega",
      html: "<ul><li>Alpha<ul><li>Nested</li></ul></li><li>Omega</li></ul>",
    },
  ],
  [{ kind: "delete" }, { kind: "pastePlain", plain: "Pasted", html: "" }],
  [{ kind: "enter" }, { kind: "delete" }],
  [
    { kind: "delete" },
    {
      kind: "imeReplacement",
      updates: ["shall", "shall café"],
      completion: "cancel",
    },
  ],
];

const BROWSER_INVISIBLE_MARKER_TRACES = (
  ["comments", "invisible-markers"] as const
).flatMap((shape) =>
  INVISIBLE_MARKER_ACTIONS.map((actions, index) => ({
    seed: 0,
    path: `${shape}-${String(index)}`,
    fingerprint: "pinned-for-its-markers",
    trace: { shape, actions },
  })),
);

export const BROWSER_CANCELLED_COMPOSITION_SEED_1793946574 = {
  seed: 1793946574,
  path: "10",
  fingerprint: "caed1eca237779c4",
  trace: {
    shape: "single-bullet-list",
    actions: [
      { kind: "enter" },
      { kind: "selectionDrag", target: "list" },
      { kind: "delete" },
      {
        kind: "imeReplacement",
        updates: ["shall مرحبا alpha", "invoice valid", "valid", "shall café"],
        completion: "cancel",
      },
      { kind: "pastePlain", plain: "é مرحبا invoice", html: "" },
      { kind: "cut" },
    ],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** Backspace before tracked-deleted text at a paragraph's start joins the paragraph. */
export const CANONICAL_JOIN_BESIDE_DELETION_SEED_29 = {
  authority: "canonical",
  seed: 29,
  path: "15:2",
  fingerprint: "bad194d7a1bfd1c6",
  trace: {
    shape: "image",
    actions: [
      { kind: "selectionDrag", target: "inlineObject" },
      { kind: "pasteListHtml", ...BROWSER_PASTE_PAYLOADS.pasteListHtml },
      { kind: "backspace" },
      { kind: "pasteMultiBlock", ...BROWSER_PASTE_PAYLOADS.pasteMultiBlock },
    ],
  },
} satisfies {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

type CanonicalRegression = {
  authority: "canonical";
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** ProseMirror types over a selection across paragraphs from keypress, with no beforeinput. */
export const CANONICAL_TYPING_OVER_LIST_ITEMS_SEED_431 = {
  authority: "canonical",
  seed: 431,
  path: "14:2:3:2:2:1:1",
  fingerprint: "d80b63ac597ad07a",
  trace: {
    shape: "single-decimal-list",
    actions: [
      { kind: "selectionDrag", target: "list" },
      { kind: "typing", text: "alpha" },
    ],
  },
} satisfies CanonicalRegression;

export const CANONICAL_TYPING_OVER_LIST_ITEMS_AFTER_TYPING_SEED_431 = {
  authority: "canonical",
  seed: 431,
  path: "14",
  fingerprint: "1d751358bf327e40",
  trace: {
    shape: "single-decimal-list",
    actions: [
      { kind: "typing", text: "shall alpha" },
      { kind: "selectionDrag", target: "list" },
      { kind: "typing", text: "é 東京 東京" },
    ],
  },
} satisfies CanonicalRegression;

export const CANONICAL_TYPING_ACROSS_PARAGRAPHS = {
  authority: "canonical",
  seed: 0,
  path: "paragraphs",
  fingerprint: "typing-across-paragraphs",
  trace: {
    shape: "plain-markdown",
    actions: [
      { kind: "selectionDrag", target: "paragraphs" },
      { kind: "typing", text: "shall" },
    ],
  },
} satisfies CanonicalRegression;

export const CANONICAL_TYPING_OVER_EVERYTHING = {
  authority: "canonical",
  seed: 0,
  path: "selectAll",
  fingerprint: "typing-over-select-all",
  trace: {
    shape: "plain-markdown",
    actions: [{ kind: "selectAll" }, { kind: "typing", text: "shall" }],
  },
} satisfies CanonicalRegression;

/** Typing over selected cells refuses with the session's own reason, leaving them as they were. */
export const CANONICAL_TYPING_OVER_CELLS = {
  authority: "canonical",
  seed: 0,
  path: "cells",
  fingerprint: "typing-over-selected-cells",
  trace: {
    shape: "tables",
    actions: [
      { kind: "selectionDrag", target: "table" },
      { kind: "typing", text: "shall" },
    ],
  },
} satisfies CanonicalRegression;

const SELECT_ALL = "ControlOrMeta+a";
const SELECT_DOWN = "Shift+ArrowDown";

/** A key replacing a selection across paragraphs. */
const BROWSER_KEY_OVER_SELECTION_TRACES = (
  [
    ["plain-markdown", [SELECT_ALL, "Tab"]],
    ["plain-markdown", [SELECT_DOWN, "Tab"]],
    ["sections", [SELECT_ALL, "Shift+Enter"]],
    ["notes", [SELECT_DOWN, "Shift+Enter"]],
  ] as const
).map(([shape, keys], index) => ({
  seed: 0,
  path: `${shape}-${String(index)}`,
  fingerprint: "pinned-for-a-key-over-a-selection",
  trace: { shape, actions: [{ kind: "keys", keys: [...keys] }] },
})) satisfies readonly {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
}[];

/** Ctrl, Cmd or Shift with Backspace at the start of a list item joins it to the paragraph before. */
const BROWSER_MODIFIED_BACKSPACE_TRACES = (
  [
    ["single-bullet-list", "ControlOrMeta+Backspace"],
    ["single-bullet-list", "Shift+Backspace"],
    ["single-decimal-list", "ControlOrMeta+Backspace"],
  ] as const
).map(([shape, key], index) => ({
  seed: 0,
  path: `${shape}-${String(index)}`,
  fingerprint: "pinned-for-a-modified-backspace",
  trace: {
    shape,
    actions: [
      {
        kind: "keys",
        keys: ["ControlOrMeta+Home", "ArrowDown", "ArrowDown", "Home", key],
      },
    ],
  },
})) satisfies readonly {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
}[];

/** A list pasted after cutting a block of table cells, while suggesting. */
export const BROWSER_TABLE_CUT_PASTE_SEED_1794032767 = {
  seed: 1794032767,
  path: "13",
  fingerprint: "407e8612a9687fad",
  trace: {
    shape: "tables",
    actions: [
      { kind: "pasteMultiBlock", ...BROWSER_PASTE_PAYLOADS.pasteMultiBlock },
      { kind: "dragCellDelete" },
      { kind: "cut" },
      { kind: "pasteListHtml", ...BROWSER_PASTE_PAYLOADS.pasteListHtml },
      { kind: "selectionDrag", target: "table" },
    ],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** A list pasted over words running to a paragraph's end, while suggesting. */
export const BROWSER_NOTES_LIST_PASTE_SEED_29 = {
  seed: 29,
  path: "15:2",
  fingerprint: "8d01cc728c151a9c",
  trace: {
    shape: "notes",
    actions: [
      { kind: "selectionDrag", target: "note" },
      { kind: "pasteListHtml", ...BROWSER_PASTE_PAYLOADS.pasteListHtml },
      { kind: "backspace" },
      { kind: "pasteMultiBlock", ...BROWSER_PASTE_PAYLOADS.pasteMultiBlock },
    ],
  },
} satisfies {
  seed: number;
  path: string;
  fingerprint: string;
  trace: BrowserInputTrace;
};

/** A composition right after a list pasted before a bookmarked heading. */
export const CANONICAL_COMPOSITION_AFTER_LIST_PASTE_SEED_11 = {
  authority: "canonical",
  seed: 11,
  path: "3:2:1:1",
  fingerprint: "9b669b5c44effff2",
  trace: {
    shape: "fields-links-bookmarks",
    actions: [
      { kind: "pasteListHtml", ...BROWSER_PASTE_PAYLOADS.pasteListHtml },
      { kind: "imeReplacement", updates: ["alpha"], completion: "commit" },
    ],
  },
} satisfies CanonicalRegression;

/** Everything cut while suggesting, across a comment's range and its reference. */
export const CANONICAL_CUT_EVERYTHING_BESIDE_COMMENT_SEED_431 = {
  authority: "canonical",
  seed: 431,
  path: "4:3:1",
  fingerprint: "5713f8fb706927fd",
  trace: {
    shape: "comments",
    actions: [{ kind: "selectAll" }, { kind: "cut" }],
  },
} satisfies CanonicalRegression;

/** A table pasted where everything around a field, a link and a bookmark was deleted. */
export const CANONICAL_TABLE_PASTE_OVER_EMPTIED_FIELDS_SEED_11 = {
  authority: "canonical",
  seed: 11,
  path: "2:3:2:2:2",
  fingerprint: "234f75ae8b2de9d4",
  trace: {
    shape: "fields-links-bookmarks",
    actions: [
      { kind: "selectAll" },
      { kind: "delete" },
      { kind: "pasteTable", ...BROWSER_PASTE_PAYLOADS.pasteTable },
    ],
  },
} satisfies CanonicalRegression;

export const BROWSER_INPUT_REGRESSIONS = [
  ...BROWSER_INVISIBLE_MARKER_TRACES,
  ...BROWSER_MODIFIED_BACKSPACE_TRACES,
  ...BROWSER_KEY_OVER_SELECTION_TRACES,
  BROWSER_NOTES_SEED_197,
  BROWSER_LIST_SEED_47,
  BROWSER_TABLE_SEED_29,
  BROWSER_TABLE_CUT_PASTE_SEED_1794032767,
  BROWSER_CANCELLED_COMPOSITION_SEED_1793946574,
  BROWSER_PASTE_BEHIND_BOOKMARK_SEED_1793921366,
  BROWSER_RTL_COMPOSITION_SEED_1793914965,
  BROWSER_PASTE_AFTER_TABLE_SEED_29,
  BROWSER_PASTE_BEFORE_STRUCK_SEED_431,
  BROWSER_SECOND_CELL_PASTE_SEED_47,
  BROWSER_CANCELLED_COMPOSITION_SEED_431,
  BROWSER_CANCELLED_COMPOSITION_SEED_263,
  BROWSER_CANCELLED_COMPOSITION_SEED_1793897521,
  BROWSER_CANCELLED_COMPOSITION_SEED_131,
  BROWSER_DELETE_BESIDE_BOOKMARK_SEED_83,
  CANONICAL_INSERTION_SEED_557,
  CANONICAL_ACROSS_PARAGRAPHS_SEED_347,
  CANONICAL_DELETION_SEED_11,
  CANONICAL_CELL_PASTE_SEED_431,
  CANONICAL_CELL_COMPOSITION_CANCEL,
  CANONICAL_RESTARTED_COMPOSITION_SEED_47,
  CANONICAL_MISPLACED_COMPOSITION_SEED_29,
  CANONICAL_DELETION_UPDATES_SEED_11,
  CANONICAL_BOOKMARK_SEED_263,
  CANONICAL_PARAGRAPH_START_SEED_263,
  CANONICAL_PARAGRAPH_START_SEED_431,
  CANONICAL_RESTARTED_AFTER_PASTE_SEED_197,
  CANONICAL_RESTARTED_AFTER_TYPING,
  CANONICAL_COMPOSITION_REREAD_AFTER_PASTE,
  CANONICAL_COMPOSITION_BESIDE_PASTED_RUN,
  CANONICAL_TRACKED_RUN_REREAD_SEED_1793897521,
  CANONICAL_JOIN_BESIDE_DELETION_SEED_29,
  CANONICAL_TYPING_OVER_LIST_ITEMS_SEED_431,
  CANONICAL_TYPING_OVER_LIST_ITEMS_AFTER_TYPING_SEED_431,
  CANONICAL_TYPING_ACROSS_PARAGRAPHS,
  CANONICAL_TYPING_OVER_EVERYTHING,
  CANONICAL_TYPING_OVER_CELLS,
  BROWSER_NOTES_LIST_PASTE_SEED_29,
  CANONICAL_COMPOSITION_AFTER_LIST_PASTE_SEED_11,
  CANONICAL_TABLE_PASTE_OVER_EMPTIED_FIELDS_SEED_11,
  CANONICAL_CUT_EVERYTHING_BESIDE_COMMENT_SEED_431,
];
