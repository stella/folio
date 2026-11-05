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

export const BROWSER_INPUT_REGRESSIONS = [
  BROWSER_NOTES_SEED_197,
  BROWSER_LIST_SEED_47,
  BROWSER_TABLE_SEED_29,
  CANONICAL_INSERTION_SEED_557,
  CANONICAL_DELETION_SEED_11,
  CANONICAL_CELL_PASTE_SEED_431,
  CANONICAL_CELL_COMPOSITION_CANCEL,
];
