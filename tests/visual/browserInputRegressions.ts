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

export const BROWSER_INPUT_REGRESSIONS = [
  BROWSER_NOTES_SEED_197,
  BROWSER_LIST_SEED_47,
  BROWSER_TABLE_SEED_29,
  BROWSER_PASTE_AFTER_TABLE_SEED_29,
  BROWSER_PASTE_BEFORE_STRUCK_SEED_431,
  BROWSER_SECOND_CELL_PASTE_SEED_47,
  CANONICAL_INSERTION_SEED_557,
  CANONICAL_DELETION_SEED_11,
  CANONICAL_CELL_PASTE_SEED_431,
  CANONICAL_CELL_COMPOSITION_CANCEL,
  CANONICAL_RESTARTED_COMPOSITION_SEED_47,
  CANONICAL_MISPLACED_COMPOSITION_SEED_29,
  CANONICAL_DELETION_UPDATES_SEED_11,
  CANONICAL_BOOKMARK_SEED_263,
];
