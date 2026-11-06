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

export const BROWSER_INPUT_REGRESSIONS = [
  [
    {
      seed: 8,
      path: "full",
      fingerprint: "zzfull",
      trace: {
        shape: "single-bullet-list",
        actions: [
          { kind: "enter" },
          { kind: "selectionDrag", target: "list" },
          { kind: "delete" },
          {
            kind: "imeReplacement",
            updates: [
              "shall مرحبا alpha",
              "invoice valid",
              "valid",
              "shall café",
            ],
            completion: "cancel",
          },
          { kind: "pastePlain", plain: "é مرحبا invoice", html: "" },
          { kind: "cut" },
        ],
      },
    },
    {
      seed: 8,
      path: "nocut",
      fingerprint: "zznocut",
      trace: {
        shape: "single-bullet-list",
        actions: [
          { kind: "enter" },
          { kind: "selectionDrag", target: "list" },
          { kind: "delete" },
          {
            kind: "imeReplacement",
            updates: [
              "shall مرحبا alpha",
              "invoice valid",
              "valid",
              "shall café",
            ],
            completion: "cancel",
          },
          { kind: "pastePlain", plain: "é مرحبا invoice", html: "" },
        ],
      },
    },
    {
      seed: 8,
      path: "noenter",
      fingerprint: "zznoenter",
      trace: {
        shape: "single-bullet-list",
        actions: [
          { kind: "selectionDrag", target: "list" },
          { kind: "delete" },
          {
            kind: "imeReplacement",
            updates: [
              "shall مرحبا alpha",
              "invoice valid",
              "valid",
              "shall café",
            ],
            completion: "cancel",
          },
          { kind: "pastePlain", plain: "é مرحبا invoice", html: "" },
          { kind: "cut" },
        ],
      },
    },
    {
      seed: 8,
      path: "noime",
      fingerprint: "zznoime",
      trace: {
        shape: "single-bullet-list",
        actions: [
          { kind: "enter" },
          { kind: "selectionDrag", target: "list" },
          { kind: "delete" },
          { kind: "pastePlain", plain: "é مرحبا invoice", html: "" },
          { kind: "cut" },
        ],
      },
    },
    {
      seed: 8,
      path: "dragpaste",
      fingerprint: "zzdragpaste",
      trace: {
        shape: "single-bullet-list",
        actions: [
          { kind: "selectionDrag", target: "list" },
          { kind: "delete" },
          { kind: "pastePlain", plain: "é مرحبا invoice", html: "" },
        ],
      },
    },
    {
      seed: 8,
      path: "dragimepaste",
      fingerprint: "zzdragimepaste",
      trace: {
        shape: "single-bullet-list",
        actions: [
          { kind: "selectionDrag", target: "list" },
          { kind: "delete" },
          {
            kind: "imeReplacement",
            updates: [
              "shall مرحبا alpha",
              "invoice valid",
              "valid",
              "shall café",
            ],
            completion: "cancel",
          },
          { kind: "pastePlain", plain: "é مرحبا invoice", html: "" },
        ],
      },
    },
  ] as never[],
][0] as never[];
