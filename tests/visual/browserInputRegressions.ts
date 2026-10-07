import { BROWSER_PASTE_PAYLOADS, type BrowserInputTrace } from "./browserInputTrace";

/** Exact findings stay as standing traces even when generator choices change. */
export const BROWSER_INPUT_REGRESSIONS = [
  {
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
  },
] satisfies { seed: number; path: string; fingerprint: string; trace: BrowserInputTrace }[];
