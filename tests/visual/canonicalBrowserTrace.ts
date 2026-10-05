import fc from "fast-check";
import { commonActionArbitraries } from "./browserInputTrace";

export const canonicalBrowserTraceArbitrary = fc.array(fc.oneof(...commonActionArbitraries), {
  minLength: 1,
  maxLength: 12,
});

export const CANONICAL_BROWSER_REDO_REPLAYS = [
  { seed: 29, path: "2:2:0:0:0:2", kinds: ["backspace", "undo"] },
  { seed: 47, path: "5:1:1", kinds: ["cut", "imeReplacement"] },
  { seed: 431, path: "9:1:0:1", kinds: ["cut", "pasteListHtml"] },
  { seed: 197, path: "1:1:2:2:2", kinds: ["typing", "pasteHtml"] },
] as const;

export const CANONICAL_BROWSER_HISTORY_REPLAYS = [
  {
    seed: 197,
    path: "1",
    kinds: [
      "imeReplacement",
      "pastePlain",
      "pastePlain",
      "pastePlain",
      "pasteListHtml",
      "typing",
      "imeReplacement",
      "undo",
      "redo",
      "pasteListHtml",
      "historyBurst",
      "pasteHtml",
    ],
  },
  { seed: 131, path: "8:1:0:0", kinds: ["backspace", "pasteWordHtml"] },
  { seed: 347, path: "0:0", kinds: ["enter"] },
  { seed: 29, path: "3:0", kinds: ["enter"] },
  { seed: 431, path: "5:2:4:6:6:6", kinds: ["typing", "historyBurst", "cut"] },
  { seed: 83, path: "1:4:1:1:1", kinds: ["backspace", "historyBurst"] },
  { seed: 557, path: "2:2:0:1", kinds: ["backspace", "pasteListHtml"] },
  ...CANONICAL_BROWSER_REDO_REPLAYS,
] as const;

export const CANONICAL_BROWSER_SAVE_REPLAYS = [
  { seed: 11, path: "2:1:1:0:1:1", kinds: ["pasteListHtml", "pastePlain"] },
  { seed: 83, path: "1:1:1:1:1:1:1", kinds: ["pasteListHtml", "historyBurst"] },
  { seed: 131, path: "3:1:2", kinds: ["pasteListHtml", "undo"] },
  { seed: 263, path: "2:1:0:0:0", kinds: ["pasteListHtml", "cut"] },
  { seed: 557, path: "1:1:0:0", kinds: ["pasteListHtml", "enter"] },
  { seed: 1791183757, path: "5:1:1", kinds: ["pasteListHtml", "pasteHtml"] },
] as const;
