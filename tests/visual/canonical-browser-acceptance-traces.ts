import { BROWSER_PASTE_PAYLOADS, type BrowserInputTrace } from "./browserInputTrace";

/** Fixed native-input sequences supplement the generated canonical lane. */
export const canonicalBrowserAcceptances = [
  {
    seed: 29,
    trace: {
      shape: "tables",
      actions: [{ kind: "dragCellDelete" }, { kind: "typing", text: "alpha" }],
    },
  },
  {
    seed: 47,
    trace: {
      shape: "single-bullet-list",
      actions: [
        { kind: "pastePlain", plain: "alpha", html: "" },
        { kind: "pasteMultiBlock", ...BROWSER_PASTE_PAYLOADS.pasteMultiBlock },
        { kind: "selectionDrag", target: "list" },
        { kind: "backspace" },
      ],
    },
  },
  {
    seed: 131,
    trace: {
      shape: "header-footer",
      actions: [{ kind: "imeReplacement", updates: ["alpha"], completion: "cancel" }],
    },
  },
  {
    seed: 263,
    trace: {
      shape: "mixed-lists",
      actions: [
        { kind: "pasteHtml", ...BROWSER_PASTE_PAYLOADS.pasteHtml },
        { kind: "historyBurst", keys: ["undo", "undo"] },
        { kind: "selectionDrag", target: "list" },
      ],
    },
  },
  {
    seed: 347,
    trace: {
      shape: "mixed-lists",
      actions: [
        { kind: "selectionDrag", target: "list" },
        { kind: "imeReplacement", updates: ["alpha"], completion: "commit" },
      ],
    },
  },
  {
    seed: 431,
    trace: {
      shape: "single-decimal-list",
      actions: [
        { kind: "imeReplacement", updates: ["alpha"], completion: "cancel" },
        { kind: "redo" },
      ],
    },
  },
  {
    seed: 557,
    trace: {
      shape: "image",
      actions: [
        { kind: "imeReplacement", updates: ["alpha"], completion: "cancel" },
        { kind: "pasteWordHtml", ...BROWSER_PASTE_PAYLOADS.pasteWordHtml },
      ],
    },
  },
  {
    seed: 1791010163,
    trace: {
      shape: "notes",
      actions: [
        { kind: "imeReplacement", updates: ["alpha"], completion: "cancel" },
        { kind: "enter" },
      ],
    },
  },
] satisfies readonly { seed: number; trace: BrowserInputTrace }[];
