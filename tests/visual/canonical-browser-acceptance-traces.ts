import { BROWSER_PASTE_PAYLOADS, type BrowserInputTrace } from "./browserInputTrace";
import type { CanonicalSessionError } from "../../packages/core/src/controller/canonicalSession";
import { CANONICAL_GAP } from "../../packages/core/src/types/canonicalCapabilities";

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

type AcceptanceSourceShape = (typeof canonicalBrowserAcceptances)[number]["trace"]["shape"];
type SourceRefusal = Pick<CanonicalSessionError, "name" | "gap" | "reason" | "message">;

const unsupportedInlineSource = {
  name: "CanonicalSessionError",
  gap: CANONICAL_GAP.dispatch,
  reason: "refused",
  message: "Canonical sessions currently require plain paragraphs and supported inline atoms.",
} as const satisfies SourceRefusal;

/** Expected source limits apply only on refusal; newly supported sources run their traces. */
export const canonicalBrowserSourceRefusals = {
  tables: {
    name: "CanonicalSessionError",
    gap: CANONICAL_GAP.tableActivation,
    reason: "refused",
    message: "Canonical sessions cannot activate documents containing tables.",
  },
  "single-bullet-list": unsupportedInlineSource,
  "header-footer": unsupportedInlineSource,
  "mixed-lists": unsupportedInlineSource,
  "single-decimal-list": unsupportedInlineSource,
  image: unsupportedInlineSource,
  notes: unsupportedInlineSource,
} as const satisfies Record<AcceptanceSourceShape, SourceRefusal>;
