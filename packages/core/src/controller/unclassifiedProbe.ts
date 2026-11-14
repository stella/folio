// Scratch instrumentation: report every unclassified-transaction refusal with
// the input and the transactions that reached the projection boundary.
import type { Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

let lastInput: string | undefined;
let context: unknown;

const probed = new WeakSet<EditorView>();

/** Remember the latest DOM input a view saw, so a refusal can name it. */
export const probeInput = (view: EditorView): void => {
  if (probed.has(view)) return;
  probed.add(view);
  const record = (value: string) => {
    lastInput = value;
  };
  view.dom.addEventListener(
    "beforeinput",
    (event) => record(`beforeinput:${(event as InputEvent).inputType}`),
    true,
  );
  view.dom.addEventListener(
    "keydown",
    (event) => {
      const key = event as KeyboardEvent;
      record(
        `keydown:${key.ctrlKey ? "Ctrl-" : ""}${key.metaKey ? "Mod-" : ""}${key.altKey ? "Alt-" : ""}${key.shiftKey ? "Shift-" : ""}${key.key}`,
      );
    },
    true,
  );
  for (const type of ["paste", "drop", "cut", "compositionstart"])
    view.dom.addEventListener(type, () => record(type), true);
};

const sliceSummary = (json: unknown): string => {
  if (typeof json !== "object" || json === null) return "";
  const slice = Reflect.get(json, "slice");
  if (typeof slice !== "object" || slice === null) return "";
  const content = Reflect.get(slice, "content");
  if (!Array.isArray(content)) return "";
  return content
    .map((node: unknown) =>
      typeof node === "object" && node !== null
        ? String(Reflect.get(node, "type"))
        : "?",
    )
    .join("+");
};

export const describeTransaction = (transaction: Transaction) => {
  const meta = Reflect.get(transaction, "meta");
  return {
    docChanged: transaction.docChanged,
    meta:
      typeof meta === "object" && meta !== null
        ? Object.keys(meta).toSorted()
        : [],
    uiEvent: transaction.getMeta("uiEvent") ?? null,
    steps: transaction.steps.map((step) => {
      const json: unknown = step.toJSON();
      const type =
        typeof json === "object" && json !== null
          ? String(Reflect.get(json, "stepType"))
          : "?";
      const summary = sliceSummary(json);
      return summary === "" ? type : `${type}:${summary}`;
    }),
  };
};

export const probeContext = (value: unknown): void => {
  context = value;
};

export const reportUnclassified = (site: string, message: string): void => {
  const stack = new Error("probe").stack?.split("\n").slice(2, 10) ?? [];
  // oxlint-disable-next-line no-console -- scratch instrumentation
  console.warn(
    `[folio-unclassified]${JSON.stringify({ site, message, lastInput, context, stack })}`,
  );
};
