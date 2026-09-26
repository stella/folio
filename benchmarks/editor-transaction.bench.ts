/**
 * Editor transaction benchmark: one many-step transaction through folio's
 * plugin stack.
 *
 * Paste, replace-all, accept-all, an AI edit batch and a document compare all
 * dispatch one transaction with a step per change. Every plugin that carries
 * positions through that transaction (the paragraph change tracker, the
 * paragraph-id allocator, base-direction detection, run identity) runs on
 * dispatch, so their cost per step is what this measures. The document and
 * the transaction are built once during setup; only applying it is measured.
 */
import { withCodSpeed } from "@codspeed/tinybench-plugin";
import { schema, singletonManager, type DocxNode } from "@stll/folio-core/prosemirror/schema";

import { EditorState, type Transaction } from "prosemirror-state";
import { Bench } from "tinybench";

import { MICRO_BENCH_OPTIONS } from "./config";

const PARAGRAPHS = 2_000;

const largeDocument = (): DocxNode =>
  schema.node(
    "doc",
    null,
    Array.from({ length: PARAGRAPHS }, (_, index) =>
      schema.node(
        "paragraph",
        { paraId: (index + 1).toString(16).padStart(8, "0").toUpperCase() },
        [schema.text(`Clause ${String(index + 1)}. The supplier shall deliver the goods on time.`)],
      ),
    ),
  );

const editedParagraphStarts = (state: EditorState): number[] => {
  const starts: number[] = [];
  state.doc.forEach((_node, offset, index) => {
    if (index % 2 === 0) {
      starts.push(offset);
    }
  });
  return starts;
};

/** One word replaced in every other paragraph, back to front: a replace-all. */
const replaceAll = (state: EditorState): Transaction => {
  const tr = state.tr;
  for (const start of editedParagraphStarts(state).toReversed()) {
    tr.insertText("Section", start + 1, start + 1 + "Clause".length);
  }
  return tr;
};

/** A mark over one span in every other paragraph: formatting many ranges at once. */
const formatAll = (state: EditorState): Transaction => {
  const tr = state.tr;
  const bold = schema.marks["bold"];
  if (!bold) {
    throw new Error("bold mark missing from schema");
  }
  for (const start of editedParagraphStarts(state)) {
    tr.addMark(start + 5, start + 20, bold.create());
  }
  return tr;
};

export function editorTransactionBench(): Bench {
  const bench = withCodSpeed(new Bench(MICRO_BENCH_OPTIONS));
  const state = EditorState.create({
    schema,
    doc: largeDocument(),
    plugins: singletonManager.getPlugins(),
  });

  for (const [label, build] of [
    ["replace-all", replaceAll],
    ["format-all", formatAll],
  ] as const) {
    const tr = build(state);
    bench.add(
      `${label} · ${String(tr.steps.length)} steps · ${String(PARAGRAPHS)} paragraphs`,
      () => {
        state.applyTransaction(tr);
      },
    );
  }

  return bench;
}
