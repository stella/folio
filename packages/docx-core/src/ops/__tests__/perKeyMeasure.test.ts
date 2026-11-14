// Measurement only: compile and apply typed keys in a long document.
import { expect, test } from "bun:test";

import type { Document } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { compileEditorIntent } from "../editorIntent";
import { formatParaId } from "../idLease";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

const PARAGRAPHS = 2_000;
const KEYS = 300;

const longDocument = (): Document => ({
  package: {
    document: {
      content: Array.from({ length: PARAGRAPHS }, (_, index) => ({
        type: "paragraph" as const,
        paraId: formatParaId(0x1000_0000 + index),
        content: [
          {
            type: "run" as const,
            content: [
              {
                type: "text" as const,
                text: `Clause ${String(index + 1)}. The supplier shall deliver the goods on time.`,
              },
            ],
          },
          {
            type: "run" as const,
            formatting: { bold: true },
            content: [{ type: "text" as const, text: " Bold tail." }],
          },
        ],
      })),
    },
  },
});

const percentile = (values: readonly number[], share: number): number => {
  const sorted = values.toSorted((left, right) => left - right);
  return (
    sorted[Math.min(sorted.length - 1, Math.floor(share * sorted.length))] ?? 0
  );
};

test("PERKEY typing in a long document", () => {
  let document = longDocument();
  const blockId = formatParaId(0x1000_0000 + PARAGRAPHS / 2);
  const at = (offset: number) => ({ story: OP_STORIES.MAIN, blockId, offset });
  const timings: number[] = [];
  const editTimings: number[] = [];
  for (let key = 0; key < KEYS; key += 1) {
    const started = performance.now();
    const compiled = compileEditorIntent(document, {
      intent: {
        type: "replaceText",
        from: at(10 + key),
        to: at(10 + key),
        text: "x",
      },
      mode: { type: "editing" },
    });
    if (compiled.isErr()) throw compiled.error;
    const applied = applyDocumentOps(document, compiled.value.ops);
    if (applied.isErr()) throw applied.error;
    document = applied.value.document;
    timings.push(performance.now() - started);
    if (key % 10 === 9) {
      const editStarted = performance.now();
      const formatted = applyDocumentOps(document, [
        {
          type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
          from: at(2),
          to: at(4),
          patch: { italic: key % 20 === 9 ? true : null },
        },
        {
          type: DOCUMENT_OP_TYPES.DELETE_RANGE,
          from: at(10 + key),
          to: at(11 + key),
        },
      ]);
      if (formatted.isErr()) throw formatted.error;
      document = formatted.value.document;
      editTimings.push(performance.now() - editStarted);
    }
  }
  const report = {
    keyMedianMs: percentile(timings.slice(20), 0.5),
    keyP95Ms: percentile(timings.slice(20), 0.95),
    formatEraseMedianMs: percentile(editTimings, 0.5),
  };
  // oxlint-disable-next-line no-console
  console.log(`PERKEY ${JSON.stringify(report)}`);
  expect(timings.length).toBe(KEYS);
});
