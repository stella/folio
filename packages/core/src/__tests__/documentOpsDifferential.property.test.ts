/** Independent PM transactions and portable model ops must edit the same text and marks. */
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";

import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  INHERIT_RUN_PROPS,
  normalizeForOps,
  OP_STORIES,
} from "../../../docx-core/src/ops/documentOps";
import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { createHarnessState, textblocks } from "./editorHarness";

setDefaultTimeout(propertyTestTimeout(60_000));

const KINDS = ["insert", "delete", "bold", "italic"] as const;
const stepArbitrary = fc.record({
  kind: fc.constantFrom(...KINDS),
  block: fc.nat(20),
  start: fc.nat(100),
  width: fc.nat(100),
  text: fc.constantFrom("x", "café", "東京", "😀", "e\u0301"),
});
const flowArbitrary = fc.record({
  source: fc.constantFrom(
    "Alpha beta.\n\nSecond paragraph.",
    "**Bold** and *italic* text.\n\nCafé 東京 😀.",
    "# Heading\n\n- First item\n- Second item",
  ),
  steps: fc.array(stepArbitrary, { minLength: 1, maxLength: 30 }),
});

/** Ignore allocator/capture metadata; retain every text unit, mark, and block container. */
const contentView = (node: PMNode): unknown => {
  const children: unknown[] = [];
  node.forEach((child) => children.push(contentView(child)));
  return {
    type: node.type.name,
    text: node.text,
    marks: node.marks.map((mark) => mark.toJSON()),
    children,
  };
};

test("PM and Document ops agree after every generated text and formatting step", () => {
  fc.assert(
    fc.property(flowArbitrary, ({ source, steps }) => {
      const original = fromMarkdown(source);
      let state = createHarnessState(original, "editing");
      let model = normalizeForOps(fromProseDoc(state.doc, original));
      for (const step of steps) {
        const blocks = textblocks(state.doc);
        const target = blocks.at(step.block % blocks.length);
        if (!target) throw new Error("generated document has no textblock");
        const blockId: unknown = target.node.attrs["paraId"];
        if (typeof blockId !== "string") throw new Error("generated textblock has no paraId");
        const offset = step.start % (target.node.content.size + 1);
        const end = Math.min(target.node.content.size, offset + step.width);
        const at = { story: OP_STORIES.MAIN, blockId, offset };
        const to = { story: OP_STORIES.MAIN, blockId, offset: end };
        const fromPosition = target.pos + 1 + offset;
        const toPosition = target.pos + 1 + end;
        let transaction = state.tr;
        const operation = (() => {
          switch (step.kind) {
            case "insert":
              transaction = transaction.insertText(step.text, fromPosition, fromPosition);
              return {
                type: DOCUMENT_OP_TYPES.INSERT_TEXT,
                at,
                text: step.text,
                runProps: INHERIT_RUN_PROPS,
              } as const;
            case "delete":
              transaction = transaction.delete(fromPosition, toPosition);
              return { type: DOCUMENT_OP_TYPES.DELETE_RANGE, from: at, to } as const;
            case "bold":
            case "italic": {
              const mark = state.schema.marks[step.kind];
              if (!mark) throw new Error(`missing ${step.kind} mark`);
              transaction = transaction.addMark(fromPosition, toPosition, mark.create());
              const patch = step.kind === "bold" ? { bold: true } : { italic: true };
              return { type: DOCUMENT_OP_TYPES.SET_RUN_PROPS, from: at, to, patch } as const;
            }
            default: {
              const unreachable: never = step.kind;
              return unreachable;
            }
          }
        })();
        // Empty range edits are PM no-ops; they need no model operation.
        if (step.kind !== "insert" && offset === end) continue;
        const result = applyDocumentOp(model, operation);
        if (result.isErr()) throw result.error;
        model = result.value.document;
        state = state.apply(transaction);
        expect(contentView(toProseDoc(model))).toEqual(contentView(state.doc));
      }
    }),
    propertyConfig({ numRuns: 100 }),
  );
});
