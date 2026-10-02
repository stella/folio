/** Independent PM transactions and portable model ops must edit the same text and marks. */
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";

import { extractRunFormatting } from "../layout-bridge/convert/runMarkFormatting";
import { mergeRunFormatting } from "../layout-bridge/convert/runFormattingMerge";
import { paragraphRunDefaults } from "../layout-bridge/convert/textFormattingConversion";
import { expectParagraphAttrs } from "../prosemirror/attrs";

import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  INHERIT_RUN_PROPS,
  normalizeForOps,
  OP_STORIES,
} from "@stll/docx-core/ops";
import {
  assertPinnedProperty,
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { OOXML_NAMESPACES } from "../docx/serializer/partNamespaces";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { expectRunIdentityMarkAttrs } from "../prosemirror/attrs";
import { RUN_IDENTITY_MARK_NAME, runIdentityAttrs } from "../prosemirror/runIdentity";
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

/**
 * RunIdentityExtension strips authored-run identity from typed text. The portable
 * model inserts into its existing run instead, so identity and empty w:rPr
 * provenance may split PM text leaves without changing rendered content.
 * Preserve authored attribute/XML remainders; only the allocator id and empty
 * property-set capture are excluded from this rendered-content comparison.
 */
const renderedMarks = (node: PMNode): unknown[] =>
  node.marks.flatMap((mark) => {
    // Font marks can be omitted when text uses its paragraph defaults.
    // Compare their effective values separately, as the painter does.
    if (
      mark.type.name === "fontFamily" ||
      mark.type.name === "fontSize" ||
      mark.type.name === "runFormattingOverride"
    )
      return [];
    if (mark.type.name !== RUN_IDENTITY_MARK_NAME) return [mark.toJSON()];
    const { preservedAttributes, preserved } = expectRunIdentityMarkAttrs(mark);
    if (!preservedAttributes?.length && !preserved?.children?.length) return [];
    return [{ type: RUN_IDENTITY_MARK_NAME, attrs: { preservedAttributes, preserved } }];
  });

/** Compare every character and mark independently of internal text-leaf partitioning. */
const contentView = (node: PMNode): unknown => {
  const children: unknown[] = [];
  const defaults = node.isTextblock ? paragraphRunDefaults(expectParagraphAttrs(node)) : {};
  node.forEach((child) => {
    if (!child.isText) {
      children.push(contentView(child));
      return;
    }
    const marks = renderedMarks(child);
    const formatting = mergeRunFormatting(defaults, extractRunFormatting(child.marks));
    for (const character of child.text ?? "") {
      children.push({ type: child.type.name, text: character, marks, formatting });
    }
  });
  return { type: node.type.name, text: node.text, marks: renderedMarks(node), children };
};

test("rendered differential oracle detects text, formatting and preserved metadata changes", () => {
  const state = createHarnessState(fromMarkdown("Alpha beta."), "editing");
  const target = textblocks(state.doc).at(0);
  if (!target) throw new Error("oracle fixture has no textblock");
  const from = target.pos + 1;
  const to = from + 1;
  const before = contentView(state.doc);
  const identity = state.schema.marks[RUN_IDENTITY_MARK_NAME];
  if (!identity) throw new Error("oracle fixture has no run identity mark");
  const captureOnly = identity.create(runIdentityAttrs(42, { emptyFormatting: true }));
  expect(contentView(state.tr.addMark(from, to, captureOnly).doc)).toEqual(before);
  expect(contentView(state.tr.insertText("x", from).doc)).not.toEqual(before);
  expect(contentView(state.tr.delete(from, to).doc)).not.toEqual(before);
  const fontSize = state.schema.marks["fontSize"];
  const fontFamily = state.schema.marks["fontFamily"];
  if (!fontSize || !fontFamily) throw new Error("oracle fixture has no font marks");
  expect(contentView(state.tr.addMark(from, to, fontSize.create({ size: 40 })).doc)).not.toEqual(
    before,
  );
  expect(
    contentView(
      state.tr.addMark(from, to, fontFamily.create({ ascii: "Courier New", hAnsi: "Courier New" }))
        .doc,
    ),
  ).not.toEqual(before);
  const inheritedFonts = state.tr
    .removeMark(from, to, fontSize)
    .removeMark(from, to, fontFamily).doc;
  expect(contentView(inheritedFonts)).toEqual(before);
  for (const kind of ["bold", "italic"] as const) {
    const mark = state.schema.marks[kind];
    if (!mark) throw new Error(`oracle fixture has no ${kind} mark`);
    expect(contentView(state.tr.addMark(from, to, mark.create()).doc)).not.toEqual(before);
  }
  const attributeRemainder = identity.create(
    runIdentityAttrs(42, {
      preservedAttributes: [
        { namespace: OOXML_NAMESPACES.w.uri, name: "rsidR", value: "00112233" },
      ],
    }),
  );
  expect(contentView(state.tr.addMark(from, to, attributeRemainder).doc)).not.toEqual(before);
  const xmlRemainder = identity.create(
    runIdentityAttrs(42, {
      preserved: { children: [{ index: 0, xml: '<custom xmlns="urn:folio:test"/>' }] },
    }),
  );
  expect(contentView(state.tr.addMark(from, to, xmlRemainder).doc)).not.toEqual(before);
});

test("PM and Document ops agree after every generated text and formatting step", () => {
  const property = fc.property(flowArbitrary, ({ source, steps }) => {
    const original = fromMarkdown(source);
    let state = createHarnessState(original, "editing");
    let model = normalizeForOps(fromProseDoc(state.doc, original));
    expect(contentView(toProseDoc(model))).toEqual(contentView(state.doc));
    for (const step of steps) {
      const blocks = textblocks(state.doc);
      const target = blocks.at(step.block % blocks.length);
      if (!target) throw new Error("generated document has no textblock");
      const blockId: unknown = target.node.attrs["paraId"];
      if (typeof blockId !== "string") throw new Error("generated textblock has no paraId");
      // The portable contract rejects offsets between UTF-16 surrogate halves.
      // Draw both endpoints in the same valid coordinate space for both paths.
      const boundaries = [0];
      for (const character of target.node.textContent) {
        boundaries.push((boundaries.at(-1) ?? 0) + character.length);
      }
      const startIndex = step.start % boundaries.length;
      const offset = boundaries.at(startIndex) ?? 0;
      const end = boundaries.at(Math.min(boundaries.length - 1, startIndex + step.width)) ?? offset;
      const at = { story: OP_STORIES.MAIN, blockId, offset };
      const to = { story: OP_STORIES.MAIN, blockId, offset: end };
      const fromPosition = target.pos + 1 + offset;
      const toPosition = target.pos + 1 + end;
      // PM insertText inherits stored/cursor marks, whereas Document inherits
      // at the addressed position. Move the independent simulator's cursor
      // there so a previous edit in another paragraph cannot supply its font.
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, fromPosition)));
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
  });
  if (process.env["FOLIO_DOCUMENT_OPS_RANDOM"] === "1") {
    const event = process.env["GITHUB_EVENT_NAME"];
    if (event === "pull_request" || event === "merge_group") {
      throw new Error("Random differential flows run only in the nightly lane");
    }
    assertProperty(property, { numRuns: 100 });
    return;
  }
  assertPinnedProperty(property);
});
