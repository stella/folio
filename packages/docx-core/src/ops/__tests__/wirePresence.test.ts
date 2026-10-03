import { expect, setDefaultTimeout, test } from "bun:test";
import { deepStrictEqual, notDeepStrictEqual } from "node:assert/strict";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { Document, Paragraph, ParagraphMarkChange } from "../../model/document";
import { applyDocumentOp, applyDocumentOpEnvelope, applyDocumentOps } from "../apply";
import { allocateEditorIntentIds, compileEditorIntent, type EditorIntent } from "../editorIntent";
import { paragraphLogicalText } from "../offsets";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  DOCUMENT_OP_SCHEMA_VERSION,
  OP_STORIES,
  toOpEnvelope,
  type DocumentOp,
  type ParagraphPropsPatch,
  type RunPropsPatch,
} from "../types";
import { captureDocumentOp, restoreDocumentOp } from "../wire";

setDefaultTimeout(propertyTestTimeout(30_000));

const jsonTransport = <T>(value: T): T => JSON.parse(JSON.stringify(value));
const ownUndefined = (record: object, key: string) =>
  Object.defineProperty(record, key, {
    value: undefined,
    enumerable: true,
    configurable: true,
    writable: true,
  });

const sourceParagraph = (): Paragraph => {
  const paragraph = {
    type: "paragraph",
    paraId: "00000001",
    formatting: { runProperties: {} },
    propertyChanges: [
      {
        type: "paragraphPropertyChange",
        info: { id: 10, author: "Source" },
        previousFormatting: {},
        currentFormatting: {},
      },
    ],
    content: [
      {
        type: "hyperlink",
        href: "https://example.com/source",
        children: [
          {
            type: "run",
            formatting: {},
            propertyChanges: [
              {
                type: "runPropertyChange",
                info: { id: 11, author: "Source" },
                previousFormatting: {},
                currentFormatting: {},
              },
            ],
            content: [{ type: "text", text: "abc" }],
          },
        ],
      },
    ],
  } satisfies Paragraph;
  ownUndefined(paragraph.formatting, "numPr");
  ownUndefined(paragraph.formatting.runProperties, "bold");
  const paragraphReview = paragraph.propertyChanges.at(0);
  const run = paragraph.content.at(0)?.children.at(0);
  const runReview = run?.propertyChanges.at(0);
  if (!paragraphReview || !run || !runReview)
    throw new TypeError("Nested source fixture disappeared.");
  ownUndefined(paragraphReview.previousFormatting, "alignment");
  ownUndefined(paragraphReview.currentFormatting, "keepNext");
  ownUndefined(run.formatting, "italic");
  ownUndefined(runReview.previousFormatting, "bold");
  ownUndefined(runReview.currentFormatting, "color");
  return paragraph;
};

const replacement = (): DocumentOp => {
  const original = sourceParagraph();
  const changed = structuredClone(original);
  changed.content.push({ type: "run", content: [{ type: "text", text: "replacement" }] });
  const op = {
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: OP_STORIES.MAIN,
    expected: [original],
    blocks: [changed],
  } satisfies DocumentOp;
  ownUndefined(op, "sectionView");
  return op;
};

test("non-object wire operations refuse atomically for current and unsupported schemas", () => {
  const document = { package: { document: { content: [sourceParagraph()] } } } satisfies Document;
  const before = structuredClone(document);
  for (const op of [null, undefined, true, 1, "insertText", []]) {
    // JSON decoding is deliberately unchecked here to exercise the runtime boundary.
    const decoded = JSON.parse(JSON.stringify({ op }));
    const restored = restoreDocumentOp(decoded.op);
    expect(restored.isErr()).toBe(true);
    if (restored.isOk()) throw new TypeError("Non-object operation unexpectedly restored.");
    expect(restored.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH);
    expect(restored.error.opType).toBeUndefined();
    for (const schema of [DOCUMENT_OP_SCHEMA_VERSION, DOCUMENT_OP_SCHEMA_VERSION + 1]) {
      const applied = applyDocumentOpEnvelope(document, { schema, op: decoded.op });
      expect(applied.isErr()).toBe(true);
      if (applied.isOk()) throw new TypeError("Non-object operation unexpectedly applied.");
      expect(applied.error.reason).toBe(
        schema === DOCUMENT_OP_SCHEMA_VERSION
          ? DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH
          : DOCUMENT_OP_REFUSAL_REASONS.UNSUPPORTED_SCHEMA,
      );
      expect(applied.error.opType).toBeUndefined();
      deepStrictEqual(document, before);
    }
  }
});

test("capture and envelope transport retain recursively owned undefined fields that raw JSON loses", () => {
  const op = replacement();
  const snapshot = structuredClone(op);
  notDeepStrictEqual(jsonTransport(op), op);
  const captured = captureDocumentOp(op);
  expect(captured.undefinedFields?.length).toBeGreaterThanOrEqual(15);
  deepStrictEqual(restoreDocumentOp(jsonTransport(captured)).unwrap(), op);
  deepStrictEqual(restoreDocumentOp(jsonTransport(toOpEnvelope(op)).op).unwrap(), op);
  deepStrictEqual(captureDocumentOp(captured), captured);
  deepStrictEqual(op, snapshot);
  const restored = restoreDocumentOp(jsonTransport(captured)).unwrap();
  notDeepStrictEqual(restored, jsonTransport(op));
});

test("centrally captured inverses restore exact authored snapshots after JSON without recapture", () => {
  const op = replacement();
  if (op.type !== DOCUMENT_OP_TYPES.REPLACE_BLOCKS)
    throw new TypeError("Expected replacement fixture.");
  const document = {
    package: { document: { content: structuredClone(op.expected) } },
  } satisfies Document;
  const before = structuredClone(document);
  const applied = applyDocumentOpEnvelope(document, jsonTransport(toOpEnvelope(op))).unwrap();
  expect(applied.inverse.some((inverse) => (inverse.undefinedFields?.length ?? 0) > 0)).toBe(true);
  const undo = applyDocumentOps(applied.document, jsonTransport(applied.inverse)).unwrap();
  deepStrictEqual(undo.document, before);
  const redo = applyDocumentOps(undo.document, jsonTransport(undo.inverse)).unwrap();
  deepStrictEqual(redo.document, applied.document);
  deepStrictEqual(document, before);
});

const invalidPaths = [
  { name: "empty", paths: [[]] },
  ...["__proto__", "constructor", "prototype"].map((segment) => ({
    name: segment,
    paths: [["expected", "0", segment]],
  })),
  { name: "type identity", paths: [["type"]] },
  { name: "presence identity", paths: [["undefinedFields"]] },
  {
    name: "duplicate",
    paths: [
      ["expected", "0", "absent"],
      ["expected", "0", "absent"],
    ],
  },
  {
    name: "ancestor first",
    paths: [
      ["expected", "0", "absent"],
      ["expected", "0", "absent", "child"],
    ],
  },
  {
    name: "descendant first",
    paths: [
      ["expected", "0", "absent", "child"],
      ["expected", "0", "absent"],
    ],
  },
  { name: "absent parent", paths: [["expected", "0", "missing", "child"]] },
  { name: "primitive parent", paths: [["story", "child"]] },
  { name: "defined leaf", paths: [["expected", "0", "paraId"]] },
  ...["1", "01", "-1", "1.0", "length"].map((slot) => ({
    name: `array slot ${slot}`,
    paths: [["expected", slot]],
  })),
];

test.each(invalidPaths)("malformed presence metadata refuses atomically: $name", ({ paths }) => {
  const original = replacement();
  if (original.type !== DOCUMENT_OP_TYPES.REPLACE_BLOCKS)
    throw new TypeError("Expected replacement fixture.");
  const document = {
    package: { document: { content: structuredClone(original.expected) } },
  } satisfies Document;
  const op = { ...jsonTransport(original), undefinedFields: paths };
  const before = structuredClone(document);
  const opBefore = structuredClone(op);
  const restored = restoreDocumentOp(op);
  expect(restored.isErr()).toBe(true);
  if (restored.isOk()) throw new TypeError("Malformed presence metadata unexpectedly restored.");
  expect(restored.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH);
  const applied = applyDocumentOp(document, op);
  expect(applied.isErr()).toBe(true);
  if (applied.isOk()) throw new TypeError("Malformed presence metadata unexpectedly applied.");
  expect(applied.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH);
  deepStrictEqual(document, before);
  deepStrictEqual(op, opBefore);
});

test("presence metadata cannot traverse an inherited intermediate object", () => {
  const op = replacement();
  Object.setPrototypeOf(op, { inherited: { leaf: undefined } });
  op.undefinedFields = [["inherited", "leaf"]];
  const restored = restoreDocumentOp(op);
  expect(restored.isErr()).toBe(true);
  if (restored.isOk()) throw new TypeError("Inherited metadata parent unexpectedly accepted.");
  expect(restored.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH);
  expect(Object.hasOwn(op, "inherited")).toBe(false);
});

test("generated multi-step transported operations keep exact immediate undo and the complete journal", () => {
  assertProperty(
    fc.property(
      fc.array(
        fc.record({
          target: fc.nat(1000),
          gap: fc.nat(1000),
          missing: fc.boolean(),
          kind: fc.constantFrom("insert", "properties"),
          text: fc.constantFrom("x", "yz", "é"),
          alignment: fc.constantFrom("left" as const, "right" as const, "center" as const),
        }),
        { minLength: 8, maxLength: 16 },
      ),
      (steps) => {
        const paragraphs = ["00000001", "00000002"].map((paraId) => {
          const paragraph = sourceParagraph();
          paragraph.paraId = paraId;
          delete paragraph.propertyChanges;
          const formatting = {};
          ownUndefined(formatting, "italic");
          paragraph.content = [
            { type: "run", formatting, content: [{ type: "text", text: "abc" }] },
          ];
          return paragraph;
        });
        const original = { package: { document: { content: paragraphs } } } satisfies Document;
        let document: Document = original;
        const journal: {
          before: Document;
          after: Document;
          inverse: readonly DocumentOp[];
          redo: readonly DocumentOp[];
        }[] = [];
        const refusals = new Map<string, number>();
        for (const step of steps) {
          const paragraph = document.package.document.content.at(step.target % 2);
          if (paragraph?.type !== "paragraph" || !paragraph.paraId)
            throw new TypeError("Generated target disappeared.");
          const blockId = step.missing ? "FFFFFFFF" : paragraph.paraId;
          const textBefore = paragraphLogicalText(paragraph);
          const offset = step.gap % (textBefore.length + 1);
          const op: DocumentOp =
            step.kind === "insert"
              ? {
                  type: DOCUMENT_OP_TYPES.INSERT_TEXT,
                  at: { story: OP_STORIES.MAIN, blockId, offset },
                  text: step.text,
                  runProps: "inherit",
                }
              : {
                  type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
                  story: OP_STORIES.MAIN,
                  blockId,
                  patch: { alignment: step.alignment },
                };
          const before = structuredClone(document);
          const applied = applyDocumentOpEnvelope(document, jsonTransport(toOpEnvelope(op)));
          if (applied.isErr()) {
            refusals.set(applied.error.reason, (refusals.get(applied.error.reason) ?? 0) + 1);
            expect(step.missing).toBe(true);
            expect(applied.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND);
            deepStrictEqual(document, before);
            continue;
          }
          expect(step.missing).toBe(false);
          const direct = applyDocumentOp(structuredClone(before), op).unwrap();
          deepStrictEqual(applied.value.document, direct.document);
          const changed = applied.value.document.package.document.content.at(step.target % 2);
          if (changed?.type !== "paragraph") throw new TypeError("Changed target disappeared.");
          if (step.kind === "insert")
            expect(paragraphLogicalText(changed)).toBe(
              textBefore.slice(0, offset) + step.text + textBefore.slice(offset),
            );
          const undo = applyDocumentOps(
            applied.value.document,
            jsonTransport(applied.value.inverse),
          ).unwrap();
          deepStrictEqual(undo.document, before);
          const redo = applyDocumentOps(undo.document, jsonTransport(undo.inverse)).unwrap();
          deepStrictEqual(redo.document, applied.value.document);
          journal.push({
            before,
            after: structuredClone(applied.value.document),
            inverse: jsonTransport(applied.value.inverse),
            redo: jsonTransport(undo.inverse),
          });
          document = redo.document;
        }
        expect([...refusals.keys()]).toEqual(
          steps.some((step) => step.missing) ? [DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND] : [],
        );
        expect(refusals.get(DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND) ?? 0).toBe(
          steps.filter((step) => step.missing).length,
        );
        for (const entry of journal.toReversed()) {
          document = applyDocumentOps(document, entry.inverse).unwrap().document;
          deepStrictEqual(document, entry.before);
        }
        deepStrictEqual(document, original);
        for (const entry of journal) {
          document = applyDocumentOps(document, entry.redo).unwrap().document;
          deepStrictEqual(document, entry.after);
        }
      },
    ),
    { numRuns: 100 },
  );
});

const propertyDocument = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          formatting: { keepNext: false },
          content: [
            {
              type: "run",
              formatting: { italic: false },
              content: [{ type: "text", text: "abc" }],
            },
          ],
        },
      ],
    },
  },
});

const propertyRecords = (document: Document) => {
  const paragraph = document.package.document.content.at(0);
  const run = paragraph?.type === "paragraph" ? paragraph.content.at(0) : undefined;
  if (
    paragraph?.type !== "paragraph" ||
    !paragraph.formatting ||
    run?.type !== "run" ||
    !run.formatting
  )
    throw new TypeError("Property presence fixture disappeared.");
  return { paragraph: paragraph.formatting, run: run.formatting };
};

test("generated property patches preserve omitted, owned undefined and concrete values through every inverse", () => {
  assertProperty(
    fc.property(
      fc.constantFrom("omitted", "undefined", "concrete"),
      fc.constantFrom("omitted", "undefined", "concrete"),
      fc.array(
        fc.record({
          field: fc.constantFrom("paragraph", "run"),
          value: fc.constantFrom("omitted", "undefined", "concrete", "remove"),
          bold: fc.boolean(),
        }),
        { minLength: 8, maxLength: 16 },
      ),
      (paragraphState, runState, steps) => {
        const original = propertyDocument();
        const records = propertyRecords(original);
        if (paragraphState === "undefined") ownUndefined(records.paragraph, "numPr");
        if (paragraphState === "concrete") records.paragraph.numPr = { kind: "none" };
        if (runState === "undefined") ownUndefined(records.run, "bold");
        if (runState === "concrete") records.run.bold = true;
        let document = original;
        const journal: {
          before: Document;
          after: Document;
          inverse: readonly DocumentOp[];
          redo: readonly DocumentOp[];
        }[] = [];
        for (const step of steps) {
          const before = structuredClone(document);
          const paragraphPatch: ParagraphPropsPatch = {};
          const runPatch: RunPropsPatch = {};
          if (step.value === "undefined")
            ownUndefined(
              step.field === "paragraph" ? paragraphPatch : runPatch,
              step.field === "paragraph" ? "numPr" : "bold",
            );
          if (step.value === "remove") {
            paragraphPatch.numPr = null;
            runPatch.bold = null;
          }
          if (step.value === "concrete") {
            paragraphPatch.numPr = { kind: "none" };
            runPatch.bold = step.bold;
          }
          const op: DocumentOp =
            step.field === "paragraph"
              ? {
                  type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
                  story: OP_STORIES.MAIN,
                  blockId: "00000001",
                  patch: paragraphPatch,
                }
              : {
                  type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
                  from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
                  to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 3 },
                  patch: runPatch,
                };
          const expected = structuredClone(before);
          const expectedRecords = propertyRecords(expected);
          const target =
            step.field === "paragraph" ? expectedRecords.paragraph : expectedRecords.run;
          const key = step.field === "paragraph" ? "numPr" : "bold";
          if (step.value === "remove") Reflect.deleteProperty(target, key);
          if (step.value === "undefined") ownUndefined(target, key);
          if (step.value === "concrete")
            Reflect.set(target, key, step.field === "paragraph" ? { kind: "none" } : step.bold);
          const applied = applyDocumentOp(document, op).unwrap();
          deepStrictEqual(applied.document, expected);
          const transported = applyDocumentOpEnvelope(
            before,
            jsonTransport(toOpEnvelope(op)),
          ).unwrap();
          deepStrictEqual(transported.document, expected);
          const memoryUndo = applyDocumentOps(applied.document, applied.inverse).unwrap();
          deepStrictEqual(memoryUndo.document, before);
          const undo = applyDocumentOps(applied.document, jsonTransport(applied.inverse)).unwrap();
          deepStrictEqual(undo.document, before);
          const redo = applyDocumentOps(undo.document, jsonTransport(undo.inverse)).unwrap();
          deepStrictEqual(redo.document, expected);
          journal.push({
            before,
            after: expected,
            inverse: jsonTransport(applied.inverse),
            redo: jsonTransport(undo.inverse),
          });
          document = redo.document;
        }
        for (const entry of journal.toReversed()) {
          document = applyDocumentOps(document, entry.inverse).unwrap().document;
          deepStrictEqual(document, entry.before);
        }
        deepStrictEqual(document, original);
        for (const entry of journal) {
          document = applyDocumentOps(document, entry.redo).unwrap().document;
          deepStrictEqual(document, entry.after);
        }
      },
    ),
    { numRuns: 100 },
  );
});

test("stale property preconditions distinguish absent fields from owned undefined in both directions", () => {
  for (const field of ["paragraph", "run"] as const) {
    for (const owned of [false, true]) {
      const document = propertyDocument();
      const records = propertyRecords(document);
      if (owned)
        ownUndefined(
          field === "paragraph" ? records.paragraph : records.run,
          field === "paragraph" ? "numPr" : "bold",
        );
      const paragraphExpected: ParagraphPropsPatch = {};
      const runExpected: RunPropsPatch = {};
      if (owned) {
        paragraphExpected.numPr = null;
        runExpected.bold = null;
      } else {
        ownUndefined(paragraphExpected, "numPr");
        ownUndefined(runExpected, "bold");
      }
      const op: DocumentOp =
        field === "paragraph"
          ? {
              type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
              story: OP_STORIES.MAIN,
              blockId: "00000001",
              patch: { numPr: { kind: "none" } },
              expected: paragraphExpected,
            }
          : {
              type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
              from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
              to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 3 },
              patch: { bold: true },
              expected: runExpected,
            };
      const before = structuredClone(document);
      for (const refused of [
        applyDocumentOp(document, op),
        applyDocumentOpEnvelope(document, jsonTransport(toOpEnvelope(op))),
      ]) {
        expect(refused.isErr()).toBe(true);
        if (refused.isOk())
          throw new TypeError("Wrong property presence precondition unexpectedly applied.");
        expect(refused.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
        deepStrictEqual(document, before);
      }
    }
  }
});

test("generated paragraph mark cut-depth presence survives captured journals and exact histories", () => {
  assertProperty(
    fc.property(fc.integer({ min: 1, max: 2147483647 }), (depth) => {
      const marks = [
        { kind: "ins", info: { id: 20, author: "Source" } },
        { kind: "ins", info: { id: 20, author: "Source" }, resolutionJoin: undefined },
        { kind: "ins", info: { id: 20, author: "Source" }, resolutionJoin: 0 },
        { kind: "ins", info: { id: 20, author: "Source" }, resolutionJoin: depth },
      ] satisfies ParagraphMarkChange[];
      for (const mark of marks) {
        const paragraph = {
          type: "paragraph",
          paraId: "00000001",
          pPrMark: mark,
          content: [{ type: "run", content: [{ type: "text", text: "before" }] }],
        } satisfies Paragraph;
        const document = { package: { document: { content: [paragraph] } } } satisfies Document;
        const before = structuredClone(document);
        const changed = structuredClone(paragraph);
        changed.content = [{ type: "run", content: [{ type: "text", text: "after" }] }];
        const op = {
          type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
          story: OP_STORIES.MAIN,
          expected: [paragraph],
          blocks: [changed],
        } satisfies DocumentOp;
        const restored = restoreDocumentOp(jsonTransport(toOpEnvelope(op)).op).unwrap();
        deepStrictEqual(restored, op);
        if (Object.hasOwn(mark, "resolutionJoin") && mark.resolutionJoin === undefined)
          notDeepStrictEqual(jsonTransport(op), op);
        const applied = applyDocumentOpEnvelope(document, jsonTransport(toOpEnvelope(op))).unwrap();
        deepStrictEqual(applied.document, applyDocumentOp(document, op).unwrap().document);
        const undo = applyDocumentOps(applied.document, applied.inverse).unwrap();
        const transportedUndo = applyDocumentOps(
          applied.document,
          jsonTransport(applied.inverse),
        ).unwrap();
        deepStrictEqual(undo.document, before);
        deepStrictEqual(transportedUndo.document, before);
        const redo = applyDocumentOps(undo.document, jsonTransport(undo.inverse)).unwrap();
        deepStrictEqual(redo.document, applied.document);
        deepStrictEqual(document, before);
      }
    }),
    { numRuns: 100 },
  );
});

test("generated split marks transport source cut depths without losing independent authored bidi siblings", () => {
  assertProperty(
    fc.property(fc.integer({ min: 0, max: 3 }), (offset) => {
      const document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                paraId: "00000001",
                content: [
                  {
                    type: "hyperlink",
                    href: "https://example.com/source",
                    children: [
                      {
                        type: "inlineWrapper",
                        kind: "bidi",
                        control: "embedding",
                        content: [
                          {
                            type: "preservedInline",
                            xml: "<w:proofErr w:type='spellStart'/>",
                            text: "",
                          },
                        ],
                      },
                      {
                        type: "inlineWrapper",
                        kind: "bidi",
                        control: "embedding",
                        content: [{ type: "run", content: [{ type: "text", text: "abc" }] }],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        },
      } satisfies Document;
      const before = structuredClone(document);
      const intent = {
        type: "splitParagraph",
        at: { story: OP_STORIES.MAIN, blockId: "00000001", offset },
        newBlockId: "00000002",
      } as const satisfies EditorIntent;
      const allocation = allocateEditorIntentIds(document, intent);
      const compiled = compileEditorIntent(document, {
        intent,
        mode: {
          type: "suggesting",
          revision: { id: allocation.revisionId, author: "Editor" },
          newIds: allocation.newIds,
        },
      }).unwrap();
      const wireOps = compiled.ops.map((op) => jsonTransport(toOpEnvelope(op)).op);
      for (const [index, op] of wireOps.entries()) {
        const originalOp = compiled.ops.at(index);
        if (originalOp === undefined)
          throw new TypeError("Transport preserves every generated operation.");
        deepStrictEqual(restoreDocumentOp(op).unwrap(), restoreDocumentOp(originalOp).unwrap());
      }
      const memory = applyDocumentOps(document, compiled.ops).unwrap();
      const leading = memory.document.package.document.content.at(0);
      if (leading?.type !== "paragraph" || leading.pPrMark === undefined)
        throw new TypeError("Applying the source split must emit a paragraph mark.");
      expect(Object.hasOwn(leading.pPrMark, "resolutionJoin")).toBe(true);
      expect(Number.isInteger(leading.pPrMark.resolutionJoin)).toBe(true);
      const transported = applyDocumentOps(document, wireOps).unwrap();
      deepStrictEqual(transported.document, memory.document);
      const undo = applyDocumentOps(
        transported.document,
        jsonTransport(transported.inverse),
      ).unwrap();
      deepStrictEqual(undo.document, before);
      const redo = applyDocumentOps(undo.document, jsonTransport(undo.inverse)).unwrap();
      deepStrictEqual(redo.document, transported.document);
      deepStrictEqual(document, before);
    }),
    { numRuns: 25 },
  );
});
