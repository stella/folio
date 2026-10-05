/** Initial wire comparison; full own-field/model differential coverage follows in S2. */
import { test, expect } from "bun:test";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import fc from "fast-check";
import {
  documentArbitrary,
  opSeedArbitrary,
  opFor,
  GENERATED_OP_KINDS,
} from "../../../packages/docx-core/src/ops/__tests__/documentArbitraries";
import { applyDocumentOps } from "../../../packages/docx-core/src/ops/apply";
import { DOCUMENT_OP_TYPES } from "../../../packages/docx-core/src/ops/types";
import { normalizeForOps } from "../../../packages/docx-core/src/ops/contract";
import type { Document, Paragraph } from "../../../packages/docx-core/src/model/document";
import type { DocumentOp } from "../../../packages/docx-core/src/ops/types";
import {
  allocateEditorIntentIds,
  compileEditorIntent,
} from "../../../packages/docx-core/src/ops/editorIntent";
import { captureDocumentOp } from "../../../packages/docx-core/src/ops/wire";
import { decodeTagged, encodeTagged, withoutCaptureSymbols, type Encoded } from "./harness-codec";

const binary =
  process.env["RUST_SPIKE_NATIVE_BINARY"] ??
  fileURLToPath(new URL("../target/debug/canonical-spike", import.meta.url));
const paragraphPatchKind = GENERATED_OP_KINDS.indexOf(DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS);
if (paragraphPatchKind < 0) throw new TypeError("Existing generator has no paragraph patch kind.");

const native = (request: unknown): unknown => {
  const result = spawnSync(binary, [], { input: `${JSON.stringify(request)}\n`, encoding: "utf8" });
  if (result.error) throw result.error;
  expect(result.status, result.stderr).toBe(0);
  return JSON.parse(result.stdout);
};

const wireValue = (value: unknown): unknown => JSON.parse(JSON.stringify(value));

for (const seed of [20261005, 20261006, 20261007]) {
  test(`paragraph patch wire result matches existing generator seed ${seed}`, () => {
    const cases = fc.sample(fc.tuple(documentArbitrary, opSeedArbitrary), { seed, numRuns: 40 });
    for (const [document, generated] of cases) {
      const op = opFor(document, { ...generated, kind: paragraphPatchKind });
      const expected = applyDocumentOps(document, [op]);
      const actual = native({ document, ops: [op] });
      if (expected.isErr()) {
        expect(actual).toEqual({
          status: "refused",
          opType: expected.error.opType,
          reason: expected.error.reason,
          message: expected.error.message,
        });
        continue;
      }
      expect(actual).toEqual(wireValue(expected.value));
      const restored = native({ document: expected.value.document, ops: expected.value.inverse });
      const tsRestored = applyDocumentOps(expected.value.document, expected.value.inverse).unwrap();
      expect(restored).toEqual(wireValue(tsRestored));
    }
  });
}

test("an atomic refused batch matches the TS result", () => {
  const document = normalizeForOps({
    package: { document: { content: [{ type: "paragraph", paraId: "00000001", content: [] }] } },
  });
  const ops = [
    {
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
      story: "main",
      blockId: "00000001",
      patch: { alignment: "center" },
    },
    { type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS, story: "main", blockId: "00000002", patch: {} },
  ] as const;
  const expected = applyDocumentOps(document, ops);
  if (expected.isOk()) throw new TypeError("Missing block must be refused.");
  expect(native({ document, ops })).toEqual({
    status: "refused",
    opType: expected.error.opType,
    reason: expected.error.reason,
    message: expected.error.message,
  });
});

const assertNativeSequence = (document: Document, ops: readonly DocumentOp[]) => {
  const expected = applyDocumentOps(document, ops).unwrap();
  expect(native({ document, ops })).toEqual(wireValue(expected));
  const undo = applyDocumentOps(expected.document, expected.inverse).unwrap();
  expect(native({ document: expected.document, ops: expected.inverse })).toEqual(wireValue(undo));
  const redo = applyDocumentOps(undo.document, undo.inverse).unwrap();
  expect(native({ document: undo.document, ops: undo.inverse })).toEqual(wireValue(redo));
};

/** Object key order is immaterial; Map entry order and array holes are facts. */
const orderedTagged = (value: Encoded): Encoded => {
  if (value === null || typeof value !== "object") return value;
  switch (value.tag) {
    case "object":
    case "nullObject":
      return {
        tag: value.tag,
        entries: value.entries
          .map(([key, entry]) => [key, orderedTagged(entry)] satisfies [string, Encoded])
          .sort(([left], [right]) => left.localeCompare(right)),
      };
    case "map":
      return {
        tag: "map",
        entries: value.entries.map(
          ([key, entry]) => [key, orderedTagged(entry)] satisfies [string, Encoded],
        ),
      };
    case "array":
      return { tag: "array", items: value.items.map(orderedTagged) };
    case "undefined":
    case "hole":
    case "date":
    case "uint8Array":
    case "arrayBuffer":
      return value;
    default: {
      const unreachable: never = value;
      return unreachable;
    }
  }
};

const taggedSnapshot = (value: unknown) =>
  orderedTagged(encodeTagged(withoutCaptureSymbols(value)));

const assertHarnessPhase = (document: unknown, ops: readonly DocumentOp[], expected: unknown) => {
  const reply = native({
    harness: { document: encodeTagged(document), ops: ops.map(captureDocumentOp) },
  });
  if (typeof reply !== "object" || reply === null || !("harness" in reply))
    throw new TypeError(`Tagged operation did not succeed: ${JSON.stringify(reply)}`);
  const actual = decodeTagged(reply.harness);
  expect(taggedSnapshot(actual)).toEqual(taggedSnapshot(expected));
  if (typeof actual !== "object" || actual === null || !("document" in actual))
    throw new TypeError("A tagged Applied result must contain its document.");
  return actual.document;
};

const assertHarnessSequence = (document: Document, ops: readonly DocumentOp[]) => {
  const expected = applyDocumentOps(document, ops).unwrap();
  const nativeDocument = assertHarnessPhase(document, ops, expected);
  const undo = applyDocumentOps(expected.document, expected.inverse).unwrap();
  // Replay the native document, preserving the output sidecars across requests.
  const nativeRestored = assertHarnessPhase(nativeDocument, expected.inverse, undo);
  expect(taggedSnapshot(nativeRestored)).toEqual(taggedSnapshot(document));
  const redo = applyDocumentOps(undo.document, undo.inverse).unwrap();
  const nativeRedone = assertHarnessPhase(nativeRestored, undo.inverse, redo);
  expect(taggedSnapshot(nativeRedone)).toEqual(taggedSnapshot(expected.document));
};

test("tagged operation replay preserves inert package Date, binary, Map order and owned undefined", () => {
  const bytes = new Uint8Array([0, 127, 255, 1]).buffer;
  const document = {
    originalBuffer: bytes,
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "2F48B967",
            content: [{ type: "run", content: [{ type: "text", text: "ž😀abc" }] }],
          },
        ],
      },
      styles: undefined,
      numbering: undefined,
      relationships: new Map([
        [
          "rId9",
          {
            id: "rId9",
            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            target: "media/retained.png",
          },
        ],
        [
          "rId2",
          {
            id: "rId2",
            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
            target: "theme/theme1.xml",
          },
        ],
      ]),
      media: new Map([
        [
          "word/media/retained.png",
          {
            path: "word/media/retained.png",
            mimeType: "image/png",
            data: bytes,
          },
        ],
      ]),
      properties: {
        created: new Date("2026-10-05T09:00:00.123Z"),
        modified: new Date("2026-10-05T09:01:00.456Z"),
      },
    },
  } satisfies Document;
  const at = (offset: number) => ({ story: "main", blockId: "2F48B967", offset }) as const;
  const insertion = {
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: at(3),
    text: "ř😀",
    runProps: "inherit",
  } as const satisfies DocumentOp;
  const patch = {
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story: "main",
    blockId: "2F48B967",
    patch: { alignment: "center" },
  } as const satisfies DocumentOp;
  const deletion = {
    type: DOCUMENT_OP_TYPES.DELETE_RANGE,
    from: at(1),
    to: at(4),
  } as const satisfies DocumentOp;
  for (const ops of [[insertion], [deletion], [patch], [insertion, patch]])
    assertHarnessSequence(document, ops);
});

test("tagged targeted content sidecars refuse relocation instead of claiming equivalence", () => {
  const document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "2F48B967",
            content: [
              {
                type: "run",
                formatting: { bold: undefined },
                content: [{ type: "text", text: "ž😀abc" }],
              },
            ],
          },
        ],
      },
    },
  } satisfies Document;
  const op = {
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: { story: "main", blockId: "2F48B967", offset: 3 },
    text: "x",
    runProps: "inherit",
  } as const satisfies DocumentOp;
  expect(applyDocumentOps(document, [op]).isOk()).toBe(true);
  expect(
    native({ harness: { document: encodeTagged(document), ops: [captureDocumentOp(op)] } }),
  ).toEqual({ status: "unsupported", opType: "insertText", dimension: "harnessSidecarMovement" });
});

test("tracked scalar insertions and deletions match TS provenance and both replay directions", () => {
  const document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "2F48B967",
            content: [
              {
                type: "run",
                formatting: { bold: true },
                content: [
                  { type: "text", text: "ž😀a" },
                  { type: "text", text: "bc" },
                ],
              },
              {
                type: "run",
                formatting: { italic: true },
                content: [{ type: "text", text: "DE" }],
              },
            ],
          },
        ],
      },
    },
  } satisfies Document;
  const at = (offset: number) => ({ story: "main", blockId: "2F48B967", offset }) as const;
  const boundaries = [0, 1, 3, 4, 5, 6, 7, 8];
  const revision = { id: 100, author: "A", date: "2026-01-02T03:04:05Z", initials: "AA" };
  for (const offset of boundaries)
    for (const runProps of ["inherit", {}, { italic: true }] as const) {
      assertNativeSequence(document, [
        { type: DOCUMENT_OP_TYPES.INSERT_TEXT, at: at(offset), text: "é😀", runProps, revision },
      ]);
    }
  for (const from of boundaries)
    for (const to of boundaries.filter((end) => end > from)) {
      assertNativeSequence(document, [
        { type: DOCUMENT_OP_TYPES.DELETE_RANGE, from: at(from), to: at(to), revision },
      ]);
    }
});

const replacementContents = [
  [{ type: "run", content: [{ type: "text", text: "a😀bc" }] }],
  [
    {
      type: "run",
      preservedAttributes: [{ name: "rsidR", value: "1234ABCD" }],
      content: [
        { type: "text", text: "a😀" },
        { type: "text", text: "b" },
      ],
    },
    { type: "run", formatting: { bold: true }, content: [{ type: "text", text: "c" }] },
  ],
] satisfies Paragraph["content"][];

for (const [fixture, content] of replacementContents.entries())
  test(`replaceText compiler scalar spans and atomic inverse closure, fixture ${fixture}`, () => {
    const document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "00000001",
              content,
            },
          ],
        },
      },
    } satisfies Document;
    const position = (offset: number) => ({ story: "main", blockId: "00000001", offset }) as const;
    const boundaries = [0];
    let offset = 0;
    for (const run of content)
      for (const leaf of run.content)
        for (const scalar of leaf.text) {
          offset += scalar.length;
          boundaries.push(offset);
        }
    for (const from of boundaries)
      for (const to of boundaries.filter((end) => end >= from))
        for (const text of ["", "ž😀"]) {
          const intent = {
            type: "replaceText",
            from: position(from),
            to: position(to),
            text,
          } as const;
          const allocation = allocateEditorIntentIds(document, intent);
          const compiled = compileEditorIntent(document, {
            intent,
            mode: { type: "editing", newIds: allocation.newIds },
          }).unwrap();
          assertNativeSequence(document, compiled.ops);
        }
  });

test("plain inline splits and joins compare both seam depths and exact refusals", () => {
  const document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "2F48B967",
            content: [
              {
                type: "run",
                content: [
                  { type: "text", text: "a😀" },
                  { type: "text", text: "bc" },
                ],
              },
            ],
          },
        ],
      },
    },
  } satisfies Document;
  const at = (offset: number) => ({ story: "main", blockId: "2F48B967", offset }) as const;
  const exercised = new Set<number>();
  for (const offset of [1, 3, 4])
    for (const depth of [1, 2]) {
      const op = { type: DOCUMENT_OP_TYPES.SPLIT_INLINE, at: at(offset), depth } as const;
      const expected = applyDocumentOps(document, [op]);
      if (expected.isOk()) {
        exercised.add(depth);
        assertNativeSequence(document, [op]);
      } else {
        expect(native({ document, ops: [op] })).toEqual({
          status: "refused",
          opType: expected.error.opType,
          reason: expected.error.reason,
          message: expected.error.message,
        });
      }
    }
  expect([...exercised].sort()).toEqual([1, 2]);
  for (const type of [DOCUMENT_OP_TYPES.SPLIT_INLINE, DOCUMENT_OP_TYPES.JOIN_INLINE])
    for (const offset of [0, 1, 2, 3, 5])
      for (const depth of [0, 1, 2, 3]) {
        const op = { type, at: at(offset), depth } as const;
        const expected = applyDocumentOps(document, [op]);
        if (expected.isOk()) {
          assertNativeSequence(document, [op]);
          continue;
        }
        expect(native({ document, ops: [op] })).toEqual({
          status: "refused",
          opType: expected.error.opType,
          reason: expected.error.reason,
          message: expected.error.message,
        });
      }
});

test("copied wrapper identities never silently duplicate package identities", () => {
  const wrapper = {
    type: "insertion",
    info: { id: 7, author: "A", date: "2026-01-02T03:04:05Z" },
    content: [{ type: "run", content: [{ type: "text", text: "x" }] }],
  } as const;
  const document = {
    package: {
      document: {
        content: [
          { type: "paragraph", paraId: "00000001", content: [] },
          { type: "paragraph", paraId: "00000002", content: [wrapper] },
        ],
      },
    },
  };
  for (const content of [[wrapper], [wrapper, wrapper]]) {
    expect(
      native({
        document,
        ops: [
          {
            type: "insertContent",
            at: { story: "main", blockId: "00000001", offset: 0 },
            slice: { content, openStart: 0, openEnd: 0 },
          },
        ],
      }),
    ).toEqual({
      status: "unsupported",
      opType: "insertContent",
      dimension: "copiedInlineIdentityFreshening",
    });
  }
});
