/** Schema-10 oracle pinned independently of main's schema-9 operation union. */
import { expect, test } from "bun:test";
import { deepStrictEqual } from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import {
  decodeTagged,
  encodeTagged,
  withoutCaptureSymbols,
  HarnessCodecError,
  type Encoded,
} from "./harness-codec";
import {
  assertCommentOraclePrepared,
  commentOracleModule,
  COMMENT_ORACLE_COMMIT,
} from "./prepare-comment-oracle";
import type {
  Document,
  Comment,
} from "../target/comment-oracle/617c0f4703a6cf43adb7ecda34366c29b93cc60f/packages/docx-core/src/model/document";
import type { DocumentOp } from "../target/comment-oracle/617c0f4703a6cf43adb7ecda34366c29b93cc60f/packages/docx-core/src/ops/types";

assertCommentOraclePrepared();
// These types come from the generated pinned source, not the main schema-9 union.
const oracle: typeof import("../target/comment-oracle/617c0f4703a6cf43adb7ecda34366c29b93cc60f/packages/docx-core/src/ops/apply") =
  await import(commentOracleModule("ops/apply.ts").href);
const wire: typeof import("../target/comment-oracle/617c0f4703a6cf43adb7ecda34366c29b93cc60f/packages/docx-core/src/ops/wire") =
  await import(commentOracleModule("ops/wire.ts").href);
const types: typeof import("../target/comment-oracle/617c0f4703a6cf43adb7ecda34366c29b93cc60f/packages/docx-core/src/ops/types") =
  await import(commentOracleModule("ops/types.ts").href);

const native = (document: unknown, operations: readonly DocumentOp[]): unknown => {
  const process = spawnSync(
    process.env["RUST_SPIKE_NATIVE_BINARY"] ??
      fileURLToPath(new URL("../target/debug/canonical-spike", import.meta.url)),
    [],
    {
      input: `${JSON.stringify({ harness: { document: encodeTagged(document), ops: operations.map((op) => wire.captureDocumentOp(op)) } })}\n`,
      encoding: "utf8",
    },
  );
  if (process.error) throw process.error;
  expect(process.status, process.stderr).toBe(0);
  const response: unknown = JSON.parse(process.stdout);
  return response;
};

/** Ignore plain object field order; keep Map order, holes and all presence tags. */
const orderedTagged = (encoded: Encoded): Encoded => {
  if (encoded === null || typeof encoded !== "object") return encoded;
  switch (encoded.tag) {
    case "object":
      return {
        tag: "object",
        entries: encoded.entries
          .map(([key, value]) => [key, orderedTagged(value)] satisfies [string, Encoded])
          .sort(([left], [right]) => {
            if (left < right) return -1;
            if (left > right) return 1;
            return 0;
          }),
      };
    case "map":
      return {
        tag: "map",
        entries: encoded.entries.map(
          ([key, value]) => [key, orderedTagged(value)] satisfies [string, Encoded],
        ),
      };
    case "array":
      return { tag: "array", items: encoded.items.map(orderedTagged) };
    case "date":
    case "undefined":
    case "hole":
    case "uint8Array":
    case "arrayBuffer":
      return encoded;
    default: {
      const unreachable: never = encoded;
      return unreachable;
    }
  }
};
const snapshot = (value: unknown) => orderedTagged(encodeTagged(withoutCaptureSymbols(value)));

const PRESENCE = ["absent", "undefined", "present"] as const;
type Presence = (typeof PRESENCE)[number];
const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const THEME_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const EXTENDED_REL = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";

const seed = (presence: Presence) => {
  const bytes = new Uint8Array([0, 127, 255]).buffer;
  const packageValue = {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "2F48B967",
          content: [
            {
              type: "run",
              preservedAttributes: [{ name: "rsidR", value: "00E10F27" }],
              content: [{ type: "text", text: "aé😀bc" }],
            },
          ],
        },
        { type: "paragraph", paraId: "32BC4F19", content: [] },
      ],
    },
    styles: undefined,
    numbering: undefined,
    properties: { created: new Date("2026-10-05T09:00:00.123Z") },
    media: new Map([
      [
        "word/media/retained.png",
        { path: "word/media/retained.png", mimeType: "image/png", data: bytes },
      ],
    ]),
  } satisfies Document["package"];
  switch (presence) {
    case "absent":
      return { package: packageValue };
    case "undefined":
      return { package: { ...packageValue, relationships: undefined } };
    case "present":
      return {
        package: {
          ...packageValue,
          relationships: new Map([
            ["rId9", { id: "rId9", type: IMAGE_REL, target: "media/retained.png" }],
            ["rId2", { id: "rId2", type: THEME_REL, target: "theme/theme1.xml" }],
          ]),
        },
      };
    default: {
      const unreachable: never = presence;
      return unreachable;
    }
  }
};

const comment = (id: number, paraId: string) =>
  ({
    id,
    author: "Author",
    initials: "AA",
    date: "2026-01-02T03:04:05Z",
    content: [
      {
        type: "paragraph",
        paraId,
        content: [{ type: "run", content: [{ type: "text", text: "Comment" }] }],
      },
    ],
    preserved: { children: [{ index: 0, xml: '<opaque xmlns="urn:fixture"/>' }] },
  }) satisfies Comment;
const position = (offset: number) => ({ story: "main", blockId: "2F48B967", offset }) as const;
const range = (from: number, to: number) =>
  ({
    type: "createComment",
    comment: comment(90, "70000010"),
    anchor: { kind: "range", from: position(from), to: position(to) },
  }) as const satisfies DocumentOp;
const reply = {
  type: "createComment",
  comment: comment(3, "70000011"),
  anchor: { kind: "reply", parentId: 90 },
} as const satisfies DocumentOp;
const point = {
  type: "createComment",
  comment: comment(5, "70000012"),
  anchor: { kind: "point", at: { story: "main", blockId: "32BC4F19", offset: 0 } },
} as const satisfies DocumentOp;

const phase = (document: unknown, ops: readonly DocumentOp[], expected: unknown) => {
  const response = native(document, ops);
  if (typeof response !== "object" || response === null || !("harness" in response))
    throw new HarnessCodecError({
      message: `Pinned comment phase failed: ${JSON.stringify(response)}`,
    });
  const actual = decodeTagged(response.harness);
  deepStrictEqual(snapshot(actual), snapshot(expected));
  if (typeof actual !== "object" || actual === null || !("document" in actual))
    throw new HarnessCodecError({ message: "A comment Applied result needs a document." });
  return actual.document;
};

const exactSequence = (document: Document, operations: readonly DocumentOp[]) => {
  const expected = oracle.applyDocumentOps(document, operations).unwrap();
  const changed = phase(document, operations, expected);
  const undo = oracle.applyDocumentOps(expected.document, expected.inverse).unwrap();
  const restored = phase(changed, expected.inverse, undo);
  deepStrictEqual(snapshot(restored), snapshot(document));
  const redo = oracle.applyDocumentOps(undo.document, undo.inverse).unwrap();
  const redone = phase(restored, undo.inverse, redo);
  deepStrictEqual(snapshot(redone), snapshot(expected.document));
  return expected;
};

test("comment oracle is exactly the pinned schema-10 source", () => {
  expect(COMMENT_ORACLE_COMMIT).toBe("617c0f4703a6cf43adb7ecda34366c29b93cc60f");
  expect(types.DOCUMENT_OP_SCHEMA_VERSION).toBe(10);
  assertCommentOraclePrepared();
});

test.each(PRESENCE)(
  "every scalar span preserves full comment inverse closure and relationship presence (%s)",
  (presence) => {
    const boundaries = [0, 1, 2, 4, 5, 6];
    for (const from of boundaries)
      for (const to of boundaries.filter((offset) => offset >= from))
        exactSequence(seed(presence), [range(from, to)]);
  },
);

test.each(PRESENCE)(
  "point, cross-block span, reply and thread lifecycle match the full tagged oracle (%s)",
  (presence) => {
    exactSequence(seed(presence), [point]);
    const crossBlock = {
      type: "createComment",
      comment: comment(90, "70000010"),
      anchor: {
        kind: "range",
        from: position(2),
        to: { story: "main", blockId: "32BC4F19", offset: 0 },
      },
    } as const satisfies DocumentOp;
    exactSequence(seed(presence), [crossBlock]);
    const created = exactSequence(seed(presence), [range(1, 5), reply, point]);
    exactSequence(created.document, [{ type: "deleteComment", id: 3, scope: "reply" }]);
    exactSequence(created.document, [{ type: "deleteComment", id: 90, scope: "thread" }]);
    exactSequence(seed(presence), [
      range(1, 5),
      reply,
      { type: "deleteComment", id: 90, scope: "thread" },
    ]);
  },
);

test("comment parts preserve Map order and restore existing relationship indices and metadata", () => {
  const base = seed("present");
  const document = {
    package: {
      ...base.package,
      relationships: new Map([
        ["rId8", { id: "rId8", type: EXTENDED_REL, target: "commentsExtended.xml" }],
        ["rId3", { id: "rId3", type: IMAGE_REL, target: "media/retained.png" }],
        ["rId5", { id: "rId5", type: COMMENTS_REL, target: "comments.xml" }],
      ]),
    },
  } satisfies Document;
  exactSequence(document, [range(1, 5)]);
});

const assertRefused = (document: Document, operations: readonly DocumentOp[]) => {
  const expected = oracle.applyDocumentOps(document, operations);
  expect(expected.isErr()).toBe(true);
  if (expected.isErr())
    deepStrictEqual(native(document, operations), {
      status: "refused",
      opType: expected.error.opType,
      reason: expected.error.reason,
      message: expected.error.message,
    });
};

test("comment ownership, identity, fresh paragraph, stale and hostile inverse refusals match exactly", () => {
  const original = seed("present");
  const created = oracle.applyDocumentOps(original, [range(1, 5), reply]).unwrap();
  assertRefused(created.document, [range(1, 5)]);
  assertRefused(original, [
    {
      type: "createComment",
      comment: comment(2_147_483_648, "70000010"),
      anchor: { kind: "point", at: position(0) },
    },
  ]);
  assertRefused(original, [
    {
      type: "createComment",
      comment: comment(90, "00000000"),
      anchor: { kind: "point", at: position(0) },
    },
  ]);
  assertRefused(original, [
    {
      type: "createComment",
      comment: { ...comment(90, "70000010"), content: [] },
      anchor: { kind: "point", at: position(0) },
    },
  ]);
  assertRefused(original, [
    {
      type: "createComment",
      comment: comment(90, "70000010"),
      anchor: { kind: "reply", parentId: 999 },
    },
  ]);
  assertRefused(created.document, [{ type: "deleteComment", id: 999, scope: "thread" }]);
  assertRefused(created.document, [{ type: "deleteComment", id: 90, scope: "reply" }]);
  assertRefused(created.document, [{ type: "deleteComment", id: 3, scope: "thread" }]);
  // A successful early operation must not escape an atomic later refusal.
  assertRefused(original, [range(1, 5), { type: "deleteComment", id: 999, scope: "thread" }]);
  const owned = oracle.applyDocumentOps(original, [range(1, 5)]).unwrap();
  const inverse = owned.inverse.at(0);
  if (!inverse || inverse.type !== "restoreCommentState")
    throw new HarnessCodecError({ message: "Expected a comment state inverse." });
  const stale = {
    ...inverse,
    expected: {
      ...inverse.expected,
      records: inverse.expected.records.map((entry) => ({
        ...entry,
        comment: { ...entry.comment, author: "stale" },
      })),
    },
  };
  assertRefused(owned.document, [stale]);
  assertRefused(owned.document, [{ ...inverse, scaffoldIds: { revision: [123], control: [] } }]);
  const replaced = {
    ...inverse,
    state: {
      ...inverse.state,
      anchors: inverse.state.anchors.map(({ story, blockId, content }) => ({
        story,
        blockId,
        content: [
          { type: "run", content: [{ type: "text", text: "forged" }] },
        ] satisfies typeof content,
      })),
    },
  };
  assertRefused(owned.document, [replaced]);
});

test("comment list owned undefined remains explicitly outside current Rust sidecar relocation", () => {
  const base = seed("present");
  const document = {
    package: { ...base.package, document: { ...base.package.document, comments: undefined } },
  } satisfies Document;
  expect(oracle.applyDocumentOps(document, [range(1, 5)]).isOk()).toBe(true);
  deepStrictEqual(native(document, [range(1, 5)]), {
    status: "unsupported",
    opType: "createComment",
    dimension: "harnessSidecarMovement",
  });
});
