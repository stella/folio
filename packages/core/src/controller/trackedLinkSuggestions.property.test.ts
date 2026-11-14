/**
 * A suggested link change against the same change made directly: rejecting
 * every change gives what rejecting them in the original gives, accepting
 * every change gives what accepting them after the direct edit gives, the
 * pending suggestion saves and reopens as it was, and its inverse restores
 * the original. A suggestion the engine refuses names the shape it cannot
 * record.
 */

import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  applyDocumentOp,
  applyDocumentOps,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  DOCUMENT_OP_REFUSAL_REASONS,
  DOCUMENT_OP_TYPES,
  type EditorIntent,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type {
  Document,
  Hyperlink,
  ParagraphContent,
  Run,
} from "../types/document";

setDefaultTimeout(propertyTestTimeout(120_000));

const PARA_ID = "12345678";
const DATE = "2026-01-01T00:00:00Z";
const TYPED_REFUSALS = [
  "A link suggestion cannot strike text beside a marker inside the same link.",
  "A link suggestion cannot copy content that names a record of its own, such as a drawing or a note reference.",
];

type Piece =
  | { kind: "run"; text: string; bold: boolean; styled: boolean }
  | {
      kind: "link";
      href: string;
      texts: readonly string[];
      marker: "none" | "start" | "middle" | "end";
    }
  | { kind: "bookmark" }
  | { kind: "insertion" | "deletion"; text: string; linked: boolean };

const pieceArbitrary: fc.Arbitrary<Piece> = fc.oneof(
  fc.record({
    kind: fc.constant("run" as const),
    text: fc.constantFrom("ab", "c", "de"),
    bold: fc.boolean(),
    styled: fc.boolean(),
  }),
  fc.record({
    kind: fc.constant("link" as const),
    href: fc.constantFrom("https://a.example/", "https://b.example/"),
    texts: fc.array(fc.constantFrom("xy", "z"), { minLength: 1, maxLength: 3 }),
    marker: fc.constantFrom("none", "none", "start", "middle", "end"),
  }),
  fc.record({ kind: fc.constant("bookmark" as const) }),
  fc.record({
    kind: fc.constantFrom("insertion" as const, "deletion" as const),
    text: fc.constantFrom("pq", "r"),
    linked: fc.boolean(),
  }),
);

const textRun = (
  text: string,
  formatting?: Run["formatting"],
): ParagraphContent =>
  formatting === undefined
    ? { type: "run", content: [{ type: "text", text }] }
    : { type: "run", formatting, content: [{ type: "text", text }] };

/** A paragraph from pieces, with unique bookmark and revision ids. */
const documentOf = (pieces: readonly Piece[]): Document => {
  let bookmark = 1;
  let revision = 100;
  const content: ParagraphContent[] = [];
  for (const piece of pieces) {
    switch (piece.kind) {
      case "run":
        content.push(
          textRun(piece.text, {
            ...(piece.bold ? { bold: true } : {}),
            ...(piece.styled
              ? { styleId: "Hyperlink" }
              : { color: { rgb: "123456" } }),
          }),
        );
        break;
      case "link": {
        const children: Hyperlink["children"] = piece.texts.map(
          (text): Run => ({ type: "run", content: [{ type: "text", text }] }),
        );
        const id = bookmark++;
        const markers = [
          { type: "bookmarkStart", id, name: `mark${String(id)}` } as const,
          { type: "bookmarkEnd", id } as const,
        ];
        const at = {
          none: undefined,
          start: 0,
          middle: Math.ceil(children.length / 2),
          end: children.length,
        }[piece.marker];
        if (at !== undefined) children.splice(at, 0, ...markers);
        content.push({ type: "hyperlink", href: piece.href, children });
        break;
      }
      case "bookmark": {
        const id = bookmark++;
        content.push(
          { type: "bookmarkStart", id, name: `mark${String(id)}` },
          { type: "bookmarkEnd", id },
        );
        break;
      }
      case "insertion":
      case "deletion":
        content.push({
          type: piece.kind,
          info: { id: revision++, author: "Other", date: DATE },
          content: [
            piece.linked
              ? {
                  type: "hyperlink",
                  href: "https://c.example/",
                  children: [textRun(piece.text)],
                }
              : textRun(piece.text),
          ],
        });
        break;
      default: {
        const unreachable: never = piece;
        return unreachable;
      }
    }
  }
  return {
    package: {
      document: {
        content: [{ type: "paragraph", paraId: PARA_ID, content }],
      },
    },
  };
};

const physicalLength = (pieces: readonly Piece[]): number => {
  let length = 0;
  for (const piece of pieces) {
    switch (piece.kind) {
      case "run":
      case "insertion":
      case "deletion":
        length += piece.text.length;
        break;
      case "link":
        length += piece.texts.join("").length;
        break;
      case "bookmark":
        break;
      default: {
        const unreachable: never = piece;
        return unreachable;
      }
    }
  }
  return length;
};

/** Every revision id a document holds, wherever its record sits. */
const revisionIds = (value: unknown, out = new Set<number>()): Set<number> => {
  if (Array.isArray(value)) {
    for (const item of value) revisionIds(item, out);
    return out;
  }
  if (typeof value !== "object" || value === null) return out;
  for (const [key, field] of Object.entries(value)) {
    if (
      key === "info" &&
      typeof field === "object" &&
      field !== null &&
      "id" in field &&
      typeof field.id === "number"
    )
      out.add(field.id);
    revisionIds(field, out);
  }
  return out;
};

const resolvedAll = (
  document: Document,
  decision: RevisionDecision,
): Document => {
  const ids = [...revisionIds(document.package.document)];
  if (ids.length === 0) return document;
  return applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: ids,
    decision,
  }).unwrap().document;
};

type Token =
  | {
      text: string;
      link: string | undefined;
      change: "insertion" | "deletion" | undefined;
      bold: boolean;
      color: string | undefined;
      styleId: string | undefined;
    }
  | { marker: "bookmarkStart" | "bookmarkEnd"; id: number }
  | { emptyLink: string | undefined };

/**
 * What a reader sees, character by character, and where each bookmark
 * stands: independent of how runs and link records are split.
 */
const authored = (document: Document): Token[] => {
  const tokens: Token[] = [];
  const walk = (
    items: readonly ParagraphContent[],
    link: string | undefined,
    change: "insertion" | "deletion" | undefined,
  ): void => {
    for (const item of items) {
      switch (item.type) {
        case "run":
          for (const part of item.content)
            if (part.type === "text")
              for (const text of part.text)
                tokens.push({
                  text,
                  link,
                  change,
                  bold: item.formatting?.bold === true,
                  color: item.formatting?.color?.rgb,
                  styleId: item.formatting?.styleId,
                });
          break;
        case "hyperlink":
          if (item.children.length === 0) tokens.push({ emptyLink: item.href });
          walk(item.children, item.href, change);
          break;
        case "insertion":
        case "deletion":
          walk(item.content, link, item.type);
          break;
        case "bookmarkStart":
        case "bookmarkEnd":
          tokens.push({ marker: item.type, id: item.id });
          break;
        default:
          throw new TypeError(`Unexpected ${item.type} in a link fixture.`);
      }
    }
  };
  for (const block of document.package.document.content)
    if (block.type === "paragraph") walk(block.content, undefined, undefined);
  return tokens;
};

const intentOf = (
  type: "setHyperlink" | "removeHyperlink" | "insertHyperlink",
  from: number,
  to: number,
): EditorIntent => {
  const position = (offset: number) => ({
    story: OP_STORIES.MAIN,
    blockId: PARA_ID,
    offset,
  });
  switch (type) {
    case "setHyperlink":
      return {
        type,
        from: position(from),
        to: position(to),
        href: "https://n.example/",
      };
    case "removeHyperlink":
      return { type, from: position(from), to: position(to) };
    case "insertHyperlink":
      return {
        type,
        from: position(from),
        to: position(to),
        text: "New",
        href: "https://n.example/",
      };
    default: {
      const unreachable: never = type;
      return unreachable;
    }
  }
};

const compiled = (
  document: Document,
  intent: EditorIntent,
  suggesting: boolean,
) => {
  const ids = createEditorIntentIdAllocator()(document, intent);
  return compileEditorIntent(document, {
    intent,
    mode: suggesting
      ? {
          type: "suggesting",
          revision: { id: ids.revisionId, author: "Suggester", date: DATE },
          newIds: ids.newIds,
        }
      : { type: "editing", newIds: ids.newIds },
  }).andThen(({ ops }) =>
    applyDocumentOps(document, ops).map((applied) => ({ ops, applied })),
  );
};

/** Check one link edit both ways; the typed refusal it met, if any. */
const checkLinkEdit = async (
  document: Document,
  intent: EditorIntent,
): Promise<string | undefined> => {
  const direct = compiled(document, intent, false);
  if (direct.isErr()) return undefined;
  const tracked = compiled(document, intent, true);
  if (tracked.isErr()) {
    expect(tracked.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
    expect(TYPED_REFUSALS).toContain(tracked.error.message);
    return tracked.error.message;
  }
  const suggested = tracked.value.applied.document;
  const edited = direct.value.applied.document;
  // Undo of the suggestion restores the original exactly.
  expect(
    applyDocumentOps(suggested, tracked.value.applied.inverse).unwrap()
      .document,
  ).toEqual(document);
  expect(authored(resolvedAll(suggested, REVISION_DECISIONS.REJECT))).toEqual(
    authored(resolvedAll(document, REVISION_DECISIONS.REJECT)),
  );
  expect(authored(resolvedAll(suggested, REVISION_DECISIONS.ACCEPT))).toEqual(
    authored(resolvedAll(edited, REVISION_DECISIONS.ACCEPT)),
  );
  // Saved and reopened, the suggestion reads and resolves as it did.
  const reopened = await parseDocx(await createDocx(suggested));
  expect(authored(reopened)).toEqual(authored(suggested));
  for (const decision of [REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT])
    expect(authored(resolvedAll(reopened, decision))).toEqual(
      authored(resolvedAll(suggested, decision)),
    );
  return undefined;
};

test("a suggested link change resolves to the direct edit or the original, saved or not", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.array(pieceArbitrary, { minLength: 1, maxLength: 6 }),
      fc.constantFrom(
        "setHyperlink" as const,
        "removeHyperlink" as const,
        "insertHyperlink" as const,
      ),
      fc.nat(),
      fc.nat(),
      async (pieces, type, left, right) => {
        const length = physicalLength(pieces);
        const from = Math.min(left % (length + 1), right % (length + 1));
        const to = Math.max(left % (length + 1), right % (length + 1));
        await checkLinkEdit(documentOf(pieces), intentOf(type, from, to));
      },
    ),
    {
      numRuns: 120,
      id: "a suggested link change resolves to the direct edit or the original, saved or not",
    },
  );
});

test("a suggested link change over a link split by a marker is a typed refusal", async () => {
  // Accepting would leave the marker in a link around no text, where the
  // direct edit leaves the bare marker between the two unlinked runs.
  const document = documentOf([
    {
      kind: "link",
      href: "https://a.example/",
      texts: ["xy", "z"],
      marker: "middle",
    },
  ]);
  expect(await checkLinkEdit(document, intentOf("removeHyperlink", 0, 3))).toBe(
    TYPED_REFUSALS[0],
  );
  // Over text beside the marker only, the link keeps text on that side.
  expect(
    await checkLinkEdit(document, intentOf("removeHyperlink", 0, 1)),
  ).toBeUndefined();
});

test("suggested link removal and setting over plain and linked text record the change", async () => {
  for (const pieces of [
    [{ kind: "run", text: "ab", bold: true, styled: true }],
    [
      {
        kind: "link",
        href: "https://a.example/",
        texts: ["xy", "z"],
        marker: "none",
      },
    ],
    [
      { kind: "run", text: "ab", bold: false, styled: false },
      { kind: "bookmark" },
      {
        kind: "link",
        href: "https://b.example/",
        texts: ["z"],
        marker: "none",
      },
      { kind: "deletion", text: "pq", linked: true },
      { kind: "insertion", text: "r", linked: false },
    ],
  ] satisfies Piece[][]) {
    const length = physicalLength(pieces);
    for (const type of [
      "setHyperlink",
      "removeHyperlink",
      "insertHyperlink",
    ] as const)
      expect(
        await checkLinkEdit(documentOf(pieces), intentOf(type, 0, length)),
      ).toBeUndefined();
  }
});
