/**
 * A suggested paste of a paragraph whose copied deletion or move away of its
 * break holds a cut of its own: the copied removal goes beside the paste's
 * insertion of the break, and accepting or rejecting the paste and the
 * copied removal, in either order or at once, leaves one document; accepting
 * both leaves what the direct paste accepted leaves.
 */
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
  ParagraphMarkChange,
  Run,
} from "../../model/document";
import { applyDocumentOps } from "../apply";
import { compileEditorIntent, type EditorIntent } from "../editorIntent";
import { identityKeysIn } from "../ids";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

setDefaultTimeout(propertyTestTimeout(120_000));

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});
const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});
const apply = (document: Document, ops: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, ops);
  if (result.isErr()) throw result.error;
  return result.value.document;
};
const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
) =>
  apply(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds,
      decision,
    },
  ]);
const revisionIdsOf = (document: Document): number[] => [
  ...new Set(
    identityKeysIn(document.package.document).flatMap((key) =>
      key.startsWith("revision:")
        ? [Number(key.slice("revision:".length))]
        : [],
    ),
  ),
];

/** The document without paragraph identities, which each order mints apart. */
const shapeOf = (value: unknown): unknown => {
  if (Array.isArray(value)) return value.map(shapeOf);
  if (typeof value !== "object" || value === null) return value;
  const out: Record<string, unknown> = {};
  for (const [key, field] of Object.entries(value))
    if (key !== "paraId") out[key] = shapeOf(field);
  return out;
};
const shape = (document: Document): string =>
  JSON.stringify(shapeOf(document.package.document.content));

const STAMP = { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" };
const NEW_IDS = {
  revision: Array.from({ length: 24 }, (_, index) => 101 + index),
};
const EARLIER = { author: "Earlier", date: "2026-01-01T00:00:00Z" };

/**
 * A copied piece either side of the cut: a link, which a cut of depth one
 * joins, and whose run a cut of depth two joins too.
 */
const pieces = (
  text: string,
  side: "left" | "right",
  record: boolean,
): ParagraphContent[] => {
  const link: ParagraphContent = {
    type: "hyperlink",
    rId: "rId9",
    href: "https://example.org",
    children: [run(text)],
  };
  if (!record) return [link];
  return side === "left"
    ? [link, { type: "bookmarkStart", id: 7, name: "copied" }]
    : [{ type: "bookmarkEnd", id: 7 }, link];
};

type Case = {
  left: string;
  right: string;
  depth: 1 | 2;
  kind: "del" | "moveFrom";
  splitDate: boolean;
  seamRecord: boolean;
  copied: RevisionDecision;
  pasted: RevisionDecision;
};

const intentFor = ({
  left,
  right,
  depth,
  kind,
  splitDate,
  seamRecord,
}: Case) => {
  const pPrMark: ParagraphMarkChange = {
    kind,
    info: { id: 93, ...EARLIER },
    resolutionJoin: depth,
    ...(splitDate ? { splitDate: "2025-12-01T00:00:00Z" } : {}),
  };
  const hasSeamRecord = seamRecord && left !== "" && right !== "";
  const paragraphs: Paragraph[] = [
    {
      type: "paragraph",
      paraId: "0000000D",
      pPrMark,
      content: left === "" ? [] : pieces(left, "left", hasSeamRecord),
    },
    {
      type: "paragraph",
      paraId: "0000000E",
      content: right === "" ? [] : pieces(right, "right", hasSeamRecord),
    },
    { type: "paragraph", content: [run("tail")] },
  ];
  const intent: EditorIntent = {
    type: "replaceFragment",
    from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 5 },
    to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 5 },
    openStart: 0,
    openEnd: 0,
    paragraphs,
  };
  return intent;
};

const destination = (): Document =>
  documentOf([
    { type: "paragraph", paraId: "00000001", content: [run("Here.")] },
  ]);

const caseArbitrary: fc.Arbitrary<Case> = fc.record({
  // An empty side holds no piece for the cut to join.
  left: fc.stringMatching(/^[a-z]{0,5}$/u),
  right: fc.stringMatching(/^[a-z]{0,5}$/u),
  depth: fc.constantFrom<1 | 2>(1, 2),
  kind: fc.constantFrom<"del" | "moveFrom">("del", "moveFrom"),
  splitDate: fc.boolean(),
  seamRecord: fc.boolean(),
  copied: fc.constantFrom(REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT),
  pasted: fc.constantFrom(REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT),
});

test("a copied removal holding a cut resolves alike in every order beside the paste's break", () => {
  assertProperty(
    fc.property(caseArbitrary, (input) => {
      const intent = intentFor(input);
      const suggested = compileEditorIntent(destination(), {
        intent,
        mode: { type: "suggesting", revision: STAMP, newIds: NEW_IDS },
      });
      if (suggested.isErr()) throw suggested.error;
      const pasted = apply(destination(), suggested.value.ops);
      const mark = pasted.package.document.content
        .flatMap((block) => (block.type === "paragraph" ? [block] : []))
        .find((block) => block.pPrMark?.deletion !== undefined)?.pPrMark;
      expect(mark?.kind).toBe("ins");
      expect(mark?.deletion?.kind).toBe(
        input.kind === "moveFrom" ? "moveFrom" : undefined,
      );
      const copied = mark?.deletion?.info.id ?? -1;
      const all = revisionIdsOf(pasted);
      const own = all.filter((id) => id !== copied);
      const outcomes = new Set([
        shape(
          resolve(resolve(pasted, [copied], input.copied), own, input.pasted),
        ),
        shape(
          resolve(resolve(pasted, own, input.pasted), [copied], input.copied),
        ),
      ]);
      if (input.copied === input.pasted)
        outcomes.add(shape(resolve(pasted, all, input.pasted)));
      expect(outcomes.size).toBe(1);
      if (
        input.copied === REVISION_DECISIONS.ACCEPT &&
        input.pasted === REVISION_DECISIONS.ACCEPT
      ) {
        const direct = compileEditorIntent(destination(), {
          intent,
          mode: { type: "editing", newIds: NEW_IDS },
        });
        if (direct.isErr()) throw direct.error;
        const copiedDirect = apply(destination(), direct.value.ops);
        expect([...outcomes]).toEqual([
          shape(
            resolve(
              copiedDirect,
              revisionIdsOf(copiedDirect),
              REVISION_DECISIONS.ACCEPT,
            ),
          ),
        ]);
      }
      if (
        input.pasted === REVISION_DECISIONS.REJECT &&
        // Rejection retains pasted range boundaries as zero-width positions.
        !(input.seamRecord && input.left !== "" && input.right !== "")
      )
        expect([...outcomes]).toEqual([shape(destination())]);
    }),
    { id: "copied cut mark beside a pasted break", numRuns: 60 },
  );
});
