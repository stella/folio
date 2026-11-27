import { expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { commentDocumentIssue, sourceCommentAnchorsIn } from "../comments";
import { validateOpsDocument } from "../contract";
import { leafSpans } from "../leaves";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { OP_STORIES } from "../types";
import { planWholeParagraphReplacement } from "../wholeParagraphReplacement";

const reference = {
  type: "commentReference",
  id: 7,
} satisfies ParagraphContent;
const content = [
  { type: "commentRangeStart", id: 7 },
  { type: "run", content: [{ type: "text", text: "Host" }] },
  { type: "commentRangeEnd", id: 7 },
  reference,
] satisfies ParagraphContent[];
const documentOf = (first: ParagraphContent[] = content): Document => ({
  package: {
    document: {
      content: [
        { type: "paragraph", paraId: "00000001", content: first },
        {
          type: "paragraph",
          paraId: "00000002",
          content: [{ type: "run", content: [{ type: "text", text: "Tail" }] }],
        },
      ],
      comments: [
        { id: 7, author: "Reviewer", anchorKind: "range", content: [] },
      ],
    },
  },
});
const plan = (document: Document) =>
  planWholeParagraphReplacement({
    document,
    from: {
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      offset: 0,
      zeroWidthBefore: 0,
    },
    to: {
      story: OP_STORIES.MAIN,
      blockId: "00000002",
      offset: 0,
      zeroWidthBefore: 0,
    },
  }).unwrap();
const field = (anchor: ParagraphContent) =>
  ({
    type: "simpleField",
    instruction: " REF target ",
    fieldType: "REF",
    content: [
      {
        type: "inlineWrapper",
        kind: "bidi",
        control: "embedding",
        content: [anchor],
      },
    ],
  }) satisfies ParagraphContent;
const tailOf = (document: Document): Paragraph => {
  const tail = document.package.document.content.at(-1);
  if (tail?.type !== "paragraph")
    throw new TypeError("The fixture has its tail");
  return tail;
};

test("complete source range ownership certifies only the retiring paragraph", () => {
  const document = documentOf();
  expect(validateOpsDocument(document).isOk()).toBe(true);
  expect(commentDocumentIssue(document, "all")).toBeUndefined();
  expect(plan(document)).toEqual({
    blockIds: ["00000001"],
    commentIds: new Set([7]),
  });
});

test("an empty pending wrapper is itself a leaf and cannot certify retired host ownership", () => {
  const pending = {
    type: "insertion",
    info: { id: 501, author: "Writer" },
    content: [],
  } satisfies ParagraphContent;
  const document = documentOf([...content, pending]);
  const spans = leafSpans([...content, pending]);
  expect(
    spans.some(
      ({ node, ancestors }) => node === pending && ancestors.length === 0,
    ),
  ).toBe(true);
  // This pure preflight must fail closed even before the seed contract normalizes empty records.
  expect(plan(document)).toBeNull();
});

test("a source reference outside the retiring range endpoints retains refusal", () => {
  const document = documentOf(content.filter((node) => node !== reference));
  tailOf(document).content.push(reference);
  expect(commentDocumentIssue(document, "all")).toBeUndefined();
  expect(plan(document)).toBeNull();
});

test("a duplicate source object inside an atomic sibling counts as another occurrence", () => {
  const document = documentOf();
  tailOf(document).content.push(field(reference));
  expect(validateOpsDocument(document).isOk()).toBe(true);
  // Addressable comment validation intentionally does not descend into field results.
  expect(commentDocumentIssue(document, "all")).toBeUndefined();
  expect(
    sourceCommentAnchorsIn(document).filter((anchor) => anchor === reference),
  ).toHaveLength(2);
  expect(plan(document)).toBeNull();
});

test("an unrelated atomic sibling anchor does not claim the retiring comment", () => {
  const document = documentOf();
  tailOf(document).content.push(field({ type: "commentReference", id: 9 }));
  document.package.document.comments?.push({
    id: 9,
    author: "Other",
    anchorKind: "point",
    content: [],
  });
  expect(plan(document)).toEqual({
    blockIds: ["00000001"],
    commentIds: new Set([7]),
  });
});

test("a source boundary gap omitted after leading markers is not whole-host selection", () => {
  const document = documentOf();
  expect(
    planWholeParagraphReplacement({
      document,
      from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
      to: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 0 },
    }).unwrap(),
  ).toBeNull();
});

test("source retirement cannot cross story or parent block-list ownership", () => {
  const document = documentOf();
  const first = document.package.document.content.at(0);
  if (first?.type !== "paragraph")
    throw new TypeError("The fixture has its host");
  document.package.document.content[0] = {
    type: "blockSdt",
    properties: { id: 33 },
    content: [first],
  };
  expect(plan(document)).toBeNull();
  const crossed = planWholeParagraphReplacement({
    document,
    from: {
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      offset: 0,
      zeroWidthBefore: 0,
    },
    to: {
      story: { kind: "header", rId: "rIdHeader" },
      blockId: "00000009",
      offset: 0,
    },
  });
  expect(crossed.isErr() && crossed.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.CROSS_BLOCK_RANGE,
  );
});

const firstOf = (document: Document): Paragraph => {
  const first = document.package.document.content.at(0);
  if (first?.type !== "paragraph")
    throw new TypeError("The fixture has its host");
  return first;
};

test("pending insertions, words or break, retire with their host", () => {
  const info = { id: 501, author: "Drafter" };
  const inserted = {
    type: "insertion",
    info,
    content: [{ type: "run", content: [{ type: "text", text: "New " }] }],
  } satisfies ParagraphContent;
  const document = documentOf([inserted, ...content]);
  firstOf(document).pPrMark = { kind: "ins", info: { ...info, id: 503 } };
  expect(validateOpsDocument(document).isOk()).toBe(true);
  expect(plan(document)).toEqual({
    blockIds: ["00000001"],
    commentIds: new Set([7]),
  });
  const suggested = (author: string, suggestionId?: string) =>
    planWholeParagraphReplacement({
      document,
      from: {
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        offset: 0,
        zeroWidthBefore: 0,
      },
      to: {
        story: OP_STORIES.MAIN,
        blockId: "00000002",
        offset: 0,
        zeroWidthBefore: 0,
      },
      suggesting: {
        author,
        ...(suggestionId === undefined ? {} : { suggestionId }),
      },
    }).unwrap();
  // A suggestion retracts its author's own insertions, not another's, nor
  // an ordinary insertion of the same author, which rejecting it gives back.
  expect(suggested("Drafter")).toEqual(plan(document));
  expect(suggested("Reviewer")).toBeNull();
  expect(suggested("Drafter", "s1")).toBeNull();
});

test("pending removals in a host still refuse to certify retirement", () => {
  const info = { id: 502, author: "Drafter" };
  const removed = {
    type: "deletion",
    info,
    content: [{ type: "run", content: [{ type: "deletedText", text: "Old" }] }],
  } satisfies ParagraphContent;
  expect(plan(documentOf([removed, ...content]))).toBeNull();
  const document = documentOf();
  firstOf(document).pPrMark = { kind: "del", info };
  expect(plan(document)).toBeNull();
});
