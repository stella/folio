import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import { FolioDocxReviewer } from "../../packages/core/src/ai-edits/headless";
import { prepareCanonicalDocxInput } from "../../packages/core/src/docx/canonicalSessionInput";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { createDocx } from "../../packages/core/src/docx/rezip";
import type {
  BlockContent,
  Document,
  ParagraphContent,
} from "../../packages/docx-core/src/model/document";
import { reviewDocumentArbitrary } from "../../packages/docx-core/src/ops/__tests__/documentArbitraries";
import { storyParagraphs } from "../../packages/docx-core/src/ops/blocks";
import {
  commentAnchorsIn,
  commentDocumentIssue,
} from "../../packages/docx-core/src/ops/comments";
import { contractViolation } from "../../packages/docx-core/src/ops/contract";
import { packageIdentityKeys } from "../../packages/docx-core/src/ops/ids";
import {
  isCommentAnchor,
  leafSpans,
} from "../../packages/docx-core/src/ops/leaves";
import { paragraphLogicalText } from "../../packages/docx-core/src/ops/offsets";
import { withCoherentSyntheticReviewComments } from "../../test/generators/coherentReviewerComments";
import {
  bodyReviewerDocumentArbitrary,
  reviewerDocumentArbitrary,
} from "../../test/generators/reviewerDocumentArbitraries";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../test/property-testing";

const paragraph = (content: ParagraphContent[]) => ({
  type: "paragraph" as const,
  paraId: "00000002",
  content,
});
const source = (content: BlockContent[]): Document => {
  const properties = { pageWidth: 12240, pageHeight: 15840 };
  return {
    package: {
      document: {
        content,
        sections: [{ properties, content: structuredClone(content) }],
        finalSectionProperties: properties,
        comments: [
          {
            id: 7,
            author: "A",
            content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }],
          },
        ],
      },
      numbering: undefined,
      footnotes: [
        {
          type: "footnote",
          id: 2,
          content: [{ type: "paragraph", paraId: "7FFFFFF1", content: [] }],
        },
      ],
      settings: { defaultTabStop: 720 },
      properties: { title: "synthetic" },
    },
    warnings: ["kept"],
  };
};

/** The complete raw source class of seed289121159, before reviewer adaptation. */
const rawDuplicate = () =>
  source([
    { type: "paragraph", paraId: "00000001", content: [] },
    {
      type: "blockSdt",
      properties: { sdtType: "group", id: 1 },
      content: [
        paragraph([
          { type: "commentReference", id: 7 },
          {
            type: "inlineWrapper",
            kind: "bidi",
            control: "embedding",
            content: [{ type: "commentReference", id: 7 }],
          },
        ]),
        { type: "paragraph", paraId: "00000003", content: [] },
      ],
    },
  ]);

const paragraphFacts = (document: Document) =>
  storyParagraphs(document.package.document).map(({ paragraph: value }) => {
    const { content, ...attributes } = value;
    return {
      attributes,
      text: paragraphLogicalText(value),
      leaves: leafSpans(content)
        .filter(({ node }) => !isCommentAnchor(node))
        .map(({ node }) => node),
    };
  });
const locations = (document: Document) =>
  commentAnchorsIn(document).map(({ node, blockId, before, after, story }) => ({
    type: node.type,
    blockId,
    from: before.offset,
    to: after.offset,
    story,
  }));

const expectPreserved = (before: Document, after: Document) => {
  expect(paragraphFacts(after)).toEqual(paragraphFacts(before));
  expect(packageIdentityKeys(after.package)).toEqual(
    packageIdentityKeys(before.package),
  );
  expect(
    after.package.document.comments?.slice(
      0,
      before.package.document.comments?.length,
    ),
  ).toEqual(before.package.document.comments);
  const remaining = locations(after);
  let index = 0;
  for (const original of locations(before)) {
    while (
      index < remaining.length &&
      JSON.stringify(remaining.at(index)) !== JSON.stringify(original)
    )
      index++;
    expect(remaining.at(index++)).toEqual(original);
  }
};

test(
  "reviewer generators construct native-valid comments without filtering draws",
  () => {
    assertProperty(
      fc.property(
        reviewerDocumentArbitrary,
        bodyReviewerDocumentArbitrary,
        (whole, body) => {
          for (const document of [whole, body]) {
            expect(commentDocumentIssue(document, "all")).toBeUndefined();
            expect(contractViolation(document)).toBeUndefined();
          }
        },
      ),
      {
        id: "reviewer generators construct native-valid comments without filtering draws",
        numRuns: 200,
      },
    );
  },
  propertyTestTimeout(15_000),
);

test(
  "raw review adaptation keeps every source location, text, definition and existing identity",
  () => {
    assertProperty(
      fc.property(reviewDocumentArbitrary, (document) => {
        const original = structuredClone(document);
        const adapted = withCoherentSyntheticReviewComments(document);
        expect(document).toEqual(original);
        expectPreserved(document, adapted);
        expect(commentDocumentIssue(adapted, "all")).toBeUndefined();
        expect(contractViolation(adapted)).toBeUndefined();
        expect(withCoherentSyntheticReviewComments(adapted)).toEqual(adapted);
      }),
      { id: "raw review adaptation preserves source facts", numRuns: 200 },
    );
  },
  propertyTestTimeout(15_000),
);

test("an unadapted duplicate-reference source opens with one reference on both sessions", async () => {
  const raw = rawDuplicate();
  expect(commentDocumentIssue(raw, "all")).toBe(
    "Comment anchors are unbalanced or duplicated.",
  );
  const bytes = (await prepareCanonicalDocxInput(await createDocx(raw)))
    .unwrap()
    .slice().buffer;
  // Opening drops the repeated reference and says so.
  const parsed = await parseDocx(bytes, {
    detectVariables: false,
    preloadFonts: false,
  });
  expect(commentDocumentIssue(parsed, "all")).toBeUndefined();
  expect(parsed.parseWarnings?.map(({ code }) => code)).toContain(
    "duplicate-comment-reference",
  );
  for (const session of ["legacy", "canonical"] as const)
    // oxlint-disable-next-line no-await-in-loop -- one session after the other
    await FolioDocxReviewer.fromBuffer(bytes.slice(0), { session });
});

const cases = {
  "surplus nested point": rawDuplicate,
  "coherent range and surplus point": () =>
    source([
      paragraph([
        { type: "commentRangeStart", id: 7 },
        { type: "run", content: [{ type: "text", text: "Kept coverage" }] },
        { type: "commentRangeEnd", id: 7 },
        { type: "commentReference", id: 7 },
        {
          type: "inlineWrapper",
          kind: "bidi",
          control: "embedding",
          content: [{ type: "commentReference", id: 7 }],
        },
      ]),
    ]),
  "unmatched range atoms and missing definitions": () =>
    source([
      paragraph([
        { type: "commentRangeEnd", id: 19 },
        { type: "commentRangeStart", id: 23 },
        { type: "run", content: [{ type: "text", text: "Standing text" }] },
        { type: "commentReference", id: 29 },
      ]),
    ]),
  "table and nested inline-control atoms": () =>
    source([
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              {
                type: "tableCell",
                content: [
                  paragraph([
                    { type: "commentReference", id: 7 },
                    {
                      type: "inlineSdt",
                      properties: { sdtType: "richText", id: 12 },
                      content: [
                        {
                          type: "hyperlink",
                          href: "https://example.org",
                          children: [
                            {
                              type: "inlineWrapper",
                              kind: "bidi",
                              control: "embedding",
                              content: [
                                { type: "commentReference", id: 7 },
                                { type: "commentRangeEnd", id: 35 },
                              ],
                            },
                          ],
                        },
                      ],
                    },
                  ]),
                ],
              },
            ],
          },
        ],
      },
      { type: "paragraph", paraId: "00000003", content: [] },
    ]),
  "multiple same-id range atoms": () =>
    source([
      paragraph([
        { type: "commentRangeStart", id: 7 },
        { type: "commentRangeStart", id: 7 },
        { type: "run", content: [{ type: "text", text: "Nested text" }] },
        { type: "commentRangeEnd", id: 7 },
        { type: "commentRangeEnd", id: 7 },
      ]),
    ]),
};

for (const [name, fixture] of Object.entries(cases))
  test(`${name} becomes coherent source held by both actual reviewers`, async () => {
    const original = fixture();
    const adapted = withCoherentSyntheticReviewComments(original);
    expectPreserved(original, adapted);
    expect(commentDocumentIssue(adapted, "all")).toBeUndefined();
    const bytes = (await prepareCanonicalDocxInput(await createDocx(adapted)))
      .unwrap()
      .slice().buffer;
    const legacy = await FolioDocxReviewer.fromBuffer(bytes.slice(0), {
      session: "legacy",
    });
    const canonical = await FolioDocxReviewer.fromBuffer(bytes.slice(0), {
      session: "canonical",
    });
    const content = (reviewer: FolioDocxReviewer) =>
      reviewer.getContent().map(({ id, text }) => ({ id, text }));
    expect(content(canonical)).toEqual(content(legacy));
    expect(canonical.getComments()).toEqual(legacy.getComments());
    const expectedIds = adapted.package.document.comments
      ?.map(({ id }) => id)
      .toSorted((a, b) => a - b);
    for (const reviewer of [legacy, canonical]) {
      const saved = await parseDocx(await reviewer.toBuffer(), {
        detectVariables: false,
        preloadFonts: false,
      });
      expect(commentDocumentIssue(saved, "all")).toBeUndefined();
      const anchors = (document: Document) =>
        commentAnchorsIn(document).map(({ node, blockId, before, after }) => ({
          type: node.type,
          id: node.id,
          blockId,
          from: before.offset,
          to: after.offset,
        }));
      expect(anchors(saved)).toEqual(anchors(adapted));
      expect(
        saved.package.document.comments
          ?.map(({ id }) => id)
          .toSorted((a, b) => a - b),
      ).toEqual(expectedIds);
    }
  });

test("blank-root construction derives its identity and author from the raw fixture", () => {
  const input = source([
    paragraph([
      { type: "commentReference", id: 37 },
      { type: "commentReference", id: 37 },
    ]),
  ]);
  const original =
    input.package.document.comments?.at(0) ??
    panic("The synthetic fixture needs its blank root.");
  original.id = 37;
  original.author = "Fixture";
  const blank =
    original.content.at(0) ?? panic("The synthetic root needs its paragraph.");
  blank.paraId = "21212121";
  const output = withCoherentSyntheticReviewComments(input);
  expect(commentDocumentIssue(output, "all")).toBeUndefined();
  expect(output.package.document.comments?.at(0)).toBe(original);
  expect(
    output.package.document.comments?.every(
      ({ author }) => author === "Fixture",
    ),
  ).toBe(true);
  expectPreserved(input, output);
});

test("coherent comment and reply definitions remain untouched", () => {
  const input = source([paragraph([{ type: "commentReference", id: 7 }])]);
  input.package.document.comments?.push({
    id: 8,
    parentId: 7,
    author: "Reply",
    content: [
      {
        type: "paragraph",
        paraId: "21212121",
        content: [
          { type: "run", content: [{ type: "text", text: "Authored reply" }] },
        ],
      },
    ],
  });
  expect(commentDocumentIssue(input, "all")).toBeUndefined();
  expect(withCoherentSyntheticReviewComments(input)).toBe(input);
});

test("repair refuses an invalid arbitrary comment graph outside the blank-root precondition", () => {
  const input = rawDuplicate();
  const root =
    input.package.document.comments?.at(0)?.content.at(0) ??
    panic("The synthetic fixture needs its root paragraph.");
  root.content.push({
    type: "run",
    content: [{ type: "text", text: "Authored body" }],
  });
  expect(() => withCoherentSyntheticReviewComments(input)).toThrow(
    "requires the raw blank-root fixture",
  );
});
