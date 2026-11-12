import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { escapeXmlAttribute } from "@stll/docx-core";
import {
  MAX_REVISION_ID,
  PARAGRAPH_MARK_CHANGE_KINDS,
} from "@stll/docx-core/model";

import { assertExactModel } from "../../../../test/exactModel";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { readParagraphAttrs } from "../prosemirror/attrs";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { schema } from "../prosemirror/schema";
import type {
  Paragraph,
  ParagraphContent,
  ParagraphMarkChange,
  TrackedRunChange,
} from "../types/document";
import { parseParagraph } from "./paragraphParser";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import {
  parseCopyOriginAncestors,
  parseParagraphMarkResolutionJoin,
  parseResolutionJoins,
  parseSplitResolutionPiece,
  reviewProvenanceIssue,
  ReviewResolutionProvenanceError,
  serializeBoundaryJoins,
  serializeCopyOriginAncestors,
  serializeParagraphMarkResolutionJoin,
  serializeResolutionJoins,
  serializeSplitResolutionPiece,
} from "./reviewResolutionProvenance";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializePartElement } from "./serializer/partNamespaces";
import {
  findChild,
  getChildElements,
  getLocalName,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const element = (xml: string) => {
  const parsed = parseXmlDocument(xml);
  if (parsed === null)
    throw new TypeError("The provenance fixture is not XML.");
  return parsed;
};

for (const wordNamespace of WORDPROCESSINGML_NAMESPACE_URIS) {
  test(`generated seam metadata survives ${wordNamespace} with arbitrary bound prefixes`, () => {
    assertProperty(
      fc.property(
        fc.constantFrom("history", "capture", "different"),
        fc.record({
          before: fc.integer({ min: 0, max: 3 }),
          after: fc.integer({ min: 0, max: 3 }),
          remove: fc.integer({ min: 0, max: 3 }),
        }),
        fc.subarray(["before", "after"] as const),
        fc.constantFrom(...PARAGRAPH_MARK_CHANGE_KINDS),
        fc.oneof(
          fc.constantFrom("absent", "undefined"),
          fc.integer({ min: 0, max: 4 }),
          fc.constant(MAX_REVISION_ID),
        ),
        fc.option(
          fc.uniqueArray(
            fc.record({
              depth: fc.integer({ min: 0, max: 4 }),
              blockers: fc.uniqueArray(
                fc.oneof(
                  fc.constant(0),
                  fc.constant(MAX_REVISION_ID),
                  fc.integer({ min: 1, max: MAX_REVISION_ID - 1 }),
                ),
                { minLength: 1, maxLength: 3 },
              ),
            }),
            {
              minLength: 1,
              maxLength: 3,
              selector: (group) => JSON.stringify(group),
            },
          ),
          { nil: undefined },
        ),
        fc.option(fc.constant("merge-plain-runs"), { nil: undefined }),
        (
          prefix,
          depths,
          boundaryJoins,
          kind,
          resolutionJoin,
          deferredRemove,
          acceptance,
        ) => {
          const paragraphMark = {
            kind,
            info: { id: 3, author: "Reviewer" },
            ...(resolutionJoin === "absent"
              ? {}
              : {
                  resolutionJoin:
                    resolutionJoin === "undefined" ? undefined : resolutionJoin,
                }),
          } satisfies ParagraphMarkChange;
          const paragraphMarkAttrs = serializeParagraphMarkResolutionJoin(
            paragraphMark,
          ).replace("folio:", `${prefix}:`);
          const joins = {
            ...depths,
            ...(acceptance === undefined ? {} : { acceptance }),
            ...(deferredRemove === undefined
              ? {}
              : {
                  deferredRemove: deferredRemove.map(({ depth, blockers }) => ({
                    depth,
                    blockers,
                  })),
                }),
            retainedAfter: [
              {
                depth: 0,
                source: [{ space: "revision", id: 20 }],
                target: [{ space: "revision", id: 21 }],
              },
            ],
          } as const;
          const attrs = serializeResolutionJoins(joins).replace(
            "folio:",
            `${prefix}:`,
          );
          const boundaryAttrs = serializeBoundaryJoins(boundaryJoins).replace(
            "folio:",
            `${prefix}:`,
          );
          const paragraph = parseParagraph(
            element(
              `<q:p xmlns:q="${wordNamespace}" xmlns:${prefix}="${FOLIO_REVIEW_HISTORY_NAMESPACE}"><q:pPr><q:rPr><q:${kind} q:id="3" q:author="Reviewer"${paragraphMarkAttrs}/></q:rPr></q:pPr><q:ins q:id="1" q:author="Reviewer"${attrs}><q:r><q:rPr><q:b/><q:rPrChange q:id="2" q:author="Reviewer"${boundaryAttrs}><q:rPr/></q:rPrChange></q:rPr><q:t>text</q:t></q:r></q:ins></q:p>`,
            ),
            null,
            null,
            null,
            null,
            null,
          );
          assertExactModel(paragraph.pPrMark, paragraphMark);
          const pmDoc = toProseDoc({
            package: { document: { content: [paragraph] } },
          });
          const projected = proseDocToBlocks(pmDoc, [paragraph]).at(0);
          if (projected?.type !== "paragraph")
            throw new TypeError("Projected paragraph disappeared.");
          assertExactModel(projected.pPrMark, paragraphMark);
          const wrapper = paragraph.content.at(0);
          if (wrapper?.type !== "insertion")
            throw new TypeError("A tracked wrapper disappeared.");
          assertExactModel(wrapper.resolutionJoins, joins);
          const run = wrapper.content.at(0);
          if (run?.type !== "run")
            throw new TypeError("A reviewed run disappeared.");
          assertExactModel(
            run.propertyChanges?.at(0)?.boundaryJoins,
            boundaryJoins,
          );
          const saved = serializePartElement({
            partPath: "word/document.xml",
            rootName: "w:document",
            baselinePrefixes: ["w"],
            sourceBindings: undefined,
            body: serializeParagraph(paragraph),
          });
          expect(saved).toContain(
            'xmlns:folio="urn:stella:folio:review-history:1"',
          );
          expect(saved).toMatch(/mc:Ignorable="[^"]*folio/u);
          const reopenedNode = findChild(element(saved), "w", "p");
          if (reopenedNode === null)
            throw new TypeError("Saved paragraph disappeared.");
          const reopened = parseParagraph(
            reopenedNode,
            null,
            null,
            null,
            null,
            null,
          );
          assertExactModel(reopened.pPrMark, paragraphMark);
          const reopenedWrapper = reopened.content.at(0);
          if (reopenedWrapper?.type !== "insertion")
            throw new TypeError("Saved wrapper disappeared.");
          assertExactModel(reopenedWrapper.resolutionJoins, joins);
          const reopenedRun = reopenedWrapper.content.at(0);
          if (reopenedRun?.type !== "run")
            throw new TypeError("Saved run disappeared.");
          assertExactModel(
            reopenedRun.propertyChanges?.at(0)?.boundaryJoins,
            boundaryJoins,
          );
        },
      ),
      {
        numRuns: 30,
        id: "generated seam metadata survives ${wordNamespace} with arbitrary bound prefixes",
      },
    );
  });
}

test("malformed and future provenance is dropped on read and reported", () => {
  for (const encoded of [
    "not-json",
    ...[
      null,
      { depth: -1, blockers: [1] },
      { depth: MAX_REVISION_ID + 1, blockers: [1] },
      { depth: 0, blockers: [] },
      { depth: 0, blockers: [1, 1] },
      { depth: 0, blockers: [-1] },
      { depth: 0, blockers: [MAX_REVISION_ID + 1] },
      { depth: 0, blockers: [1], extra: true },
    ].map((deferredRemove) =>
      JSON.stringify({
        version: 1,
        value: { before: 0, after: 0, remove: 0, deferredRemove },
      }),
    ),
    ...[
      [],
      [
        { depth: 0, blockers: [1] },
        { depth: 0, blockers: [1] },
      ],
      [
        { depth: 0, blockers: [1] },
        { depth: -1, blockers: [2] },
      ],
    ].map((deferredRemove) =>
      JSON.stringify({
        version: 1,
        value: { before: 0, after: 0, remove: 0, deferredRemove },
      }),
    ),
    ...["unknown", null, 1, {}, ["merge-plain-runs"]].map((acceptance) =>
      JSON.stringify({
        version: 1,
        value: { before: 0, after: 0, remove: 0, acceptance },
      }),
    ),
    JSON.stringify({ version: 2, value: { before: 0, after: 0, remove: 0 } }),
    JSON.stringify({ version: 1, value: { before: -1, after: 0, remove: 0 } }),
    JSON.stringify({
      version: 1,
      value: {
        before: 0,
        after: 0,
        remove: 0,
        retainedAfter: [
          { depth: 0, source: [{ space: "revision", id: 1 }], target: [] },
        ],
      },
    }),
  ]) {
    const escaped = escapeXmlAttribute(encoded);
    const node = element(
      `<r xmlns:history="${FOLIO_REVIEW_HISTORY_NAMESPACE}" history:resolutionJoins="${escaped}"/>`,
    );
    // Advisory metadata: the reader drops it, and says why.
    expect(parseResolutionJoins(node)).toBeUndefined();
    expect(reviewProvenanceIssue("resolutionJoins", encoded)).toBeDefined();
  }
  expect(
    parseResolutionJoins(
      element(
        '<r xmlns:history="urn:unrelated" history:resolutionJoins="invalid"/>',
      ),
    ),
  ).toBeUndefined();
});

test("paragraph cut provenance validates private envelopes and typed projection attributes", () => {
  for (const encoded of [
    "not-json",
    JSON.stringify({ version: 2, value: 0 }),
    JSON.stringify({ version: 1, value: -1 }),
    JSON.stringify({ version: 1, value: 0.5 }),
    JSON.stringify({ version: 1, value: MAX_REVISION_ID + 1 }),
    JSON.stringify({ version: 1, value: "undefined" }),
    JSON.stringify({ version: 1, value: 0, extra: true }),
  ]) {
    const node = element(
      `<r xmlns:history="${FOLIO_REVIEW_HISTORY_NAMESPACE}" history:resolutionJoin="${escapeXmlAttribute(encoded)}"/>`,
    );
    expect(parseParagraphMarkResolutionJoin(node)).toEqual({});
    expect(reviewProvenanceIssue("resolutionJoin", encoded)).toBeDefined();
  }
  assertExactModel(
    parseParagraphMarkResolutionJoin(
      element(
        '<r xmlns:history="urn:unrelated" history:resolutionJoin="invalid"/>',
      ),
    ),
    {},
  );
  for (const resolutionJoin of [
    -1,
    0.5,
    MAX_REVISION_ID + 1,
    Number.NaN,
    Number.POSITIVE_INFINITY,
  ]) {
    expect(() =>
      serializeParagraphMarkResolutionJoin({
        kind: "ins",
        info: { id: 3, author: "Reviewer" },
        resolutionJoin,
      }),
    ).toThrow(ReviewResolutionProvenanceError);
    const node = schema.node("paragraph", {
      pPrMark: {
        kind: "ins",
        info: { id: 3, author: "Reviewer" },
        resolutionJoin,
      },
    });
    const parsed = readParagraphAttrs(node);
    expect(parsed.ok).toBe(false);
  }
});

test("a cut record's identity handoff has its own private attribute", () => {
  const joins = {
    before: 0,
    after: 2,
    remove: 0,
    retainedEnclosing: [
      {
        levels: 1,
        source: [{ space: "revision", id: 7 }],
        target: [{ space: "revision", id: 2 }],
      },
    ],
  } as const;
  const written = serializeResolutionJoins(joins);
  // Its own attribute beside the joins, never a field inside them: a reader
  // that does not know it keeps the joins and drops only the handoff.
  expect(written).toContain("folio:retainedEnclosing=");
  const stored = /folio:resolutionJoins="([^"]*)"/u.exec(written)?.[1] ?? "";
  expect(stored).not.toContain("retainedEnclosing");
  const read = (attributes: string) =>
    parseResolutionJoins(
      element(
        `<r xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"${attributes}/>`,
      ),
    );
  expect(read(written)).toEqual(joins);
  // A split change carries it on its first piece, beside the piece's joins.
  const piece = serializeSplitResolutionPiece({ piece: 0, pieces: 2, joins });
  expect(piece).toContain("folio:retainedEnclosing=");
  expect(
    parseSplitResolutionPiece(
      element(`<r xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"${piece}/>`),
    ),
  ).toEqual({ piece: 0, pieces: 2, joins });
  // An unreadable handoff drops only itself, and says so.
  const { retainedEnclosing: _handoff, ...withoutHandoff } = joins;
  for (const encoded of [
    JSON.stringify({
      version: 1,
      value: [{ levels: 0, source: [], target: [] }],
    }),
    JSON.stringify({ version: 2, value: joins.retainedEnclosing }),
    "levels",
  ]) {
    const attributes = `${serializeResolutionJoins(withoutHandoff)} folio:retainedEnclosing="${escapeXmlAttribute(encoded)}"`;
    expect(read(attributes)).toEqual(withoutHandoff);
    expect(reviewProvenanceIssue("retainedEnclosing", encoded)).toBeDefined();
  }
  // Inside the joins attribute it is not folio's own writing: dropped.
  expect(
    read(
      ` folio:resolutionJoins="${escapeXmlAttribute(JSON.stringify({ version: 1, value: joins }))}"`,
    ),
  ).toBeUndefined();
});

test("the paragraph a rejected mark keeps has its own private attribute", () => {
  const mark = {
    kind: "ins",
    info: { id: 3, author: "Reviewer" },
    resolutionJoin: 0,
    survivor: "first",
  } as const;
  const written = serializeParagraphMarkResolutionJoin(mark);
  // Its own attribute beside the join depth, never a field inside it.
  expect(written).toContain("folio:rejectionSurvivor=");
  expect(written).toContain("folio:resolutionJoin=");
  const node = element(
    `<r xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"${written}/>`,
  );
  expect(parseParagraphMarkResolutionJoin(node)).toEqual({
    resolutionJoin: 0,
    survivor: "first",
  });
  // An unreadable value drops only the marker, and says so.
  for (const encoded of [
    JSON.stringify({ version: 1, value: "second" }),
    JSON.stringify({ version: 2, value: "first" }),
    "first",
  ]) {
    const dropped = element(
      `<r xmlns:history="${FOLIO_REVIEW_HISTORY_NAMESPACE}" history:rejectionSurvivor="${escapeXmlAttribute(encoded)}"/>`,
    );
    expect(parseParagraphMarkResolutionJoin(dropped)).toEqual({});
    expect(reviewProvenanceIssue("rejectionSurvivor", encoded)).toBeDefined();
  }
});

test("copy-origin ancestry uses the bounded versioned private envelope", () => {
  const encoded = serializeCopyOriginAncestors([0, 2, 4]);
  expect(encoded).toBe(JSON.stringify({ version: 1, value: [0, 2, 4] }));
  expect(
    parseCopyOriginAncestors(
      element(
        `<r xmlns:history="${FOLIO_REVIEW_HISTORY_NAMESPACE}" history:copyOriginAncestors="${escapeXmlAttribute(encoded ?? "")}"/>`,
      ),
    ),
  ).toEqual([0, 2, 4]);
  expect(
    parseCopyOriginAncestors(
      element(
        '<r xmlns:history="urn:unrelated" history:copyOriginAncestors="invalid"/>',
      ),
    ),
  ).toBeUndefined();

  const invalidEnvelopes = [
    "not-json",
    JSON.stringify({ version: 2, value: [0] }),
    JSON.stringify({ version: 1, value: [] }),
    JSON.stringify({ version: 1, value: [0, 0] }),
    JSON.stringify({ version: 1, value: [1, 0] }),
    JSON.stringify({ version: 1, value: [Number.MAX_SAFE_INTEGER + 1] }),
    JSON.stringify({
      version: 1,
      value: Array.from({ length: 257 }, (_, index) => index),
    }),
    " ".repeat(32_769),
  ];
  for (const invalid of invalidEnvelopes) {
    const node = element(
      `<r xmlns:history="${FOLIO_REVIEW_HISTORY_NAMESPACE}" history:copyOriginAncestors="${escapeXmlAttribute(invalid)}"/>`,
    );
    expect(parseCopyOriginAncestors(node)).toBeUndefined();
    expect(reviewProvenanceIssue("copyOriginAncestors", invalid)).toBeDefined();
  }
  expect(() => serializeCopyOriginAncestors([])).toThrow(
    ReviewResolutionProvenanceError,
  );
});

/** The paragraph read back from what the serializer wrote for it. */
const reopened = (xml: string): Paragraph => {
  const root = element(
    `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}">${xml}</root>`,
  );
  const paragraph = findChild(root, "w", "p");
  if (!paragraph) throw new TypeError("The paragraph disappeared.");
  return parseParagraph(paragraph, null, null, null);
};

test("a change holding only a link keeps its joins through a save", () => {
  const paragraph = {
    type: "paragraph",
    content: [
      {
        type: "insertion",
        info: { id: 1, author: "Reviewer" },
        resolutionJoins: { before: 1, after: 1, remove: 1 },
        content: [
          {
            type: "hyperlink",
            href: "https://example.com",
            children: [
              { type: "run", content: [{ type: "text", text: "link" }] },
            ],
          },
        ],
      },
    ],
  } satisfies Paragraph;
  const xml = serializeParagraph(paragraph);
  const change = reopened(xml).content[0];
  expect(change?.type === "insertion" && change.resolutionJoins).toEqual({
    before: 1,
    after: 1,
    remove: 1,
  });
  expect(serializeParagraph(reopened(xml))).toBe(xml);
});

test("nested hyperlink lifting splits enclosing resolution provenance and reads it back", () => {
  const paragraph = {
    type: "paragraph",
    content: [
      {
        type: "insertion",
        info: { id: 35, author: "Outer Reviewer" },
        resolutionJoins: { before: 1, after: 1, remove: 1 },
        content: [
          { type: "run", content: [{ type: "text", text: "before" }] },
          {
            type: "deletion",
            info: { id: 37, author: "Inner Reviewer" },
            content: [
              {
                type: "hyperlink",
                anchor: "target",
                children: [
                  { type: "run", content: [{ type: "text", text: "linked" }] },
                ],
              },
            ],
          },
          { type: "run", content: [{ type: "text", text: "after" }] },
        ],
      },
    ],
  } satisfies Paragraph;
  const xml = serializeParagraph(paragraph);
  expectNoHyperlinkInsideRevision(xml);
  const change = reopened(xml).content[0];
  expect(
    change?.type === "insertion" && {
      joins: change.resolutionJoins,
      kinds: change.content.map(({ type }) => type),
    },
  ).toEqual({
    joins: { before: 1, after: 1, remove: 1 },
    kinds: ["run", "deletion", "run"],
  });
  expect(serializeParagraph(reopened(xml))).toBe(xml);
});

type RecursiveInlineContainerKind = Extract<
  ParagraphContent,
  { content: ParagraphContent[] }
>["type"];
const RECURSIVE_CONTAINER_KINDS = {
  insertion: fc.constant("insertion" as const),
  deletion: fc.constant("deletion" as const),
  moveFrom: fc.constant("moveFrom" as const),
  moveTo: fc.constant("moveTo" as const),
  simpleField: fc.constant("simpleField" as const),
  inlineSdt: fc.constant("inlineSdt" as const),
  inlineWrapper: fc.constantFrom(
    "bidiEmbedding",
    "bidiOverride",
    "smartTag",
    "customXml",
  ),
} satisfies Record<RecursiveInlineContainerKind, fc.Arbitrary<string>>;

test(
  "generated nested lifting splits provenance into pieces and reads it back",
  () => {
    assertProperty(
      fc.property(
        fc.array(fc.oneof(...Object.values(RECURSIVE_CONTAINER_KINDS)), {
          minLength: 1,
          maxLength: 5,
        }),
        fc.constantFrom("plain", "field"),
        fc.constantFrom("linked", "unlinked"),
        fc.constantFrom("none", "outer", "inner", "every"),
        fc.record({
          before: fc.integer({ min: 0, max: 4 }),
          after: fc.integer({ min: 0, max: 4 }),
          remove: fc.integer({ min: 0, max: 4 }),
        }),
        (kinds, carrier, link, provenance, joins) => {
          const run = {
            type: "run",
            content: [{ type: "text", text: "result" }],
          } as const;
          const leaf =
            link === "linked"
              ? {
                  type: "hyperlink" as const,
                  anchor: "target",
                  children: [{ ...run, content: [...run.content] }],
                }
              : { ...run, content: [...run.content] };
          let content: TrackedRunChange["content"] =
            carrier === "field"
              ? [
                  {
                    type: "simpleField",
                    instruction: "REF target",
                    fieldType: "REF",
                    content: [leaf],
                  },
                ]
              : [leaf];
          let owners = 0;
          let exposedLink = link === "linked";
          // Owners whose joins travel as split pieces around the lifted link.
          let splitOwners = 0;
          for (const [index, kind] of kinds.entries()) {
            const retains =
              provenance === "every" ||
              (provenance === "inner" && index === 0) ||
              (provenance === "outer" && index === kinds.length - 1);
            const nestedContent: TrackedRunChange["content"] = [
              { type: "run", content: [{ type: "text", text: "before" }] },
            ];
            nestedContent.push(...content);
            nestedContent.push({
              type: "run",
              content: [{ type: "text", text: "after" }],
            });
            switch (kind) {
              case "insertion":
              case "deletion":
              case "moveFrom":
              case "moveTo":
                if (retains) {
                  owners++;
                  if (exposedLink) splitOwners++;
                }
                content = [
                  {
                    type: kind,
                    info: { id: index + 1, author: "Reviewer" },
                    ...(retains ? { resolutionJoins: joins } : {}),
                    content: nestedContent,
                  },
                ];
                break;
              case "simpleField":
                if (retains) {
                  owners++;
                  if (exposedLink) splitOwners++;
                }
                content = [
                  {
                    type: "simpleField",
                    instruction: "REF target",
                    fieldType: "REF",
                    content: [
                      {
                        type: "insertion",
                        info: { id: index + 1, author: "Reviewer" },
                        ...(retains ? { resolutionJoins: joins } : {}),
                        content: nestedContent,
                      },
                    ],
                  },
                ];
                break;
              case "inlineSdt":
                exposedLink = false;
                content = [
                  {
                    type: "inlineSdt",
                    properties: { sdtType: "richText" },
                    content: nestedContent,
                  },
                ];
                break;
              case "bidiEmbedding":
              case "bidiOverride":
                content = [
                  {
                    type: "inlineWrapper",
                    kind: "bidi",
                    control:
                      kind === "bidiEmbedding" ? "embedding" : "override",
                    direction: "rtl",
                    content: nestedContent,
                  },
                ];
                break;
              case "smartTag":
              case "customXml":
                content = [
                  {
                    type: "inlineWrapper",
                    kind,
                    element: "tag",
                    uri: "urn:tag",
                    content: nestedContent,
                  },
                ];
                break;
              default: {
                const untested: never = kind;
                return untested;
              }
            }
          }
          // Every generated container stack is inside a tracked owner, including
          // stacks made entirely from fields, transparent wrappers and controls.
          content = [
            {
              type: "insertion",
              info: { id: 100, author: "Root Reviewer" },
              content,
            },
          ];
          const paragraph: Paragraph = { type: "paragraph", content };
          const xml = serializeParagraph(paragraph);
          expectNoHyperlinkInsideRevision(xml);
          expect(xml.match(/folio:resolutionJoins=/gu)?.length ?? 0).toBe(
            owners - splitOwners,
          );
          // Each split owner's first piece carries its joins.
          expect(
            xml.match(
              /folio:splitResolutionJoins="[^"]*&quot;piece&quot;:0,&quot;pieces&quot;:\d+,&quot;joins&quot;/gu,
            )?.length ?? 0,
          ).toBe(splitOwners);
          if (owners > splitOwners)
            expect(xml).toContain(serializeResolutionJoins(joins));
          const parsedRoot = element(
            `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}">${xml}</root>`,
          );
          const parsedParagraph = findChild(parsedRoot, "w", "p");
          if (!parsedParagraph)
            throw new TypeError("Generated paragraph disappeared.");
          const reopenedXml = serializeParagraph(
            parseParagraph(parsedParagraph, null, null, null),
          );
          expectNoHyperlinkInsideRevision(reopenedXml);
          // Reading the pieces back rebuilds each change with its joins, so
          // writing it again writes the same pieces. The one shape the reader
          // keeps as markup instead is a link whose lift stops at a content
          // control with a transparent wrapper outermost below it and a change
          // inside that wrapper: a revision inside a wrapper inside a link.
          const lift = kinds.indexOf("inlineSdt");
          const lifted = lift < 0 ? [] : kinds.slice(0, lift);
          const wrapperKinds: ReadonlySet<string> = new Set([
            "bidiEmbedding",
            "bidiOverride",
            "smartTag",
            "customXml",
          ]);
          const capturedByReader =
            link === "linked" &&
            lifted.length > 0 &&
            wrapperKinds.has(lifted.at(-1) ?? "") &&
            lifted.some((kind) => !wrapperKinds.has(kind));
          if (!capturedByReader) expect(reopenedXml).toBe(xml);
          expect(reopenedXml.match(/<w:hyperlink\b/gu)?.length ?? 0).toBe(
            link === "linked" ? 1 : 0,
          );
        },
      ),
      {
        numRuns: 80,
        id: "generated nested lifting splits provenance into pieces and reads it back",
      },
    );
  },
  propertyTestTimeout(10_000),
);

const expectNoHyperlinkInsideRevision = (xml: string) => {
  const root = element(
    `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</root>`,
  );
  const pending = [{ node: root, revisionDepth: 0 }];
  while (pending.length > 0) {
    const next = pending.pop();
    if (!next) continue;
    const name = getLocalName(next.node.name);
    if (name === "hyperlink") expect(next.revisionDepth).toBe(0);
    const revisionDepth =
      name === "sdt"
        ? 0
        : next.revisionDepth +
          (["ins", "del", "moveFrom", "moveTo"].includes(name) ? 1 : 0);
    for (const child of getChildElements(next.node))
      pending.push({ node: child, revisionDepth });
  }
};

for (const provenance of ["none", "joins"] as const) {
  test(`tracked field wrapper hyperlink uses lifted revision placement (${provenance})`, () => {
    const paragraph = {
      type: "paragraph",
      content: [
        {
          type: "insertion",
          info: { id: 35, author: "Reviewer" },
          ...(provenance === "joins"
            ? { resolutionJoins: { before: 1, after: 1, remove: 1 } }
            : {}),
          content: [
            {
              type: "simpleField",
              instruction: "REF target",
              fieldType: "REF",
              content: [
                {
                  type: "inlineWrapper",
                  kind: "bidi",
                  control: "embedding",
                  direction: "rtl",
                  content: [
                    {
                      type: "hyperlink",
                      anchor: "target",
                      children: [
                        {
                          type: "run",
                          content: [{ type: "text", text: "linked" }],
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    } satisfies Paragraph;
    const xml = serializeParagraph(paragraph);
    if (provenance === "joins") {
      const change = reopened(xml).content[0];
      expect(change?.type === "insertion" && change.resolutionJoins).toEqual({
        before: 1,
        after: 1,
        remove: 1,
      });
    }
    expect(serializeParagraph(reopened(xml))).toBe(xml);
    expectNoHyperlinkInsideRevision(xml);
    expect(xml).toContain("<w:hyperlink");
    expect(xml).toContain('<w:dir w:val="rtl">');
    expect(xml).toContain("linked");
  });
}

test("opaque tracked captures refuse untyped hyperlinks", () => {
  expect(() =>
    serializeParagraph({
      type: "paragraph",
      content: [
        {
          type: "insertion",
          info: { id: 35, author: "Reviewer" },
          content: [
            {
              type: "preservedInline",
              text: "linked",
              xml: '<w:dir w:val="rtl"><w:hyperlink w:anchor="target"><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:dir>',
            },
          ],
        },
      ],
    }),
  ).toThrow(/tracked capture/);
});

test("hyperlink revision hoisting descends through transparent revision containers", () => {
  const paragraph = parseParagraph(
    element(
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:hyperlink w:anchor="target"><w:ins w:id="1" w:author="Outer"><w:dir w:val="rtl"><w:del w:id="2" w:author="Inner"><w:r><w:delText>linked</w:delText></w:r></w:del></w:dir></w:ins></w:hyperlink></w:p>',
    ),
    null,
    null,
    null,
  );
  const xml = serializeParagraph(paragraph);
  expect(xml.match(/<w:hyperlink\b/gu)).toHaveLength(1);
  expect(xml).toContain('w:anchor="target"');
  expect(xml).toContain('w:id="1"');
  expect(xml).toContain('w:id="2"');
  expect(JSON.stringify(paragraph)).not.toContain('"type":"preservedInline"');
  expectNoHyperlinkInsideRevision(xml);
});
