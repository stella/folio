import { expect, test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import { escapeXmlAttribute } from "@stll/docx-core";
import { MAX_REVISION_ID, PARAGRAPH_MARK_CHANGE_KINDS } from "@stll/docx-core/model";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { readParagraphAttrs } from "../prosemirror/attrs";
import { schema } from "../prosemirror/schema";
import { assertExactModel } from "../../../../test/exactModel";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializePartElement } from "./serializer/partNamespaces";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import {
  parseResolutionJoins,
  parseParagraphMarkResolutionJoin,
  serializeParagraphMarkResolutionJoin,
  ReviewResolutionProvenanceError,
  serializeBoundaryJoins,
  serializeResolutionJoins,
} from "./reviewResolutionProvenance";
import {
  getChildElements,
  getLocalName,
  findChild,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "./xmlParser";
import type {
  TrackedRunChange,
  ParagraphContent,
  Paragraph,
  ParagraphMarkChange,
} from "../types/document";

setDefaultTimeout(propertyTestTimeout(30_000));

const element = (xml: string) => {
  const parsed = parseXmlDocument(xml);
  if (parsed === null) throw new TypeError("The provenance fixture is not XML.");
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
            { minLength: 1, maxLength: 3, selector: (group) => JSON.stringify(group) },
          ),
          { nil: undefined },
        ),
        fc.option(fc.constant("merge-plain-runs"), { nil: undefined }),
        (prefix, depths, boundaryJoins, kind, resolutionJoin, deferredRemove, acceptance) => {
          const paragraphMark = {
            kind,
            info: { id: 3, author: "Reviewer" },
            ...(resolutionJoin === "absent"
              ? {}
              : { resolutionJoin: resolutionJoin === "undefined" ? undefined : resolutionJoin }),
          } satisfies ParagraphMarkChange;
          const paragraphMarkAttrs = serializeParagraphMarkResolutionJoin(paragraphMark).replace(
            "folio:",
            `${prefix}:`,
          );
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
          const attrs = serializeResolutionJoins(joins).replace("folio:", `${prefix}:`);
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
          const pmDoc = toProseDoc({ package: { document: { content: [paragraph] } } });
          const projected = proseDocToBlocks(pmDoc, [paragraph]).at(0);
          if (projected?.type !== "paragraph")
            throw new TypeError("Projected paragraph disappeared.");
          assertExactModel(projected.pPrMark, paragraphMark);
          const wrapper = paragraph.content.at(0);
          if (wrapper?.type !== "insertion") throw new TypeError("A tracked wrapper disappeared.");
          assertExactModel(wrapper.resolutionJoins, joins);
          const run = wrapper.content.at(0);
          if (run?.type !== "run") throw new TypeError("A reviewed run disappeared.");
          assertExactModel(run.propertyChanges?.at(0)?.boundaryJoins, boundaryJoins);
          const saved = serializePartElement({
            partPath: "word/document.xml",
            rootName: "w:document",
            baselinePrefixes: ["w"],
            sourceBindings: undefined,
            body: serializeParagraph(paragraph),
          });
          expect(saved).toContain('xmlns:folio="urn:stella:folio:review-history:1"');
          expect(saved).toMatch(/mc:Ignorable="[^"]*folio/u);
          const reopenedNode = findChild(element(saved), "w", "p");
          if (reopenedNode === null) throw new TypeError("Saved paragraph disappeared.");
          const reopened = parseParagraph(reopenedNode, null, null, null, null, null);
          assertExactModel(reopened.pPrMark, paragraphMark);
          const reopenedWrapper = reopened.content.at(0);
          if (reopenedWrapper?.type !== "insertion")
            throw new TypeError("Saved wrapper disappeared.");
          assertExactModel(reopenedWrapper.resolutionJoins, joins);
          const reopenedRun = reopenedWrapper.content.at(0);
          if (reopenedRun?.type !== "run") throw new TypeError("Saved run disappeared.");
          assertExactModel(reopenedRun.propertyChanges?.at(0)?.boundaryJoins, boundaryJoins);
        },
      ),
      { numRuns: 30 },
    );
  });
}

test("malformed and future provenance is refused instead of discarded", () => {
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
      JSON.stringify({ version: 1, value: { before: 0, after: 0, remove: 0, deferredRemove } }),
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
      JSON.stringify({ version: 1, value: { before: 0, after: 0, remove: 0, deferredRemove } }),
    ),
    ...["unknown", null, 1, {}, ["merge-plain-runs"]].map((acceptance) =>
      JSON.stringify({ version: 1, value: { before: 0, after: 0, remove: 0, acceptance } }),
    ),
    JSON.stringify({ version: 2, value: { before: 0, after: 0, remove: 0 } }),
    JSON.stringify({ version: 1, value: { before: -1, after: 0, remove: 0 } }),
    JSON.stringify({
      version: 1,
      value: {
        before: 0,
        after: 0,
        remove: 0,
        retainedAfter: [{ depth: 0, source: [{ space: "revision", id: 1 }], target: [] }],
      },
    }),
  ]) {
    const escaped = escapeXmlAttribute(encoded);
    const node = element(
      `<r xmlns:history="${FOLIO_REVIEW_HISTORY_NAMESPACE}" history:resolutionJoins="${escaped}"/>`,
    );
    expect(() => parseResolutionJoins(node)).toThrow(ReviewResolutionProvenanceError);
  }
  expect(
    parseResolutionJoins(
      element('<r xmlns:history="urn:unrelated" history:resolutionJoins="invalid"/>'),
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
    expect(() => parseParagraphMarkResolutionJoin(node)).toThrow(ReviewResolutionProvenanceError);
  }
  assertExactModel(
    parseParagraphMarkResolutionJoin(
      element('<r xmlns:history="urn:unrelated" history:resolutionJoin="invalid"/>'),
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
      pPrMark: { kind: "ins", info: { id: 3, author: "Reviewer" }, resolutionJoin },
    });
    const parsed = readParagraphAttrs(node);
    expect(parsed.ok).toBe(false);
  }
});

test("save refuses provenance requiring unsupported hyperlink repartition", () => {
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
            children: [{ type: "run", content: [{ type: "text", text: "link" }] }],
          },
        ],
      },
    ],
  } satisfies Paragraph;
  expect(() => serializeParagraph(paragraph)).toThrow(ReviewResolutionProvenanceError);
});

test("nested hyperlink lifting refuses enclosing resolution provenance", () => {
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
                children: [{ type: "run", content: [{ type: "text", text: "linked" }] }],
              },
            ],
          },
          { type: "run", content: [{ type: "text", text: "after" }] },
        ],
      },
    ],
  } satisfies Paragraph;
  expect(() => serializeParagraph(paragraph)).toThrow(ReviewResolutionProvenanceError);
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
  inlineWrapper: fc.constantFrom("bidiEmbedding", "bidiOverride", "smartTag", "customXml"),
} satisfies Record<RecursiveInlineContainerKind, fc.Arbitrary<string>>;

test(
  "generated nested lifting refuses provenance and unsplit revisions retain it",
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
          const run = { type: "run", content: [{ type: "text", text: "result" }] } as const;
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
          for (const [index, kind] of kinds.entries()) {
            const retains =
              provenance === "every" ||
              (provenance === "inner" && index === 0) ||
              (provenance === "outer" && index === kinds.length - 1);
            const nestedContent: TrackedRunChange["content"] = [
              { type: "run", content: [{ type: "text", text: "before" }] },
            ];
            nestedContent.push(...content);
            nestedContent.push({ type: "run", content: [{ type: "text", text: "after" }] });
            switch (kind) {
              case "insertion":
              case "deletion":
              case "moveFrom":
              case "moveTo":
                if (retains) owners++;
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
                if (retains) owners++;
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
                    control: kind === "bidiEmbedding" ? "embedding" : "override",
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
          content = [{ type: "insertion", info: { id: 100, author: "Root Reviewer" }, content }];
          const paragraph: Paragraph = { type: "paragraph", content };
          if (link === "linked" && owners > 0) {
            expect(() => serializeParagraph(paragraph)).toThrow(ReviewResolutionProvenanceError);
            return;
          }
          if (link === "linked" && kinds.includes("inlineSdt")) {
            expect(() => serializeParagraph(paragraph)).toThrow(/content control/);
            return;
          }
          const xml = serializeParagraph(paragraph);
          expectNoHyperlinkInsideRevision(xml);
          expect(xml.match(/folio:resolutionJoins=/gu)?.length ?? 0).toBe(owners);
          if (owners > 0) expect(xml).toContain(serializeResolutionJoins(joins));
          const parsedRoot = element(
            `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}">${xml}</root>`,
          );
          const parsedParagraph = findChild(parsedRoot, "w", "p");
          if (!parsedParagraph) throw new TypeError("Generated paragraph disappeared.");
          const reopenedXml = serializeParagraph(parseParagraph(parsedParagraph, null, null, null));
          expectNoHyperlinkInsideRevision(reopenedXml);
          expect(reopenedXml.match(/<w:hyperlink\b/gu)?.length ?? 0).toBe(
            link === "linked" ? 1 : 0,
          );
        },
      ),
      { numRuns: 80 },
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
      next.revisionDepth + (["ins", "del", "moveFrom", "moveTo"].includes(name) ? 1 : 0);
    for (const child of getChildElements(next.node)) pending.push({ node: child, revisionDepth });
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
                      children: [{ type: "run", content: [{ type: "text", text: "linked" }] }],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    } satisfies Paragraph;
    if (provenance === "joins") {
      expect(() => serializeParagraph(paragraph)).toThrow(ReviewResolutionProvenanceError);
      return;
    }
    const xml = serializeParagraph(paragraph);
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
              xml: '<w:sdt><w:sdtPr/><w:sdtContent><w:hyperlink w:anchor="target"><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:sdtContent></w:sdt>',
            },
          ],
        },
      ],
    }),
  ).toThrow(/tracked capture/);
});
