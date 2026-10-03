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
import { findChild, parseXmlDocument, WORDPROCESSINGML_NAMESPACE_URIS } from "./xmlParser";
import type { Paragraph, ParagraphMarkChange } from "../types/document";

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
