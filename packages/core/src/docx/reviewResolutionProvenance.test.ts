import { expect, test } from "bun:test";
import fc from "fast-check";
import { escapeXmlAttribute } from "@stll/docx-core";
import { assertExactModel } from "../../../../test/exactModel";
import { assertProperty } from "../../../../test/property-testing";
import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializePartElement } from "./serializer/partNamespaces";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import {
  parseResolutionJoins,
  ReviewResolutionProvenanceError,
  serializeBoundaryJoins,
  serializeResolutionJoins,
} from "./reviewResolutionProvenance";
import { findChild, parseXmlDocument, WORDPROCESSINGML_NAMESPACE_URIS } from "./xmlParser";
import type { Paragraph } from "../types/document";

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
        (prefix, depths, boundaryJoins) => {
          const joins = {
            ...depths,
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
              `<q:p xmlns:q="${wordNamespace}" xmlns:${prefix}="${FOLIO_REVIEW_HISTORY_NAMESPACE}"><q:ins q:id="1" q:author="Reviewer"${attrs}><q:r><q:rPr><q:b/><q:rPrChange q:id="2" q:author="Reviewer"${boundaryAttrs}><q:rPr/></q:rPrChange></q:rPr><q:t>text</q:t></q:r></q:ins></q:p>`,
            ),
            null,
            null,
            null,
            null,
            null,
          );
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
