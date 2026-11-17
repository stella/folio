import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import type {
  BlockContent,
  Document,
  ParagraphContent,
} from "@stll/docx-core/model";
import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
} from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { storyParagraphs } from "../../../docx-core/src/ops/blocks";
import { normalizeForOps } from "../../../docx-core/src/ops/contract";
import { leafSpans } from "../../../docx-core/src/ops/leaves";
import {
  paragraphHoldsContent,
  paragraphHoldsContentOnceAccepted,
  paragraphLength,
} from "../../../docx-core/src/ops/offsets";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { paragraphContentOf } from "../prosemirror/zeroWidthAnchors";
import { FolioDocxReviewer } from "./headless";

const XML = '<w:proofErr w:type="spellStart"/>';
const MARKS = ["ins", "del", "moveFrom", "moveTo"] as const;
const ANCESTRY = [
  "standing",
  "hyperlink",
  "inlineWrapper",
  "inlineSdt",
  "insertion",
  "deletion",
  "moveFrom",
  "moveTo",
] as const;
const FOLLOWERS = ["paragraph", "table", "sdt", "opaque"] as const;
type SourceOptions = {
  level: "inline" | "run";
  display: "empty" | "visible";
  mark: (typeof MARKS)[number];
  ancestry: (typeof ANCESTRY)[number];
  follower: (typeof FOLLOWERS)[number];
};
const source = ({
  level,
  display,
  mark,
  ancestry,
  follower,
}: SourceOptions) => {
  const text = display === "empty" ? "" : "Opaque";
  const xml =
    display === "empty"
      ? XML
      : "<w:customXml><w:r><w:t>Opaque</w:t></w:r></w:customXml>";
  const payload =
    level === "inline"
      ? { type: "preservedInline" as const, xml, text }
      : {
          type: "run" as const,
          content: [{ type: "preservedXml" as const, xml, text }],
        };
  let content: ParagraphContent[] = [payload];
  switch (ancestry) {
    case "hyperlink":
      content = [
        { type: "hyperlink", href: "https://example.org", children: [payload] },
      ];
      break;
    case "inlineWrapper":
      content = [
        {
          type: "inlineWrapper",
          kind: "bidi",
          control: "embedding",
          content: [payload],
        },
      ];
      break;
    case "inlineSdt":
      content = [
        {
          type: "inlineSdt",
          properties: { sdtType: "richText", id: 12 },
          content: [payload],
        },
      ];
      break;
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo":
      content = [
        {
          type: ancestry,
          info: { id: 2, author: "Content" },
          content: [payload],
        },
      ];
      break;
    case "standing":
      break;
  }
  const following = {
    type: "paragraph",
    paraId: "00000002",
    content: [],
  } satisfies BlockContent;
  let block: BlockContent;
  switch (follower) {
    case "paragraph":
      block = following;
      break;
    case "table":
      block = {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [{ type: "tableCell", content: [following] }],
          },
        ],
      };
      break;
    case "sdt":
      block = {
        type: "blockSdt",
        properties: { sdtType: "group", id: 9 },
        content: [following],
      };
      break;
    case "opaque":
      block = { type: "preservedBlock", xml: '<w:altChunk r:id="rIdOpaque"/>' };
      break;
    default: {
      const unreachable: never = follower;
      return unreachable;
    }
  }
  return normalizeForOps({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content,
            pPrMark: { kind: mark, info: { id: 1, author: "Mark" } },
          },
          block,
          {
            type: "paragraph",
            paraId: "00000003",
            content: [
              { type: "run", content: [{ type: "text", text: "Tail" }] },
            ],
          },
        ],
      },
    },
  });
};
const payloads = (document: Document) =>
  storyParagraphs(document.package.document).flatMap(({ paragraph }) =>
    leafSpans(paragraph.content).flatMap(({ node }) =>
      node.type === "preservedInline" || node.type === "preservedXml"
        ? [node]
        : [],
    ),
  );
const removesContent = (
  ancestry: SourceOptions["ancestry"],
  decision: "accept" | "reject",
) =>
  decision === "accept"
    ? ancestry === "deletion" || ancestry === "moveFrom"
    : ancestry === "insertion" || ancestry === "moveTo";

test(
  "opaque payload survives paragraph-mark resolution independently of logical width",
  () => {
    assertProperty(
      fc.property(
        fc.constantFrom(...MARKS),
        fc.constantFrom(...ANCESTRY),
        fc.constantFrom(...FOLLOWERS),
        fc.constantFrom("empty" as const, "visible" as const),
        (mark, ancestry, follower, display) => {
          for (const level of ["inline", "run"] as const) {
            const document = source({
              level,
              display,
              mark,
              ancestry,
              follower,
            });
            const paragraph = document.package.document.content.at(0);
            if (paragraph?.type !== "paragraph")
              return panic("The source must start with its marked paragraph.");
            expect(paragraphLength(paragraph)).toBe(
              level === "inline" && display === "empty" ? 0 : 1,
            );
            expect(paragraphHoldsContent(paragraph)).toBe(true);
            expect(paragraphHoldsContentOnceAccepted(paragraph)).toBe(
              ancestry !== "deletion" && ancestry !== "moveFrom",
            );
            expect(paragraphContentOf(toProseDoc(document).child(0))).toBe(
              true,
            );
            for (const decision of ["accept", "reject"] as const)
              for (const revisionIds of [[1], [1, 2]]) {
                const edited = applyDocumentOps(document, [
                  {
                    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                    story: OP_STORIES.MAIN,
                    revisionIds,
                    decision,
                  },
                ]).unwrap();
                const expected =
                  revisionIds.length === 2 && removesContent(ancestry, decision)
                    ? 0
                    : 1;
                expect(payloads(edited.document)).toEqual(
                  expected === 0 ? [] : payloads(document),
                );
                expect(
                  applyDocumentOps(
                    edited.document,
                    JSON.parse(JSON.stringify(edited.inverse)),
                  ).unwrap().document,
                ).toEqual(document);
                const undo = applyDocumentOps(
                  edited.document,
                  edited.inverse,
                ).unwrap();
                expect(
                  applyDocumentOps(undo.document, undo.inverse).unwrap()
                    .document,
                ).toEqual(edited.document);
              }
          }
        },
      ),
      { id: "opaque paragraph retention", numRuns: 80 },
    );
  },
  propertyTestTimeout(15_000),
);

for (const ancestry of ANCESTRY)
  for (const mark of MARKS)
    test(`${ancestry}/${mark} opaque source agrees in both reviewers and saved resolutions`, async () => {
      const document = source({
        level: "inline",
        display: "empty",
        mark,
        ancestry,
        follower: "sdt",
      });
      const bytes = (
        await prepareCanonicalDocxInput(await createDocx(document))
      )
        .unwrap()
        .slice().buffer;
      for (const decision of ["accept", "reject"] as const) {
        const legacy = await FolioDocxReviewer.fromBuffer(bytes.slice(0), {
          session: "legacy",
        });
        const canonical = await FolioDocxReviewer.fromBuffer(bytes.slice(0), {
          session: "canonical",
        });
        const legacyOutcome =
          decision === "accept" ? legacy.acceptAll() : legacy.rejectAll();
        const canonicalOutcome =
          decision === "accept" ? canonical.acceptAll() : canonical.rejectAll();
        expect(legacyOutcome.skipped).toEqual([]);
        expect(canonicalOutcome.skipped).toEqual([]);
        expect(legacyOutcome.resolved).toBeGreaterThan(0);
        expect(canonicalOutcome.resolved).toBeGreaterThan(0);
        expect(canonical.getContent()).toEqual(legacy.getContent());
        const expected = removesContent(ancestry, decision) ? 0 : 1;
        for (const reviewer of [legacy, canonical]) {
          expect(payloads(reviewer.toDocument())).toEqual(
            expected === 0 ? [] : payloads(document),
          );
          const saved = await parseDocx(await reviewer.toBuffer(), {
            detectVariables: false,
            preloadFonts: false,
          });
          expect(payloads(saved)).toEqual(
            expected === 0 ? [] : payloads(document),
          );
        }
      }
    });
