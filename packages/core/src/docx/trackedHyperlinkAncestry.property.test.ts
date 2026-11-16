import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  applyDocumentOp,
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type {
  Document,
  Hyperlink,
  InlineWrapper,
  Paragraph,
  ParagraphContent,
  Run,
  TrackedRunChange,
} from "../types/document";
import { parseParagraph } from "./paragraphParser";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument } from "./xmlParser";

const wrapperKinds = [
  "embedding",
  "override",
  "smartTag",
  "customXml",
] as const;
const revisionKinds = ["insertion", "deletion", "moveFrom", "moveTo"] as const;
const joins = { before: 0, after: 0, remove: 0 };
const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});

const wrapper = (
  kind: (typeof wrapperKinds)[number],
  content: ParagraphContent[],
): InlineWrapper => {
  switch (kind) {
    case "embedding":
    case "override":
      return {
        type: "inlineWrapper",
        kind: "bidi",
        control: kind,
        direction: "rtl",
        content,
      };
    case "smartTag":
    case "customXml":
      return {
        type: "inlineWrapper",
        kind,
        element: "scope",
        uri: "urn:scope",
        content,
      };
  }
};

type RevisionOptions = Pick<TrackedRunChange, "type" | "content"> & {
  id: number;
};
const revision = ({
  type,
  content,
  id,
}: RevisionOptions): TrackedRunChange => ({
  type,
  info: { id, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
  resolutionJoins: joins,
  content,
});

const reopen = (paragraph: Paragraph): Paragraph => {
  const xml = serializeParagraph(paragraph).replace(
    "<w:p",
    `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"`,
  );
  const parsed = parseXmlDocument(xml);
  if (parsed === null) throw new TypeError("The saved paragraph is not XML.");
  return parseParagraph(parsed, new Map(), null, null);
};

test(
  "saving a lifted hyperlink preserves its revision and wrapper ancestry through resolution",
  () => {
    assertProperty(
      fc.property(
        fc.constantFrom(...revisionKinds),
        fc.array(
          fc.record({
            kind: fc.constantFrom(...wrapperKinds),
            revision: fc.constantFrom(...revisionKinds),
          }),
          { minLength: 1, maxLength: 4 },
        ),
        fc.constantFrom("only", "before", "after", "both", "links"),
        fc.boolean(),
        fc.array(fc.constantFrom(...wrapperKinds), { maxLength: 3 }),
        (type, layers, layout, nested, innerLayers) => {
          let linkedContent: Hyperlink["children"] = [run("linked")];
          for (const kind of innerLayers.toReversed())
            linkedContent = [wrapper(kind, linkedContent)];
          const link = {
            type: "hyperlink",
            anchor: "Target",
            href: "#Target",
            children: linkedContent,
          } as const satisfies ParagraphContent;
          let content: TrackedRunChange["content"] = [link];
          if (layout === "before" || layout === "both")
            content.unshift(run("before"));
          if (layout === "after" || layout === "both")
            content.push(run("after"));
          if (layout === "links")
            content.push({
              type: "hyperlink",
              anchor: "Other",
              href: "#Other",
              children: [run("other")],
            });
          const revisionIds = [7];
          for (const [index, layer] of layers.toReversed().entries()) {
            const { kind } = layer;
            content = [wrapper(kind, content)];
            if (nested) {
              const id = 8 + index;
              revisionIds.push(id);
              content = [revision({ type: layer.revision, content, id })];
            }
          }
          const source = {
            type: "paragraph",
            paraId: "10000000",
            content: [revision({ type, content, id: 7 })],
          } as const satisfies Paragraph;
          const saved = reopen(source);
          expect(saved.content).toEqual(source.content);
          expect(reopen(saved).content).toEqual(source.content);
          const document = (paragraph: Paragraph): Document => ({
            package: { document: { content: [paragraph] } },
          });
          for (const decision of [
            REVISION_DECISIONS.ACCEPT,
            REVISION_DECISIONS.REJECT,
          ]) {
            const operation = {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds,
              decision,
            } as const;
            const resolved = applyDocumentOp(
              document(saved),
              operation,
            ).unwrap();
            expect(resolved.document.package.document.content).toEqual(
              applyDocumentOp(document(source), operation).unwrap().document
                .package.document.content,
            );
            expect(
              applyDocumentOps(resolved.document, resolved.inverse).unwrap()
                .document,
            ).toEqual(document(saved));
            expect(
              applyDocumentOps(
                resolved.document,
                JSON.parse(JSON.stringify(resolved.inverse)),
              ).unwrap().document,
            ).toEqual(document(saved));
          }
        },
      ),
      {
        numRuns: 100,
        id: "saving a lifted hyperlink preserves its revision and wrapper ancestry through resolution",
        examples: [
          [
            "deletion",
            [{ kind: "embedding", revision: "insertion" }],
            "only",
            false,
            [],
          ],
          [
            "insertion",
            [
              { kind: "embedding", revision: "deletion" },
              { kind: "embedding", revision: "insertion" },
            ],
            "only",
            true,
            ["customXml"],
          ],
        ],
      },
    );
  },
  propertyTestTimeout(30_000),
);
