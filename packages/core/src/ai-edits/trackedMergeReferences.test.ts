import { expect, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FOLIO_PARAGRAPH_ALIGNMENT_VALUES,
  type FolioDocumentOperation,
} from "../document-operations";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { paragraphPropertiesSnapshot } from "../prosemirror/commands/propertyChangeScope";
import { resolveStateStory } from "../prosemirror/markupViewProjection";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";

const open = async (markdown: string) => {
  const bytes = (await ensureParaIds(await createDocx(fromMarkdown(markdown)))).docx;
  return FolioDocxReviewer.fromBuffer(bytes, { author: "Test" });
};
const reopen = async (reviewer: FolioDocxReviewer) =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), { author: "Test" });

const paragraphProjection = (doc: PMNode) => {
  const rows: { text: string; properties: ReturnType<typeof paragraphPropertiesSnapshot> }[] = [];
  doc.descendants((node) => {
    if (node.type.name !== "paragraph") return;
    rows.push({ text: node.textContent, properties: paragraphPropertiesSnapshot(node) });
  });
  return rows;
};

const refusedProperty = {
  reason: "pendingParagraphMarkDeletion",
  canonicalRefusal: { gap: CANONICAL_GAP.publicPendingParagraphMarkProperties },
} as const;

const propertiesArbitrary = fc.record({
  styleId: fc.integer({ min: 1, max: 6 }).map((level) => `Heading${level}`),
  alignment: fc.constantFrom(...FOLIO_PARAGRAPH_ALIGNMENT_VALUES),
  spacing: fc.option(
    fc.record({
      spaceBefore: fc.integer({ min: 0, max: 720 }),
      spaceAfter: fc.integer({ min: 0, max: 720 }),
    }),
    { nil: null },
  ),
});
type PropertyEditsOptions = {
  blockId: string;
  text: string;
  properties: fc.Infer<typeof propertiesArbitrary>;
};

// Every operation that edits paragraph properties: a patch, and a style-only
// replaceBlock (the paragraph's own text under another style).
const propertyEdits = ({ blockId, text, properties }: PropertyEditsOptions) => {
  const patches = [
    properties,
    { spacing: properties.spacing },
    { spacing: null },
    { styleId: "Heading1" },
    { styleId: "Heading2" },
  ];
  const restyles = [...new Set([properties.styleId, "Heading2"])].filter(
    (styleId) => styleId !== "Heading1",
  );
  return [
    ...patches.map(
      (patch): FolioDocumentOperation => ({
        id: "properties",
        type: "setBlockParagraphProperties",
        blockId,
        properties: patch,
      }),
    ),
    ...restyles.map(
      (restyle): FolioDocumentOperation => ({
        id: "properties",
        type: "replaceBlock",
        blockId,
        text,
        styleId: restyle,
      }),
    ),
  ];
};

test(
  "separate patches on pending deleted paragraph marks refuse without changing any paragraph",
  async () => {
    await fc.assert(
      fc.asyncProperty(propertiesArbitrary, fc.boolean(), async (properties, atomic) => {
        for (const mode of ["direct", "tracked-changes", "suggested"] as const) {
          for (const [saved, structure] of [
            [false, "merge"],
            [true, "merge"],
            [false, "delete-final"],
            [true, "delete-final"],
          ] as const) {
            let reviewer = await open("# Heading\n\nBody clause.");
            const first = reviewer.getContent().at(0);
            if (!first) throw new Error("Missing heading");
            const original = reviewer.getContent();
            reviewer.applyDocumentOperations({
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode: "tracked-changes",
              operations: [
                structure === "merge"
                  ? { id: "merge", type: "mergeBlockWithNext", blockId: first.id, separator: " " }
                  : {
                      id: "delete",
                      type: "deleteBlock",
                      blockId: reviewer.getContent().at(1)?.id ?? "",
                    },
              ],
            });
            if (saved) reviewer = await reopen(reviewer);
            const current = reviewer.getContent().find(({ id }) => id === first.id);
            if (!current) throw new Error("Missing pending paragraph");
            for (const edit of propertyEdits({
              blockId: first.id,
              text: current.text,
              properties,
            })) {
              const before = reviewer.state.doc.toJSON();
              const changes = reviewer.getChanges();
              const result = reviewer.applyDocumentOperations({
                version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
                mode,
                atomic,
                operations: [edit],
              });
              expect(result.applied).toEqual([]);
              expect(result.skipped.at(0)).toMatchObject({ id: "properties", ...refusedProperty });
              expect(result.issues.at(0)).toMatchObject({
                code: refusedProperty.reason,
                canonicalRefusal: refusedProperty.canonicalRefusal,
                recovery: "resolveTrackedChange",
                retryable: false,
              });
              expect(reviewer.state.doc.toJSON()).toEqual(before);
              expect(reviewer.getChanges()).toEqual(changes);
            }
            reviewer.rejectAll();
            expect(reviewer.getContent()).toEqual(original);
          }
        }
        // The runtime guard also sees a mark created earlier within this batch.
        const heading = (await open("# Heading\n\nBody clause.")).getContent().at(0);
        if (!heading) throw new Error("Missing heading");
        for (const edit of propertyEdits({
          blockId: heading.id,
          text: heading.text,
          properties,
        })) {
          const reviewer = await open("# Heading\n\nBody clause.");
          const original = reviewer.getContent();
          const result = reviewer.applyDocumentOperations({
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "tracked-changes",
            atomic,
            operations: [
              { id: "merge", type: "mergeBlockWithNext", blockId: heading.id, separator: " " },
              edit,
            ],
          });
          expect(result.skipped.find(({ id }) => id === "properties")).toMatchObject(
            refusedProperty,
          );
          if (atomic) expect(reviewer.getContent()).toEqual(original);
          else {
            expect(result.applied.map(({ id }) => id)).toEqual(["merge"]);
            reviewer.acceptAll();
            expect(reviewer.getContent().at(0)?.styleId).toBe("Heading1");
          }
        }
      }),
      propertyConfig({ seed: 1873084145, numRuns: 12 }),
    );
  },
  propertyTestTimeout(30_000),
);

test(
  "atomic merge properties preserve accept-all previews and rejection ownership over generated chains",
  async () => {
    await fc.assert(
      fc.asyncProperty(
        propertiesArbitrary,
        fc.integer({ min: 1, max: 4 }),
        async (properties, chainLength) => {
          const markdown = [
            "# Heading",
            ...Array.from({ length: chainLength }, (_, i) => `Clause ${i}.`),
          ].join("\n\n");
          for (const resolution of [
            "accept",
            "reject",
            "reject-forward",
            "reject-reverse",
            "reject-first-merge",
          ] as const) {
            const reviewer = await open(markdown);
            const heading = reviewer.getContent().at(0);
            if (!heading) throw new Error("Missing heading");
            const initialSpacing = reviewer.applyDocumentOperations({
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode: "direct",
              operations: [
                {
                  id: "initial-spacing",
                  type: "setBlockParagraphProperties",
                  blockId: heading.id,
                  properties: { spacing: { spaceBefore: 240, spaceAfter: 240 } },
                },
              ],
            });
            expect(initialSpacing.skipped).toEqual([]);
            const original = reviewer.getContent();
            const originalParagraphs = paragraphProjection(reviewer.state.doc);
            const operations = original.slice(0, chainLength).map(({ id }, index) =>
              index === 0
                ? {
                    id: `merge-${index}`,
                    type: "mergeBlockWithNext" as const,
                    blockId: id,
                    separator: " ",
                    mergedParagraphProperties: properties,
                  }
                : {
                    id: `merge-${index}`,
                    type: "mergeBlockWithNext" as const,
                    blockId: id,
                    separator: " ",
                  },
            );
            const result = reviewer.applyDocumentOperations({
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode: "tracked-changes",
              operations,
            });
            expect(result.skipped).toEqual([]);
            expect(result.applied).toHaveLength(operations.length);
            const pending = await reopen(reviewer);
            expect(
              paragraphProjection(resolveStateStory(pending.state, "accept").resolved),
            ).toEqual(paragraphProjection(resolveStateStory(reviewer.state, "accept").resolved));
            const revisions = [
              ...new Set(result.applied.flatMap(({ revisionIds }) => revisionIds ?? [])),
            ];
            if (resolution === "reject") reviewer.rejectAll();
            if (resolution === "reject-forward" || resolution === "reject-reverse") {
              for (const id of resolution === "reject-forward" ? revisions : revisions.toReversed())
                reviewer.rejectChange(id);
            }
            if (resolution === "reject-first-merge") {
              const merge = result.applied.find(({ id }) => id === "merge-0");
              if (merge?.revisionId === undefined) throw new Error("Missing merge revision");
              reviewer.rejectChange(merge.revisionId);
            }
            const acceptedPreview = paragraphProjection(
              resolveStateStory(reviewer.state, "accept").resolved,
            );
            reviewer.acceptAll();
            expect(paragraphProjection(reviewer.state.doc)).toEqual(acceptedPreview);
            if (resolution === "accept") {
              expect(reviewer.getContent()).toHaveLength(1);
              expect(reviewer.getContent().at(0)?.styleId).toBe(properties.styleId);
              expect(reviewer.getContent().at(0)?.directAlignment).toBe(properties.alignment);
              expect(reviewer.getContent().at(0)?.directSpacing).toEqual(
                properties.spacing ?? undefined,
              );
            } else if (resolution === "reject-first-merge") {
              expect(reviewer.getContent()).toHaveLength(2);
              expect(paragraphProjection(reviewer.state.doc).at(0)?.properties).toEqual(
                originalParagraphs.at(0)?.properties,
              );
              // Rejecting the paragraph mark keeps the separately accepted separator.
              expect(reviewer.getContent().at(0)?.text).toBe(`${original.at(0)?.text} `);
              expect(paragraphProjection(reviewer.state.doc).at(1)?.properties).toEqual(
                originalParagraphs.at(1)?.properties,
              );
            } else expect(reviewer.getContent()).toEqual(original);
            expect((await reopen(reviewer)).getContent()).toEqual(reviewer.getContent());
          }
        },
      ),
      propertyConfig({ seed: 1873084145, numRuns: 12 }),
    );
  },
  propertyTestTimeout(30_000),
);
