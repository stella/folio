import { expect, setDefaultTimeout, test } from "bun:test";
import { panic } from "better-result";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { expectParagraphBlock } from "../../../../test/paragraphBlock";
import {
  insertedTextBoxDocument,
  INSERTED_TEXT_BOX_STYLE,
} from "../__tests__/insertedTextBoxDocument";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";

import {
  createFolioAIEditSnapshotWithStyleResolver,
  remapFolioAIEditSnapshotStyleReferences,
  sourceDocumentOf,
} from "../ai-edits/snapshot";
import { schema } from "../prosemirror/schema";
import { createStyleResolver } from "../prosemirror/styles/styleResolver";
import {
  createInsertedResourceProjection,
  rebindInsertedResourceBlocks,
} from "./inserted-resource-snapshot";

const SOURCE_STYLE = "SourceNumbered";
const IMPORTED_STYLE = "ImportedNumbered";
const ANCHOR_STYLE = "Normal";
const sourceResolver = createStyleResolver({
  styles: [
    { styleId: ANCHOR_STYLE, type: "paragraph" },
    { styleId: SOURCE_STYLE, type: "paragraph", rPr: { bold: true } },
  ],
});
const importedResolver = createStyleResolver({
  styles: [
    { styleId: ANCHOR_STYLE, type: "paragraph" },
    { styleId: IMPORTED_STYLE, type: "paragraph", rPr: { bold: true } },
  ],
});

test("inserted resource rebinding preserves full-source identities across duplicate and positional IDs", () => {
  for (const identity of ["duplicate", "positional"] as const) {
    // Construct raw PM nodes: public DOCX loading repairs package IDs before
    // this projection boundary and therefore cannot exercise these collisions.
    const leading = schema.node(
      "paragraph",
      { paraId: identity === "duplicate" ? "A1000001" : null, styleId: ANCHOR_STYLE },
      [schema.text("Leading unchanged anchor")],
    );
    const insertion = schema.node(
      "paragraph",
      {
        paraId: identity === "duplicate" ? "A1000001" : null,
        ...(identity === "positional" && { idStability: "positional" }),
        styleId: SOURCE_STYLE,
      },
      [schema.text("Inserted styled paragraph")],
    );
    const trailing = schema.node(
      "paragraph",
      { paraId: identity === "duplicate" ? "A1000003" : null, styleId: ANCHOR_STYLE },
      [schema.text("Trailing unchanged anchor")],
    );
    const base = createFolioAIEditSnapshotWithStyleResolver(
      schema.node("doc", null, [leading, trailing]),
      sourceResolver,
    );
    const revised = createFolioAIEditSnapshotWithStyleResolver(
      schema.node("doc", null, [leading, insertion, trailing]),
      sourceResolver,
    );
    const originalBlocks = revised.blocks.map((block) => ({ ...block }));
    const insertedBlock = revised.blocks.at(1);
    expect(insertedBlock?.id).toBe("seq-0002");
    if (identity === "positional") expect(insertedBlock?.idStability).toBe("positional");

    const projection = createInsertedResourceProjection({ base, revised });
    expect(projection.revisedBlockIds).toEqual([insertedBlock?.id]);
    expect(projection.snapshot.blocks).toHaveLength(1);
    expect(sourceDocumentOf(projection.snapshot).firstChild).toBe(insertion);
    // The old subset-generated-ID join cannot name the source insertion.
    expect(projection.snapshot.blocks.at(0)?.id).not.toBe(insertedBlock?.id);

    const rebound = remapFolioAIEditSnapshotStyleReferences({
      snapshot: projection.snapshot,
      styleIdMap: new Map([[SOURCE_STYLE, IMPORTED_STYLE]]),
      defaultParagraphStyleId: undefined,
      importedStyleResolver: importedResolver,
    });
    const mapped = rebindInsertedResourceBlocks(projection, rebound);
    expect([...mapped.keys()]).toEqual([insertedBlock?.id]);
    if (!insertedBlock) panic("Missing insertion fixture block");
    expect(mapped.get(insertedBlock.id)).toMatchObject({
      id: insertedBlock.id,
      text: insertedBlock.text,
      styleId: IMPORTED_STYLE,
    });
    expect(mapped.size).toBe(1);
    expect(mapped.has(revised.blocks.at(0)?.id ?? "")).toBe(false);
    expect(mapped.has(revised.blocks.at(2)?.id ?? "")).toBe(false);
    expect(revised.blocks).toEqual(originalBlocks);
    expect(sourceDocumentOf(revised).child(0)).toBe(leading);
    expect(sourceDocumentOf(revised).child(2)).toBe(trailing);
    expect(leading.attrs["styleId"]).toBe(ANCHOR_STYLE);
    expect(trailing.attrs["styleId"]).toBe(ANCHOR_STYLE);
    expect(insertion.attrs["styleId"]).toBe(SOURCE_STYLE);
  }
});

setDefaultTimeout(propertyTestTimeout(30_000));

test("inserted nested carriers project and rebind every paragraph exactly once", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.stringMatching(/^[a-z]{1,20}$/u),
      fc.integer({ min: 1, max: 4 }),
      async (suffix, innerCount) => {
        for (const innerContent of ["paragraph", "tableCell"] as const) {
          for (const carrierCount of [1, 2]) {
            const baseDocument = insertedTextBoxDocument({
              side: "base",
              innerContent,
              suffix,
              innerCount,
              carrierCount,
            });
            const revisedDocument = insertedTextBoxDocument({
              side: "revised",
              innerContent,
              suffix,
              innerCount,
              carrierCount,
            });
            const base = createFolioAIEditSnapshotWithStyleResolver(
              toProseDoc(baseDocument),
              createStyleResolver(baseDocument.package.styles),
            );
            const revised = createFolioAIEditSnapshotWithStyleResolver(
              toProseDoc(revisedDocument),
              createStyleResolver(revisedDocument.package.styles),
            );
            const insertedBlocks = revised.blocks.filter(({ text }) => text !== "Unchanged anchor");
            const expectedTexts = Array.from({ length: carrierCount }, (_, carrierIndex) => {
              const carrierSuffix =
                carrierCount === 1 ? suffix : `${suffix} carrier ${carrierIndex + 1}`;
              return [
                `Parent text ${carrierSuffix}`,
                ...Array.from(
                  { length: innerCount },
                  (_innerValue, innerIndex) =>
                    `Inner text ${carrierSuffix}${innerCount === 1 ? "" : ` ${innerIndex + 1}`}`,
                ),
              ];
            }).flat();
            const expectedCount = carrierCount * (1 + innerCount);
            expect(insertedBlocks.map(({ text }) => text)).toEqual(expectedTexts);
            const projection = createInsertedResourceProjection({ base, revised });
            expect(projection.snapshot.blocks).toHaveLength(expectedCount);
            expect(projection.revisedBlockIds).toEqual(insertedBlocks.map(({ id }) => id));
            expect(new Set(projection.revisedBlockIds).size).toBe(expectedCount);
            const rebound = remapFolioAIEditSnapshotStyleReferences({
              snapshot: projection.snapshot,
              styleIdMap: new Map([[INSERTED_TEXT_BOX_STYLE, IMPORTED_STYLE]]),
              defaultParagraphStyleId: undefined,
              importedStyleResolver: importedResolver,
            });
            const mapped = rebindInsertedResourceBlocks(projection, rebound);
            expect(mapped.size).toBe(expectedCount);
            for (const original of insertedBlocks) {
              expect(mapped.get(original.id)).toMatchObject({
                id: original.id,
                text: original.text,
                styleId: IMPORTED_STYLE,
              });
            }
            expect(revised.blocks.filter(({ text }) => text === "Unchanged anchor")).toEqual(
              base.blocks,
            );
            for (const block of insertedBlocks) {
              expect(expectParagraphBlock(block).styleId).toBe(INSERTED_TEXT_BOX_STYLE);
            }
          }
        }
      },
    ),
    {
      // Vary nested paragraph count; every run covers both shapes and sibling counts.
      numRuns: 4,
    },
  );
});
