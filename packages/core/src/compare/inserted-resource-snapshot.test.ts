import { expect, test } from "bun:test";

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
    if (!insertedBlock) throw new Error("Missing insertion fixture block");
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
