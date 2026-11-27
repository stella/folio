import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  applyDocumentOp,
  applyDocumentOps,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  DOCUMENT_OP_REFUSAL_REASONS,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import { propertyConfig } from "../../../../test/property-testing";
import type { Document, ParagraphContent, Table } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";

const read = (buffer: ArrayBuffer) =>
  parseDocx(buffer, { detectVariables: false, preloadFonts: false });

const operation = {
  type: DOCUMENT_OP_TYPES.MERGE_CELLS,
  story: OP_STORIES.MAIN,
  blockId: "00000001",
  top: 0,
  bottom: 1,
  left: 0,
  right: 2,
} as const;
const intent = { type: "table", operation } as const;
const revision = { author: "Reviewer", date: "2026-11-27T00:00:00Z" };

const documentWith = (content: ParagraphContent[]): Document => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "table",
      columnWidths: [1000, 1200],
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              content: [{ type: "paragraph", paraId: "00000001", content: [] }],
            },
            {
              type: "tableCell",
              content: [{ type: "paragraph", paraId: "00000002", content }],
            },
          ],
        },
      ],
    },
    { type: "paragraph", paraId: "00000003", content: [] },
  ];
  return document;
};

const tableOf = (document: Document): Table => {
  const table = document.package.document.content.at(0);
  if (table?.type !== "table") throw new Error("Expected a table.");
  return table;
};

const cellContentOf = (document: Document) =>
  tableOf(document).rows.map((row) =>
    row.cells.map((cell) =>
      cell.content.map((block) =>
        block.type === "paragraph"
          ? { id: block.paraId, content: block.content }
          : block,
      ),
    ),
  );

const allocate = createEditorIntentIdAllocator();

test("tracked table moves give nested controls unique IDs and restore their authored IDs after reopening", async () => {
  await fc.assert(
    fc.asyncProperty(
      fc.integer({ min: 1, max: 3 }),
      fc.integer({ min: -2048, max: -10 }),
      fc.boolean(),
      async (depth, firstId, statesKind) => {
        let content: ParagraphContent[] = [
          { type: "run", content: [{ type: "text", text: "Controlled text" }] },
        ];
        for (let index = depth - 1; index >= 0; index -= 1)
          content = [
            {
              type: "inlineSdt",
              properties: {
                sdtType: "richText",
                id: firstId + index,
                alias: `Clause ${index}`,
                ...(statesKind
                  ? {
                      preserved: {
                        children: [{ index: 18, xml: "<w:richText/>" }],
                      },
                    }
                  : {}),
              },
              content,
            },
          ];
        const original = await read(await createDocx(documentWith(content)));
        const ids = allocate(original, intent);
        expect(ids.newIds.control).toHaveLength(depth);
        const compiled = compileEditorIntent(original, {
          intent,
          mode: {
            type: "suggesting",
            revision: { ...revision, id: ids.revisionId },
            newIds: ids.newIds,
          },
        }).unwrap();
        const pending = applyDocumentOps(original, compiled.ops).unwrap();
        expect(
          applyDocumentOps(pending.document, pending.inverse).unwrap().document,
        ).toStrictEqual(original);
        const direct = applyDocumentOp(original, {
          ...operation,
          newBlockIds: [],
        }).unwrap().document;
        const file = await createDocx(pending.document);
        const archive = await ZipArchive.read(file);
        const part = archive.entry("word/document.xml");
        if (part === undefined) throw new Error("Expected word/document.xml.");
        const xml = await part.text();
        expect(xml).toContain("<w:moveFrom ");
        expect(xml).toContain("<w:moveTo ");
        const reopened = await read(file);
        const geometry =
          tableOf(reopened).formatting?.gridChange?.geometryChange;
        const move =
          geometry?.type === "mergeCells" ? geometry.moves?.at(0) : undefined;
        if (move === undefined || move.kind === "empty")
          throw new Error("Expected a paragraph move.");
        expect(move.controls).toEqual(
          ids.newIds.control.map((id, index) => [firstId + index, id]),
        );
        const authoredIds = Array.from(
          { length: depth },
          (_, index) => firstId + index,
        );
        for (const id of [...authoredIds, ...ids.newIds.control]) {
          const marker = `<w:id w:val="${id}"/>`;
          expect(xml.split(marker)).toHaveLength(2);
        }
        for (const decision of [
          REVISION_DECISIONS.ACCEPT,
          REVISION_DECISIONS.REJECT,
        ]) {
          const expected =
            decision === REVISION_DECISIONS.ACCEPT ? direct : original;
          const decide = (document: Document) =>
            applyDocumentOp(document, {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: [move.toRevision],
              decision,
            }).unwrap();
          const decided = decide(pending.document);
          expect(decided.document).toStrictEqual(expected);
          expect(
            applyDocumentOps(decided.document, decided.inverse).unwrap()
              .document,
          ).toStrictEqual(pending.document);
          expect(cellContentOf(decide(reopened).document)).toStrictEqual(
            cellContentOf(expected),
          );
        }
      },
    ),
    propertyConfig({ numRuns: 8 }),
  );
});

test("an inline control with opaque behavior remains outside the tracked move scope", () => {
  const original = documentWith([
    {
      type: "inlineSdt",
      properties: {
        sdtType: "richText",
        id: 42,
        preserved: {
          children: [
            {
              index: 8,
              xml: "<w:temporary/>",
            },
          ],
        },
      },
      content: [{ type: "run", content: [{ type: "text", text: "Bound" }] }],
    },
  ]);
  const ids = allocate(original, intent);
  expect(ids.newIds.control).toHaveLength(0);
  const result = compileEditorIntent(original, {
    intent,
    mode: {
      type: "suggesting",
      revision: { ...revision, id: ids.revisionId },
      newIds: ids.newIds,
    },
  });
  if (result.isOk()) throw new Error("Expected the opaque control to refuse.");
  expect(result.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
});

test.each([
  { control: [], reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS },
  { control: [42], reason: DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION },
])(
  "checks the temporary control IDs before recording a merge ($reason)",
  ({ control, reason }) => {
    const original = documentWith([
      {
        type: "inlineSdt",
        properties: { sdtType: "richText", id: 42 },
        content: [{ type: "run", content: [{ type: "text", text: "Moved" }] }],
      },
    ]);
    const ids = allocate(original, intent);
    const result = applyDocumentOp(original, {
      ...operation,
      newBlockIds: ["00000004"],
      revision: { ...revision, id: ids.revisionId },
      newIds: { ...ids.newIds, control },
    });
    if (result.isOk()) throw new Error("Expected the temporary IDs to refuse.");
    expect(result.error.reason).toBe(reason);
  },
);

test.each([
  { kind: "footnote", reference: "footnoteRef" },
  { kind: "endnote", reference: "endnoteRef" },
] as const)(
  "a native table move retains the single $kind body for both tracked references",
  async ({ kind, reference }) => {
    const noteId = 8;
    const source = documentWith([
      { type: "run", content: [{ type: reference, id: noteId }] },
    ]);
    const content = [
      {
        type: "paragraph",
        paraId: "00000008",
        content: [
          {
            type: "run",
            content: [
              { type: "noteMarker", kind },
              { type: "text", text: " Retained note" },
            ],
          },
        ],
      },
    ] satisfies Document["package"]["document"]["content"];
    if (kind === "footnote")
      source.package.footnotes = [{ type: kind, id: noteId, content }];
    else source.package.endnotes = [{ type: kind, id: noteId, content }];
    const original = await read(await createDocx(source));
    const ids = allocate(original, intent);
    const compiled = compileEditorIntent(original, {
      intent,
      mode: {
        type: "suggesting",
        revision: { ...revision, id: ids.revisionId },
        newIds: ids.newIds,
      },
    }).unwrap();
    const pending = applyDocumentOps(original, compiled.ops).unwrap();
    const direct = applyDocumentOp(original, {
      ...operation,
      newBlockIds: [],
    }).unwrap().document;
    const file = await createDocx(pending.document);
    const archive = await ZipArchive.read(file);
    const main = archive.entry("word/document.xml");
    if (main === undefined) throw new Error("Expected the document part.");
    const xml = await main.text();
    expect(xml.split(`<w:${kind}Reference w:id="${noteId}"`)).toHaveLength(3);
    const reopened = await read(file);
    for (const decision of [
      REVISION_DECISIONS.ACCEPT,
      REVISION_DECISIONS.REJECT,
    ]) {
      const op = {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: pending.revisions,
        decision,
      } as const;
      const expected =
        decision === REVISION_DECISIONS.ACCEPT ? direct : original;
      const settled = applyDocumentOp(pending.document, op).unwrap();
      expect(settled.document).toStrictEqual(expected);
      expect(
        applyDocumentOps(settled.document, settled.inverse).unwrap().document,
      ).toStrictEqual(pending.document);
      const reopenedSettled = applyDocumentOp(reopened, op).unwrap().document;
      expect(cellContentOf(reopenedSettled)).toEqual(cellContentOf(expected));
      const notes =
        kind === "footnote"
          ? reopenedSettled.package.footnotes
          : reopenedSettled.package.endnotes;
      expect(notes?.filter(({ id }) => id === noteId)).toHaveLength(1);
    }
  },
);

test("a tracked move keeps both control copies bound to the same saved custom XML part", async () => {
  const store = "{A0000000-0000-0000-0000-000000000001}";
  const data = "<record>Bound</record>";
  const binding = `<w:dataBinding w:storeItemID="${store}" w:xpath="/record[1]"/>`;
  const source = documentWith([
    {
      type: "inlineSdt",
      properties: {
        sdtType: "richText",
        id: 42,
        preserved: { children: [{ index: 8, xml: binding }] },
      },
      content: [{ type: "run", content: [{ type: "text", text: "Bound" }] }],
    },
  ]);
  const seed = await ZipArchive.read(await createDocx(source));
  const relationships = seed.entry("word/_rels/document.xml.rels");
  const types = seed.entry("[Content_Types].xml");
  if (relationships === undefined || types === undefined)
    throw new Error("Expected package metadata.");
  seed.set(
    "word/_rels/document.xml.rels",
    (await relationships.text()).replace(
      "</Relationships>",
      '<Relationship Id="rId88" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>',
    ),
  );
  seed.set(
    "[Content_Types].xml",
    (await types.text()).replace(
      "</Types>",
      '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>',
    ),
  );
  seed.set("customXml/item1.xml", data);
  seed.set(
    "customXml/itemProps1.xml",
    `<ds:datastoreItem ds:itemID="${store}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"/>`,
  );
  seed.set(
    "customXml/_rels/item1.xml.rels",
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>',
  );
  const original = await read((await seed.generate()).buffer);
  const ids = allocate(original, intent);
  expect(ids.newIds.control).toHaveLength(1);
  const compiled = compileEditorIntent(original, {
    intent,
    mode: {
      type: "suggesting",
      revision: { ...revision, id: ids.revisionId },
      newIds: ids.newIds,
    },
  }).unwrap();
  const pending = applyDocumentOps(original, compiled.ops).unwrap();
  const direct = applyDocumentOp(original, {
    ...operation,
    newBlockIds: [],
  }).unwrap().document;
  const file = await createDocx(pending.document);
  const archive = await ZipArchive.read(file);
  expect(await archive.entry("customXml/item1.xml")?.text()).toBe(data);
  expect(
    (await archive.entry("word/document.xml")?.text())?.split("<w:dataBinding"),
  ).toHaveLength(3);
  const reopened = await read(file);
  for (const decision of [
    REVISION_DECISIONS.ACCEPT,
    REVISION_DECISIONS.REJECT,
  ]) {
    const op = {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: pending.revisions,
      decision,
    } as const;
    const settled = applyDocumentOp(pending.document, op).unwrap();
    const expected = decision === REVISION_DECISIONS.ACCEPT ? direct : original;
    expect(settled.document).toStrictEqual(expected);
    expect(
      applyDocumentOps(settled.document, settled.inverse).unwrap().document,
    ).toStrictEqual(pending.document);
    expect(
      cellContentOf(applyDocumentOp(reopened, op).unwrap().document),
    ).toEqual(cellContentOf(expected));
  }
});
