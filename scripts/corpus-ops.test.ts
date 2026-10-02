import { propertyTestTimeout } from "../test/property-testing";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { describe, expect, test } from "bun:test";
import type { BlockContent, Document, Paragraph } from "../packages/docx-core/src/model/document";
import { DEFAULT_TAB_STOP_TWIPS } from "../packages/docx-core/src/model/document";
import {
  applyDocumentOp,
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  INHERIT_RUN_PROPS,
  normalizeForOps,
  OP_STORIES,
} from "../packages/docx-core/src/ops/documentOps";
import { buildBodySequenceDocx } from "@stll/folio-core/compare/__fixtures__/body-sequence";
import { parseDocx } from "@stll/folio-core/docx/parser";

import type { CorpusInvariantInput } from "./lib/corpus-invariants/contract";
import { inverseSequenceFailures, runOpInverseInvariant } from "./lib/corpus-invariants/op-inverse";
import {
  localityStepFailures,
  runOpLocalityInvariant,
  serializedLocalityFailures,
} from "./lib/corpus-invariants/op-locality";
import {
  generateOpSequence,
  OP_SEQUENCE_FAMILIES,
  OP_GENERATOR_ROLES,
  prepareOpDocument,
  sameOpModel,
  serializedOpParts,
  seedFromBytes,
} from "./lib/corpus-invariants/op-sequences";

const makeParagraph = (paraId: string, text: string) =>
  ({
    type: "paragraph",
    paraId,
    content: [{ type: "run", content: [{ type: "text", text }] }],
  }) satisfies Paragraph;

const documentFixture = () =>
  normalizeForOps({
    package: {
      document: {
        content: [
          makeParagraph("60000001", "First paragraph 🧭."),
          makeParagraph("60000002", "Second paragraph é."),
          {
            type: "table",
            rows: [
              {
                type: "tableRow",
                cells: [
                  { type: "tableCell", content: [makeParagraph("60000003", "Left cell")] },
                  { type: "tableCell", content: [makeParagraph("60000004", "Right cell")] },
                ],
              },
              {
                type: "tableRow",
                cells: [
                  { type: "tableCell", content: [makeParagraph("60000005", "Second row")] },
                  { type: "tableCell", content: [makeParagraph("60000006", "")] },
                ],
              },
            ],
          },
          {
            type: "blockSdt",
            properties: { sdtType: "richText", id: 9001, tag: "corpus-sequence" },
            content: [makeParagraph("60000007", "Inside a content control")],
          },
          {
            type: "blockCustomXml",
            openingXml: '<w:customXml w:element="record" w:uri="urn:corpus">',
            closingXml: "</w:customXml>",
            content: [makeParagraph("60000008", "Inside a transparent wrapper")],
          },
          {
            type: "paragraph",
            paraId: "60000009",
            content: [
              {
                type: "insertion",
                info: { id: 9002, author: "Fixture", date: "2026-01-01T00:00:00Z" },
                content: [{ type: "run", content: [{ type: "text", text: "Reviewed text" }] }],
              },
            ],
          },
          makeParagraph("6000000a", "Final paragraph"),
        ],
      },
      headers: new Map([
        [
          "rId10",
          {
            type: "header",
            hdrFtrType: "default",
            content: [makeParagraph("60000010", "Header paragraph")],
          },
        ],
      ]),
      properties: { title: "Operation corpus fixture" },
    },
  } satisfies Document);

const insertionStep = (blockId = "60000001") => {
  const before = documentFixture();
  const op = {
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: { story: OP_STORIES.MAIN, blockId, offset: 3 },
    text: "x",
    runProps: INHERIT_RUN_PROPS,
  } as const;
  const result = applyDocumentOp(before, op);
  if (result.isErr()) throw result.error;
  return { before, op, edit: result.value };
};

const inputFor = async (buffer: ArrayBuffer): Promise<CorpusInvariantInput> => ({
  bytes: new Uint8Array(buffer),
  buffer,
  parsed: await parseDocx(buffer, { preloadFonts: false }),
  documentPart: "word/document.xml",
  budgetMs: 30_000,
});

describe("corpus operation invariants", () => {
  test("omitted section patches do not own fields or hide foreign changes", () => {
    const before = documentFixture();
    const op = {
      type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
      sectionIndex: 0,
      patch: { evenAndOddHeaders: true },
    } as const;
    const result = applyDocumentOp(before, op).unwrap();
    const step = { before, op, edit: result };
    expect(localityStepFailures(step)).toEqual([]);
    const changed = structuredClone(result.document);
    const section = changed.package.document.finalSectionProperties;
    if (section === undefined) throw new Error("Section fixture missing");
    section.titlePg = true;
    expect(
      localityStepFailures({ ...step, edit: { ...result, document: changed } }).length,
    ).toBeGreaterThan(0);
  });

  test("settings locality covers absent and authored settings without hiding foreign defaults", () => {
    for (const settings of [undefined, { defaultTabStop: 900, mirrorMargins: true }]) {
      for (const evenAndOddHeaders of [true, false, null]) {
        const before = documentFixture();
        if (settings !== undefined) before.package.settings = settings;
        const op = {
          type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
          sectionIndex: 0,
          patch: { evenAndOddHeaders },
        } as const;
        const result = applyDocumentOp(before, op);
        if (result.isErr()) throw result.error;
        const step = { before, op, edit: result.value };
        expect(localityStepFailures(step)).toEqual([]);
        const changed = structuredClone(result.value.document);
        changed.package.settings = {
          ...(changed.package.settings ?? { defaultTabStop: DEFAULT_TAB_STOP_TWIPS }),
          defaultTabStop: 123,
        };
        expect(
          localityStepFailures({ ...step, edit: { ...step.edit, document: changed } }).length,
        ).toBeGreaterThan(0);
      }
    }
  });

  test(
    "seeded sequences cover every declared family and preserve exact inverse/locality",
    () => {
      const document = documentFixture();
      const snapshot = structuredClone(document);
      const exercised = new Set<string>();
      const structuralFamilies = [
        DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
        DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER,
        DOCUMENT_OP_TYPES.ADD_NOTE,
        DOCUMENT_OP_TYPES.REMOVE_NOTE,
        DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
        DOCUMENT_OP_TYPES.DELETE_BLOCKS,
        DOCUMENT_OP_TYPES.INSERT_TABLE,
        DOCUMENT_OP_TYPES.DELETE_TABLE,
        DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
      ];
      const checkedStructuralFamilies = new Set<string>();
      for (let seed = 0; seed < 64; seed += 1) {
        const sequence = generateOpSequence(document, seed);
        expect(sequence.steps.length).toBeGreaterThan(0);
        expect(sequence.mutations).toEqual([]);
        const failures = inverseSequenceFailures(sequence);
        expect(failures, `sequence seed ${seed}`).toEqual([]);
        for (const step of sequence.steps) {
          exercised.add(step.op.type);
          if (
            structuralFamilies.some((type) => type === step.op.type) &&
            !sameOpModel(step.before, step.edit.document)
          ) {
            expect(step.edit.inverse.length).toBeGreaterThan(0);
            expect(
              inverseSequenceFailures({
                ...sequence,
                steps: [{ ...step, edit: { ...step.edit, inverse: [] } }],
              }).length,
            ).toBeGreaterThan(0);
            checkedStructuralFamilies.add(step.op.type);
          }
          expect(localityStepFailures(step)).toEqual([]);
        }
        const replay = generateOpSequence(document, seed);
        expect(sameOpModel(sequence, replay)).toBe(true);
      }
      expect([...checkedStructuralFamilies].sort()).toEqual([...structuralFamilies].sort());
      expect([...exercised].sort()).toEqual([...OP_SEQUENCE_FAMILIES].sort());
      expect(
        Object.entries(OP_GENERATOR_ROLES)
          .filter(([, role]) => role === "generated")
          .map(([type]) => type)
          .sort(),
      ).toEqual([...OP_SEQUENCE_FAMILIES].sort());
      expect(Object.keys(OP_GENERATOR_ROLES).sort()).toEqual(
        Object.values(DOCUMENT_OP_TYPES).sort(),
      );
      expect(sameOpModel(document, snapshot)).toBe(true);
    },
    propertyTestTimeout(30_000),
  );

  test("lifecycle locality permits declared story parts and rejects unrelated model and ZIP changes", async () => {
    const buffer = await createDocx(documentFixture());
    const document = await prepareOpDocument(await inputFor(buffer));
    const sequence = generateOpSequence(document, 0);
    const lifecycle = sequence.steps.find(
      ({ op }) =>
        op.type === DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER ||
        op.type === DOCUMENT_OP_TYPES.ADD_NOTE ||
        op.type === DOCUMENT_OP_TYPES.SET_SECTION_PROPS,
    );
    if (!lifecycle) throw new Error("The seeded schedule must execute a lifecycle operation.");
    const changed = structuredClone(lifecycle.edit.document);
    changed.package.properties = { ...changed.package.properties, title: "foreign change" };
    expect(
      localityStepFailures({ ...lifecycle, edit: { ...lifecycle.edit, document: changed } }).length,
    ).toBeGreaterThan(0);
    const control = await serializedOpParts(sequence.original);
    const edited = await serializedOpParts(sequence.document);
    expect(
      serializedLocalityFailures({ sequence, control, edited, documentPart: "word/document.xml" }),
    ).toEqual([]);
    const corrupt = new Map(edited);
    corrupt.set("word/styles.xml", new TextEncoder().encode("foreign style change"));
    expect(
      serializedLocalityFailures({
        sequence,
        control,
        edited: corrupt,
        documentPart: "word/document.xml",
      }),
    ).toContain("sequence changed unrelated serialized part: word/styles.xml");
  });

  test("row inverses retain the outer table when nested tables precede its surviving anchor", () => {
    for (let depth = 1; depth <= 3; depth += 1) {
      const before = documentFixture();
      const table = before.package.document.content.at(2);
      if (table?.type !== "table") throw new Error("fixture table is missing");
      const survivingRow = table.rows.at(1);
      const cell = survivingRow?.cells.at(0);
      if (cell === undefined) throw new Error("fixture cell is missing");
      let nested: BlockContent = makeParagraph("61000000", "Innermost cell");
      for (let index = 0; index < depth; index += 1) {
        nested = {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: [
                {
                  type: "tableCell",
                  content: [nested, makeParagraph(`6100000${index + 1}`, "Nested final paragraph")],
                },
              ],
            },
          ],
        };
      }
      cell.content.unshift(nested);
      const op = {
        type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS,
        story: OP_STORIES.MAIN,
        blockId: "60000003",
        expected: table.rows,
        rows: table.rows.slice(1),
      } as const;
      const applied = applyDocumentOp(before, op);
      if (applied.isErr()) throw applied.error;
      const inverse = applied.value.inverse.at(0);
      expect(inverse?.type).toBe(DOCUMENT_OP_TYPES.SET_TABLE_ROWS);
      if (inverse?.type !== DOCUMENT_OP_TYPES.SET_TABLE_ROWS)
        throw new Error("row inverse is missing");
      expect(inverse.blockId).toBe("60000005");
      const restored = applyDocumentOps(applied.value.document, applied.value.inverse);
      if (restored.isErr()) throw restored.error;
      expect(sameOpModel(restored.value.document, before)).toBe(true);
      expect(localityStepFailures({ before, op, edit: applied.value })).toEqual([]);
    }
  });

  test("the byte-derived seed depends on package contents deterministically", () => {
    const bytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x00, 0x61]);
    expect(seedFromBytes(bytes)).toBe(seedFromBytes(Uint8Array.from(bytes)));
    const changed = Uint8Array.from(bytes);
    changed[5] = 0x62;
    expect(seedFromBytes(changed)).not.toBe(seedFromBytes(bytes));
  });

  test("an incomplete inverse cannot pass the sequence detector", () => {
    const sequence = generateOpSequence(documentFixture(), 0x17a3);
    expect(inverseSequenceFailures(sequence)).toEqual([]);
    expect(inverseSequenceFailures({ ...sequence, inverse: [] }).length).toBeGreaterThan(0);
    const brokenStep = insertionStep();
    expect(
      inverseSequenceFailures({
        ...sequence,
        steps: [{ ...brokenStep, edit: { ...brokenStep.edit, inverse: [] } }],
      }).length,
    ).toBeGreaterThan(0);
  });

  test("a missing touched block cannot hide an actual serialized edit", () => {
    const step = insertionStep();
    expect(localityStepFailures(step)).toEqual([]);
    expect(
      localityStepFailures({
        ...step,
        edit: { ...step.edit, touched: { modified: [], inserted: [], removed: [] } },
      }).length,
    ).toBeGreaterThan(0);
  });

  test("the locality detector catches an untouched paragraph and a sibling table cell", () => {
    for (const targetId of ["60000002", "60000004"]) {
      const step = insertionStep(targetId === "60000004" ? "60000003" : "60000001");
      const document = structuredClone(step.edit.document);
      const mutate = (blocks: typeof document.package.document.content): boolean => {
        for (const block of blocks) {
          if (block.type === "paragraph" && block.paraId === targetId) {
            block.content = makeParagraph(targetId, "Unexpected replacement").content;
            return true;
          }
          if (block.type === "table") {
            for (const row of block.rows) {
              for (const cell of row.cells) {
                if (mutate(cell.content)) return true;
              }
            }
          }
          if (
            (block.type === "blockSdt" || block.type === "blockCustomXml") &&
            mutate(block.content)
          ) {
            return true;
          }
        }
        return false;
      };
      expect(mutate(document.package.document.content)).toBe(true);
      expect(
        localityStepFailures({ ...step, edit: { ...step.edit, document } }).length,
      ).toBeGreaterThan(0);
    }
  });

  test("the locality detector catches unrelated package metadata and header changes", () => {
    const step = insertionStep();
    const metadata = structuredClone(step.edit.document);
    metadata.package.properties = { title: "Unexpected title" };
    expect(
      localityStepFailures({ ...step, edit: { ...step.edit, document: metadata } }).length,
    ).toBeGreaterThan(0);
    const headerDocument = structuredClone(step.edit.document);
    const header = headerDocument.package.headers?.get("rId10");
    if (header === undefined) throw new Error("fixture header is missing");
    header.content = [makeParagraph("60000010", "Unexpected header")];
    expect(
      localityStepFailures({ ...step, edit: { ...step.edit, document: headerDocument } }).length,
    ).toBeGreaterThan(0);
  });

  test("corpus preparation and both invariants handle a real package with rows and a header", async () => {
    const buffer = await buildBodySequenceDocx(
      [
        { kind: "paragraph", text: "First paragraph." },
        {
          kind: "table",
          rows: [
            ["Left", "Right"],
            ["Second row", ""],
          ],
        },
        { kind: "paragraph", text: "Final paragraph." },
      ],
      { header: [{ kind: "paragraph", text: "Header paragraph." }] },
    );
    const input = await inputFor(buffer);
    const prepared = await prepareOpDocument(input);
    expect(await prepareOpDocument(input)).toBe(prepared);
    expect((await runOpInverseInvariant(input)).failures).toEqual([]);
    expect((await runOpLocalityInvariant(input)).failures).toEqual([]);
  });
});
