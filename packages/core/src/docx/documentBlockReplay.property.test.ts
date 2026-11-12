/**
 * A canonical save replays the body of `word/document.xml` block by block:
 * unchanged blocks keep their source bytes and changed blocks take the
 * serializer's. The replay reads the generated part once per save and the
 * source part once per captured baseline.
 *
 *   1. For generated source parts (prefixes, quotes, declarations, gaps,
 *      tables, content controls, bookmarks, section properties, background)
 *      and chains of edits on the model (text, insertions, deletions, moves,
 *      duplicates, section and background changes), the replay writes exactly
 *      the bytes of the replay that read both parts on every save, kept only
 *      as the oracle in __fixtures__/storyBlockReplayPerSaveRead.ts.
 *   2. Fixed cases reach each replay branch the generated cases may miss.
 *   3. Work per save: the source part is read on the first save only, the
 *      generated part once per save, and only changed blocks are replayed
 *      from the serializer.
 */

import { describe, expect, setDefaultTimeout, spyOn, test } from "bun:test";
import fc from "fast-check";

import * as streamingXml from "@stll/docx-core/xml";
import { ZipArchive } from "@stll/docx-core/zip";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type { BlockContent, Document, Paragraph } from "../types/document";
import {
  buildDocumentBlockReplay as perSaveReadReplay,
  canReplayDocumentSource as perSaveReadCanReplay,
} from "./__fixtures__/storyBlockReplayPerSaveRead";
import { getDocumentSourceBaseline } from "./headerFooterVerbatim";
import { parseDocx } from "./parser";
import { createEmptyDocx } from "./rezip";
import { serializeDocument } from "./serializer/documentSerializer";
import { readRootNamespaceBindings } from "./serializer/partNamespaces";
import {
  buildDocumentBlockReplay,
  canReplayDocumentSource,
} from "./storyBlockReplay";
import * as verbatimCapture from "./verbatimCapture";

setDefaultTimeout(propertyTestTimeout(60_000));

const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORD_2010 = "http://schemas.microsoft.com/office/word/2010/wordml";

type Syntax = {
  prefix: "w" | "q";
  quote: '"' | "'";
  declaration: "none" | "double" | "single";
};
type ParagraphSpec = {
  paraId: string | undefined;
  text: string;
  bold: boolean;
  centered: boolean;
};
type CellSpec = { width: boolean; paragraphs: ParagraphSpec[] };
type BlockSpec =
  | ({ kind: "paragraph" } & ParagraphSpec)
  | { kind: "table"; rows: CellSpec[][] }
  | { kind: "sdt"; paragraphs: ParagraphSpec[] }
  | { kind: "bookmark"; id: number };
type PartSpec = {
  syntax: Syntax;
  background: boolean;
  blocks: { gap: string; block: BlockSpec }[];
  trailingGap: string;
  sectionProperties: boolean;
  /** `<w:body/>` when the body has no children. */
  selfClosingEmptyBody: boolean;
  /** Text between `</w:body>` and `</w:document>`. */
  afterBody: string;
};

const writePart = (spec: PartSpec): string => {
  const { prefix: p, quote: q } = spec.syntax;
  const attribute = (name: string, value: string) =>
    ` ${name}=${q}${value}${q}`;
  const paragraph = (value: ParagraphSpec) =>
    `<${p}:p${value.paraId === undefined ? "" : attribute("w14:paraId", value.paraId)}>` +
    (value.centered
      ? `<${p}:pPr><${p}:jc${attribute(`${p}:val`, "center")}/></${p}:pPr>`
      : "") +
    `<${p}:r>${value.bold ? `<${p}:rPr><${p}:b/></${p}:rPr>` : ""}<${p}:t>${value.text}</${p}:t></${p}:r></${p}:p>`;
  const block = (value: BlockSpec): string => {
    switch (value.kind) {
      case "paragraph":
        return paragraph(value);
      case "table": {
        const columns = Math.max(...value.rows.map((row) => row.length));
        return (
          `<${p}:tbl><${p}:tblPr><${p}:tblW${attribute(`${p}:w`, "0")}${attribute(`${p}:type`, "auto")}/></${p}:tblPr>` +
          `<${p}:tblGrid>${`<${p}:gridCol${attribute(`${p}:w`, "2000")}/>`.repeat(columns)}</${p}:tblGrid>` +
          value.rows
            .map(
              (row) =>
                `<${p}:tr>${row
                  .map(
                    (cell) =>
                      `<${p}:tc>${cell.width ? `<${p}:tcPr><${p}:tcW${attribute(`${p}:w`, "2000")}${attribute(`${p}:type`, "dxa")}/></${p}:tcPr>` : ""}${cell.paragraphs.map(paragraph).join("")}</${p}:tc>`,
                  )
                  .join("")}</${p}:tr>`,
            )
            .join("") +
          `</${p}:tbl>`
        );
      }
      case "sdt":
        return `<${p}:sdt><${p}:sdtPr><${p}:alias${attribute(`${p}:val`, "Clause")}/></${p}:sdtPr><${p}:sdtContent>${value.paragraphs.map(paragraph).join("")}</${p}:sdtContent></${p}:sdt>`;
      case "bookmark":
        return `<${p}:bookmarkStart${attribute(`${p}:id`, String(value.id))}${attribute(`${p}:name`, `b${value.id}`)}/><${p}:bookmarkEnd${attribute(`${p}:id`, String(value.id))}/>`;
      default: {
        const unreachable: never = value;
        return unreachable;
      }
    }
  };
  const declaration = {
    none: "",
    double: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
    single: "<?xml version='1.0' encoding='UTF-8'?>",
  }[spec.syntax.declaration];
  const children =
    spec.blocks.map(({ gap, block: value }) => gap + block(value)).join("") +
    spec.trailingGap +
    (spec.sectionProperties
      ? `<${p}:sectPr><${p}:pgSz${attribute(`${p}:w`, "12240")}${attribute(`${p}:h`, "15840")}/></${p}:sectPr>`
      : "");
  const body =
    children === "" && spec.selfClosingEmptyBody
      ? `<${p}:body/>`
      : `<${p}:body>${children}</${p}:body>`;
  return (
    declaration +
    `<${p}:document${attribute(`xmlns:${p}`, WORD)}${attribute("xmlns:w14", WORD_2010)}>` +
    (spec.background
      ? `<${p}:background${attribute(`${p}:color`, "FFFFFF")}/>`
      : "") +
    body +
    spec.afterBody +
    `</${p}:document>`
  );
};

const PARA_IDS = ["1A2B3C01", "1A2B3C02", "1A2B3C03", "1A2B3C04"] as const;
const TEXTS = ["same", "other", "third"] as const;
const GAPS = ["", "\n", "\n\t", " <!-- authored gap --> "] as const;

const paragraphSpecArb: fc.Arbitrary<ParagraphSpec> = fc.record({
  paraId: fc.option(fc.constantFrom(...PARA_IDS), { nil: undefined }),
  text: fc.constantFrom(...TEXTS),
  bold: fc.boolean(),
  centered: fc.boolean(),
});
const blockSpecArb: fc.Arbitrary<BlockSpec> = fc.oneof(
  {
    weight: 4,
    arbitrary: paragraphSpecArb.map((value) => ({
      kind: "paragraph" as const,
      ...value,
    })),
  },
  {
    weight: 1,
    arbitrary: fc
      .array(
        fc.array(
          fc.record({
            width: fc.boolean(),
            paragraphs: fc.array(paragraphSpecArb, {
              minLength: 1,
              maxLength: 2,
            }),
          }),
          { minLength: 1, maxLength: 2 },
        ),
        { minLength: 1, maxLength: 2 },
      )
      .map((rows) => ({ kind: "table" as const, rows })),
  },
  {
    weight: 1,
    arbitrary: fc
      .array(paragraphSpecArb, { minLength: 1, maxLength: 2 })
      .map((paragraphs) => ({ kind: "sdt" as const, paragraphs })),
  },
  {
    weight: 1,
    arbitrary: fc
      .integer({ min: 1, max: 3 })
      .map((id) => ({ kind: "bookmark" as const, id })),
  },
);
const partSpecArb: fc.Arbitrary<PartSpec> = fc.record({
  syntax: fc.record({
    prefix: fc.constantFrom("w" as const, "q" as const),
    quote: fc.constantFrom('"' as const, "'" as const),
    declaration: fc.constantFrom(
      "none" as const,
      "double" as const,
      "single" as const,
    ),
  }),
  background: fc.boolean(),
  blocks: fc.array(
    fc.record({ gap: fc.constantFrom(...GAPS), block: blockSpecArb }),
    { maxLength: 6 },
  ),
  trailingGap: fc.constantFrom(...GAPS),
  sectionProperties: fc.boolean(),
  selfClosingEmptyBody: fc.boolean(),
  afterBody: fc.constantFrom("", "\n"),
});

type Edit =
  | { kind: "text"; at: number; text: string }
  | { kind: "insert"; at: number; text: string; paraId: string | undefined }
  | { kind: "delete"; at: number }
  | { kind: "move"; from: number; to: number }
  | { kind: "duplicate"; at: number }
  | { kind: "section"; pageWidth: number | undefined }
  | { kind: "background"; color: string | undefined };

const editArb: fc.Arbitrary<Edit> = fc.oneof(
  fc.record({
    kind: fc.constant("text" as const),
    at: fc.nat(8),
    text: fc.constantFrom(...TEXTS, "edited"),
  }),
  fc.record({
    kind: fc.constant("insert" as const),
    at: fc.nat(8),
    text: fc.constantFrom(...TEXTS, "inserted"),
    paraId: fc.option(fc.constantFrom(...PARA_IDS, "1A2B3C09"), {
      nil: undefined,
    }),
  }),
  fc.record({ kind: fc.constant("delete" as const), at: fc.nat(8) }),
  fc.record({
    kind: fc.constant("move" as const),
    from: fc.nat(8),
    to: fc.nat(8),
  }),
  fc.record({ kind: fc.constant("duplicate" as const), at: fc.nat(8) }),
  fc.record({
    kind: fc.constant("section" as const),
    pageWidth: fc.option(fc.constantFrom(11906, 12240), { nil: undefined }),
  }),
  fc.record({
    kind: fc.constant("background" as const),
    color: fc.option(fc.constantFrom("ABCDEF", "FFFFFF"), { nil: undefined }),
  }),
);

const textRuns = (text: string): Paragraph["content"] => [
  { type: "run", content: [{ type: "text", text }] },
];

const textParagraph = (
  text: string,
  paraId: string | undefined,
): Paragraph => ({
  type: "paragraph",
  ...(paraId === undefined ? {} : { paraId }),
  content: textRuns(text),
});

const setText = (block: BlockContent, text: string): void => {
  switch (block.type) {
    case "paragraph":
      block.content = textRuns(text);
      return;
    case "table": {
      const first = block.rows.at(0)?.cells.at(0)?.content.at(0);
      if (first?.type === "paragraph") first.content = textRuns(text);
      return;
    }
    case "blockSdt":
    case "blockCustomXml": {
      const first = block.content.at(0);
      if (first?.type === "paragraph") first.content = textRuns(text);
      return;
    }
    default:
      return;
  }
};

const applyEdit = (document: Document, edit: Edit): void => {
  const body = document.package.document;
  const { content } = body;
  const index = (at: number) =>
    content.length === 0 ? 0 : at % content.length;
  switch (edit.kind) {
    case "text": {
      const block = content[index(edit.at)];
      if (block) setText(block, edit.text);
      return;
    }
    case "insert":
      content.splice(
        edit.at % (content.length + 1),
        0,
        textParagraph(edit.text, edit.paraId),
      );
      return;
    case "delete":
      if (content.length > 0) content.splice(index(edit.at), 1);
      return;
    case "move": {
      if (content.length === 0) return;
      const [moved] = content.splice(index(edit.from), 1);
      if (moved) content.splice(edit.to % (content.length + 1), 0, moved);
      return;
    }
    case "duplicate": {
      const block = content[index(edit.at)];
      if (block) content.splice(index(edit.at), 0, structuredClone(block));
      return;
    }
    case "section":
      body.finalSectionProperties =
        edit.pageWidth === undefined
          ? undefined
          : { ...body.finalSectionProperties, pageWidth: edit.pageWidth };
      return;
    case "background":
      body.background =
        edit.color === undefined ? undefined : { color: { rgb: edit.color } };
      return;
    default: {
      const unreachable: never = edit;
      return unreachable;
    }
  }
};

const openPart = async (xml: string): Promise<Document> => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set("word/document.xml", xml);
  return parseDocx((await zip.generate({ compression: "store" })).buffer);
};

type SaveComparison = {
  replay: string | null;
  oracle: string | null;
  canReplay: boolean;
  oracleCanReplay: boolean;
  saved: string;
  oracleSaved: string;
};

/** One save of the document by both replays, through the canonical serializer too. */
const compareSave = (document: Document): SaveComparison => {
  const baseline = getDocumentSourceBaseline(document);
  if (baseline.type !== "captured")
    throw new Error(`Source baseline ${baseline.type}`);
  const bindings = readRootNamespaceBindings(baseline.xml);
  const serializedXml = serializeDocument(document, bindings);
  const options = {
    sourceXml: baseline.xml,
    serializedXml,
    baseline: baseline.body,
    current: document.package.document,
  };
  const oracle = perSaveReadReplay(options);
  const oracleCanReplay = perSaveReadCanReplay(baseline.xml, baseline.body);
  const exact =
    JSON.stringify(document.package.document) === baseline.fingerprint;
  return {
    replay: buildDocumentBlockReplay(options),
    oracle,
    canReplay: canReplayDocumentSource(baseline.xml, baseline.body),
    oracleCanReplay,
    saved: serializeDocument(document, bindings, {
      bodyAuthority: "canonical",
      xml: baseline.xml,
    }),
    oracleSaved:
      exact && oracleCanReplay ? baseline.xml : (oracle ?? serializedXml),
  };
};

const expectSameSave = (comparison: SaveComparison): void => {
  expect(comparison.replay).toBe(comparison.oracle);
  expect(comparison.canReplay).toBe(comparison.oracleCanReplay);
  expect(comparison.saved).toBe(comparison.oracleSaved);
};

describe("document block replay", () => {
  test("writes the bytes of the replay that reads both parts every save, over edit chains", async () => {
    await fc.assert(
      fc.asyncProperty(
        partSpecArb,
        fc.array(editArb, { minLength: 1, maxLength: 5 }),
        async (spec, edits) => {
          const document = await openPart(writePart(spec));
          expectSameSave(compareSave(document));
          for (const edit of edits) {
            applyEdit(document, edit);
            expectSameSave(compareSave(document));
          }
        },
      ),
      propertyConfig({ numRuns: 150 }),
    );
  });

  const paragraphSpec = (
    text: string,
    paraId?: string,
  ): ParagraphSpec & { kind: "paragraph" } => ({
    kind: "paragraph",
    paraId,
    text,
    bold: false,
    centered: false,
  });
  const fixedPart = (
    blocks: BlockSpec[],
    overrides: Partial<PartSpec> = {},
  ): PartSpec => ({
    syntax: { prefix: "q", quote: "'", declaration: "single" },
    background: false,
    blocks: blocks.map((block) => ({ gap: "\n", block })),
    trailingGap: "",
    sectionProperties: false,
    selfClosingEmptyBody: false,
    afterBody: "",
    ...overrides,
  });
  const FIXED_CASES: readonly {
    name: string;
    part: PartSpec;
    edits: readonly Edit[];
  }[] = [
    {
      name: "a paragraph's text",
      part: fixedPart([
        paragraphSpec("same", "1A2B3C01"),
        paragraphSpec("other"),
      ]),
      edits: [{ kind: "text", at: 1, text: "edited" }],
    },
    {
      name: "a cell of a table",
      part: fixedPart([
        paragraphSpec("same"),
        {
          kind: "table",
          rows: [
            [
              { width: true, paragraphs: [paragraphSpec("other")] },
              { width: false, paragraphs: [paragraphSpec("third")] },
            ],
          ],
        },
      ]),
      edits: [{ kind: "text", at: 1, text: "edited" }],
    },
    {
      name: "a block content control",
      part: fixedPart([{ kind: "sdt", paragraphs: [paragraphSpec("same")] }]),
      edits: [{ kind: "text", at: 0, text: "edited" }],
    },
    {
      name: "an insertion into a body with no blocks, closed right before the document",
      part: fixedPart([]),
      edits: [{ kind: "insert", at: 0, text: "inserted", paraId: undefined }],
    },
    {
      name: "an insertion into a self-closing body",
      part: fixedPart([], { selfClosingEmptyBody: true }),
      edits: [{ kind: "insert", at: 0, text: "inserted", paraId: undefined }],
    },
    {
      name: "an insertion before the section properties of an empty body",
      part: fixedPart([], { sectionProperties: true }),
      edits: [{ kind: "insert", at: 0, text: "inserted", paraId: undefined }],
    },
    {
      name: "section properties and the background",
      part: fixedPart([paragraphSpec("same")], {
        sectionProperties: true,
        background: true,
      }),
      edits: [
        { kind: "section", pageWidth: 11906 },
        { kind: "background", color: "ABCDEF" },
      ],
    },
    {
      name: "moved and duplicated paragraphs sharing an identity",
      part: fixedPart([
        paragraphSpec("same", "1A2B3C01"),
        paragraphSpec("same", "1A2B3C02"),
        paragraphSpec("other"),
      ]),
      edits: [
        { kind: "move", from: 0, to: 2 },
        { kind: "duplicate", at: 1 },
      ],
    },
  ];

  for (const { name, part, edits } of FIXED_CASES) {
    test(`replays ${name} with the same bytes`, async () => {
      const document = await openPart(writePart(part));
      for (const edit of edits) {
        applyEdit(document, edit);
        const comparison = compareSave(document);
        expect(comparison.oracle).not.toBeNull();
        expectSameSave(comparison);
      }
    });
  }
});

describe("document block replay reads", () => {
  const PARAGRAPHS = 40;
  // The source part spells WordprocessingML with `q:` and the serializer with
  // `w:`, so each read of a whole part or its body names its side.
  const sourcePart = writePart({
    syntax: { prefix: "q", quote: '"', declaration: "double" },
    background: false,
    blocks: Array.from({ length: PARAGRAPHS }, (_, index) => ({
      gap: "\n",
      block: {
        kind: "paragraph" as const,
        paraId: (0x10000000 + index).toString(16).toUpperCase(),
        text: `paragraph ${index}`,
        bold: index % 3 === 0,
        centered: false,
      },
    })),
    trailingGap: "\n",
    sectionProperties: true,
    selfClosingEmptyBody: false,
    afterBody: "",
  });

  type Reads = { source: number; generated: number; replayed: number };

  /** One save's replay, counting the part reads and the replayed blocks. */
  const countedReplay = (document: Document) => {
    const baseline = getDocumentSourceBaseline(document);
    if (baseline.type !== "captured") throw new Error("Missing baseline");
    const options = {
      sourceXml: baseline.xml,
      serializedXml: serializeDocument(
        document,
        readRootNamespaceBindings(baseline.xml),
      ),
      baseline: baseline.body,
      current: document.package.document,
    };
    const reads = spyOn(streamingXml, "parseStreamingXmlWithSourceRanges");
    const replayed = spyOn(verbatimCapture, "isSafeCapturedXmlDocument");
    try {
      const replay = buildDocumentBlockReplay(options);
      const inputs = reads.mock.calls.map(([xml]) => xml);
      const counted: Reads = {
        source: inputs.filter((xml) => xml.includes("<q:body")).length,
        generated: inputs.filter((xml) => xml.includes("<w:body")).length,
        replayed: replayed.mock.calls.length,
      };
      return { replay, oracle: perSaveReadReplay(options), counted };
    } finally {
      reads.mockRestore();
      replayed.mockRestore();
    }
  };

  const edit = (document: Document, at: number, text: string) => {
    const block = document.package.document.content[at];
    if (!block) throw new Error(`Missing block ${at}`);
    setText(block, text);
  };

  test("a first save reads each part once and replays only the changed block", async () => {
    const document = await openPart(sourcePart);
    edit(document, 5, "edited");
    const { replay, oracle, counted } = countedReplay(document);
    expect(replay).not.toBeNull();
    expect(replay).toBe(oracle);
    expect(counted).toEqual({ source: 1, generated: 1, replayed: 1 });
  });

  test("later saves read only the generated part and replay only changed blocks", async () => {
    const document = await openPart(sourcePart);
    edit(document, 5, "edited");
    countedReplay(document);
    const saves: Reads[] = [];
    for (const [save, at] of [10, 20, 30].entries()) {
      edit(document, at, `edited in save ${save}`);
      const { replay, oracle, counted } = countedReplay(document);
      expect(replay).toBe(oracle);
      saves.push(counted);
    }
    expect(saves).toEqual([
      { source: 0, generated: 1, replayed: 2 },
      { source: 0, generated: 1, replayed: 3 },
      { source: 0, generated: 1, replayed: 4 },
    ]);
  });
});
