/**
 * Block operations on and next to a paragraph that carries an inline object.
 *
 * A paragraph's drawings are its content: deleting the paragraph deletes them,
 * and inserting, splitting or merging beside it keeps each one in the
 * paragraph that holds it. A text box is the case that differs in the editor:
 * its paragraphs are a sibling block of the carrier, tied to it by an anchor,
 * and are read as blocks of their own. So every kind the model has is run
 * through the same operations, in direct and tracked mode, and checked twice:
 *
 * - the requested outcome in the saved package (for a tracked edit: after
 *   accept-all, live and reopened, and after accepting one change at a time;
 *   the original back after reject-all);
 * - the reader's ordered blocks, live against the same package reopened.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "../docx/parser";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import type { Paragraph, ParagraphContent, Run, RunContent, Shape } from "../types/content";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const PNG_DATA_URL =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const FLOATING_POSITION = {
  horizontal: { relativeTo: "column", posOffset: 0 },
  vertical: { relativeTo: "paragraph", posOffset: 0 },
} as const;

const textBoxShape = (floating: boolean): Shape => ({
  type: "shape",
  shapeType: "textBox",
  id: "42",
  name: "Carried box",
  size: { width: 1_828_800, height: 914_400 },
  ...(floating
    ? { position: FLOATING_POSITION, wrap: { type: "square", wrapText: "bothSides" } }
    : {}),
  textBody: {
    content: [
      {
        type: "paragraph",
        paraId: "34000001",
        content: [{ type: "run", content: [{ type: "text", text: "Box words." }] }],
      },
    ],
  },
});

const CHART_XML =
  '<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/>' +
  '<wp:docPr id="77" name="Carried chart"/><a:graphic>' +
  '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing>';

/** Every inline object the model carries in a run. */
const OBJECT_KINDS = {
  "inline image": (): RunContent => ({
    type: "drawing",
    image: {
      type: "image",
      src: PNG_DATA_URL,
      docPrName: "Carried image",
      size: { width: 914_400, height: 457_200 },
      wrap: { type: "inline" },
    },
  }),
  "anchored image": (): RunContent => ({
    type: "drawing",
    image: {
      type: "image",
      src: PNG_DATA_URL,
      docPrName: "Carried image",
      size: { width: 914_400, height: 457_200 },
      wrap: { type: "square", wrapText: "bothSides" },
      position: FLOATING_POSITION,
    },
  }),
  "chart (preserved drawing)": (): RunContent => ({
    type: "drawing",
    image: { type: "image", size: { width: 914_400, height: 457_200 }, wrap: { type: "inline" } },
    rawXml: CHART_XML,
    rawXmlMode: "preserveOnly",
  }),
  shape: (): RunContent => ({
    type: "shape",
    shape: {
      type: "shape",
      shapeType: "rect",
      id: "43",
      name: "Carried shape",
      size: { width: 914_400, height: 457_200 },
    },
  }),
  "inline text box": (): RunContent => ({ type: "shape", shape: textBoxShape(false) }),
  "anchored text box": (): RunContent => ({ type: "shape", shape: textBoxShape(true) }),
} as const satisfies Record<string, () => RunContent>;

type ObjectKind = keyof typeof OBJECT_KINDS;
type Placement = "leading" | "trailing";

const LEAD = "Lead paragraph.";
const CARRIER = "Carrier words.";
const TAIL = "Tail paragraph.";
const INSERTED = "Inserted.";

const toArrayBuffer = (bytes: Uint8Array): ArrayBuffer =>
  bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

const buildPackage = async (kind: ObjectKind, placement: Placement): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  // Each in a run of its own, the way a drawing is usually authored.
  const words: Run = { type: "run", content: [{ type: "text", text: CARRIER }] };
  const object: Run = { type: "run", content: [OBJECT_KINDS[kind]()] };
  document.package.document.content = [
    paragraph(LEAD),
    { type: "paragraph", content: placement === "leading" ? [object, words] : [words, object] },
    paragraph(TAIL),
  ];
  const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
  return toArrayBuffer(docx);
};

// ---------------------------------------------------------------------------
// What the saved package says
// ---------------------------------------------------------------------------

/** One body paragraph: its own words, and the objects it carries, in order. */
type SavedParagraph = { text: string; objects: string[] };

const describeObject = (content: RunContent): string | null => {
  if (content.type === "drawing") {
    return content.rawXmlMode === "preserveOnly" ? "chart" : "image";
  }
  if (content.type !== "shape") {
    return null;
  }
  const body = content.shape.textBody;
  if (!body) {
    return "shape";
  }
  const words = body.content
    .flatMap((block) => (block.type === "paragraph" ? [paragraphText(block)] : []))
    .join("|");
  return `box(${words})`;
};

const inlineChildren = (item: ParagraphContent): readonly ParagraphContent[] => {
  switch (item.type) {
    case "hyperlink":
      return item.children;
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo":
    case "inlineSdt":
    case "simpleField":
    case "inlineWrapper":
      return item.content;
    default:
      return [];
  }
};

const visitRuns = (
  content: readonly ParagraphContent[],
  visit: (runContent: RunContent) => void,
): void => {
  for (const item of content) {
    if (item.type === "run") {
      item.content.forEach(visit);
      continue;
    }
    visitRuns(inlineChildren(item), visit);
  }
};

const paragraphText = (block: Paragraph): string => {
  let text = "";
  visitRuns(block.content, (content) => {
    if (content.type === "text") {
      text += content.text;
    }
  });
  return text;
};

const savedBody = async (bytes: ArrayBuffer): Promise<SavedParagraph[]> => {
  const document = await parseDocx(bytes, { detectVariables: false, preloadFonts: false });
  return document.package.document.content.flatMap((block) => {
    if (block.type !== "paragraph") {
      return [];
    }
    const objects: string[] = [];
    visitRuns(block.content, (content) => {
      const described = describeObject(content);
      if (described) {
        objects.push(described);
      }
    });
    return [{ text: paragraphText(block), objects }];
  });
};

const hasRevisions = async (bytes: ArrayBuffer): Promise<boolean> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes);
  return reviewer.getChanges().length > 0;
};

// ---------------------------------------------------------------------------
// What the reader says
// ---------------------------------------------------------------------------

/** The reader's blocks in order, with where each one lives. */
const readerOrder = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ id, kind, text, containerPath }) => ({
    id,
    kind,
    text,
    containerPath: containerPath ?? [],
  }));

const blockIdOf = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((candidate) => candidate.text === text);
  if (!block) {
    throw new Error(`no block reads "${text}"`);
  }
  return block.id;
};

// ---------------------------------------------------------------------------
// The operations
// ---------------------------------------------------------------------------

type Mode = "direct" | "tracked-changes";

type Scenario = {
  name: string;
  operation: (reviewer: FolioDocxReviewer) => FolioDocumentOperation;
  /** The body the operation asks for, `"object"` marking the carrier's object. */
  expected: (placement: Placement) => readonly (readonly [string, "object" | null])[];
  /**
   * Where the operation is refused instead. A text box is a block of its own
   * after its paragraph, and a tracked break across it cannot be resolved:
   * rejecting a split or accepting a merge joins two paragraphs, and the
   * resolver joins only neighbours.
   */
  refused?: (kind: ObjectKind, placement: Placement, mode: Mode) => boolean;
};

const isTextBox = (kind: ObjectKind): boolean => kind.endsWith("text box");

const withObject = (placement: Placement, first: string, second: string) =>
  placement === "leading" ? first : second;

const SCENARIOS: readonly Scenario[] = [
  {
    name: "deleteBlock on the carrier",
    operation: (reviewer) => ({
      id: "op",
      type: "deleteBlock",
      blockId: blockIdOf(reviewer, CARRIER),
    }),
    expected: () => [
      [LEAD, null],
      [TAIL, null],
    ],
  },
  {
    name: "insertAfterBlock on the carrier",
    operation: (reviewer) => ({
      id: "op",
      type: "insertAfterBlock",
      blockId: blockIdOf(reviewer, CARRIER),
      text: INSERTED,
    }),
    expected: () => [
      [LEAD, null],
      [CARRIER, "object"],
      [INSERTED, null],
      [TAIL, null],
    ],
  },
  {
    name: "insertBeforeBlock on the carrier",
    operation: (reviewer) => ({
      id: "op",
      type: "insertBeforeBlock",
      blockId: blockIdOf(reviewer, CARRIER),
      text: INSERTED,
    }),
    expected: () => [
      [LEAD, null],
      [INSERTED, null],
      [CARRIER, "object"],
      [TAIL, null],
    ],
  },
  {
    name: "insertAfterBlock on the paragraph before the carrier",
    operation: (reviewer) => ({
      id: "op",
      type: "insertAfterBlock",
      blockId: blockIdOf(reviewer, LEAD),
      text: INSERTED,
    }),
    expected: () => [
      [LEAD, null],
      [INSERTED, null],
      [CARRIER, "object"],
      [TAIL, null],
    ],
  },
  {
    name: "insertBeforeBlock on the paragraph after the carrier",
    operation: (reviewer) => ({
      id: "op",
      type: "insertBeforeBlock",
      blockId: blockIdOf(reviewer, TAIL),
      text: INSERTED,
    }),
    expected: () => [
      [LEAD, null],
      [CARRIER, "object"],
      [INSERTED, null],
      [TAIL, null],
    ],
  },
  {
    name: "splitBlock inside the carrier",
    operation: (reviewer) => ({
      id: "op",
      type: "splitBlock",
      blockId: blockIdOf(reviewer, CARRIER),
      offset: "Carrier".length,
      separator: " ",
    }),
    expected: (placement) => [
      [LEAD, null],
      ["Carrier", withObject(placement, "object", null) as "object" | null],
      ["words.", withObject(placement, null, "object") as "object" | null],
      [TAIL, null],
    ],
    // Only where the box's anchor stays in the first half.
    refused: (kind, placement, mode) =>
      isTextBox(kind) && placement === "leading" && mode === "tracked-changes",
  },
  {
    name: "mergeBlockWithNext on the carrier",
    operation: (reviewer) => ({
      id: "op",
      type: "mergeBlockWithNext",
      blockId: blockIdOf(reviewer, CARRIER),
      separator: " ",
    }),
    expected: () => [
      [LEAD, null],
      [`${CARRIER} ${TAIL}`, "object"],
    ],
    refused: (kind) => isTextBox(kind),
  },
  {
    name: "mergeBlockWithNext on the paragraph before the carrier",
    operation: (reviewer) => ({
      id: "op",
      type: "mergeBlockWithNext",
      blockId: blockIdOf(reviewer, LEAD),
      separator: " ",
    }),
    expected: () => [
      [`${LEAD} ${CARRIER}`, "object"],
      [TAIL, null],
    ],
  },
];

const ORIGINAL = [
  [LEAD, null],
  [CARRIER, "object"],
  [TAIL, null],
] as const;

const expectedBody = (
  rows: readonly (readonly [string, "object" | null])[],
  object: string,
): SavedParagraph[] => rows.map(([text, carries]) => ({ text, objects: carries ? [object] : [] }));

const objectDescription = (kind: ObjectKind): string => {
  const content = OBJECT_KINDS[kind]();
  const described = describeObject(content);
  if (!described) {
    throw new Error(`${kind} describes no object`);
  }
  return described;
};

const reopen = async (reviewer: FolioDocxReviewer): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), { author: "Reviewer" });

const PLACEMENTS: readonly Placement[] = ["trailing", "leading"];
const MODES: readonly Mode[] = ["direct", "tracked-changes"];

/** Accept the pending changes one revision at a time. */
const acceptEachChange = (reviewer: FolioDocxReviewer): void => {
  for (let guard = 0; guard < 20; guard += 1) {
    const change = reviewer.getChanges().at(0);
    if (!change) {
      return;
    }
    expect(reviewer.acceptChange(change)).toBe(true);
  }
  throw new Error("accepting one change at a time did not converge");
};

describe("block operations on and next to an inline object's paragraph", () => {
  for (const kind of Object.keys(OBJECT_KINDS) as ObjectKind[]) {
    const object = objectDescription(kind);
    for (const placement of PLACEMENTS) {
      for (const scenario of SCENARIOS) {
        for (const mode of MODES) {
          test(`${kind}, ${placement}: ${scenario.name} (${mode})`, async () => {
            const source = await buildPackage(kind, placement);
            expect(await savedBody(source)).toEqual(expectedBody(ORIGINAL, object));

            const reviewer = await FolioDocxReviewer.fromBuffer(source, { author: "Reviewer" });
            const result = reviewer.applyDocumentOperations({
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode,
              operations: [scenario.operation(reviewer)],
            });
            const reopened = await reopen(reviewer);
            expect(readerOrder(reopened)).toEqual(readerOrder(reviewer));

            if (scenario.refused?.(kind, placement, mode)) {
              expect(result.skipped).toEqual([{ id: "op", reason: "unsupportedBlock" }]);
              expect(result.applied).toEqual([]);
              expect(await savedBody(await reviewer.toBuffer())).toEqual(
                expectedBody(ORIGINAL, object),
              );
              return;
            }
            expect(result.skipped).toEqual([]);
            expect(result.applied.map(({ id }) => id)).toEqual(["op"]);

            const expected = expectedBody(scenario.expected(placement), object);
            if (mode === "direct") {
              const saved = await reviewer.toBuffer();
              expect(await hasRevisions(saved)).toBe(false);
              expect(await savedBody(saved)).toEqual(expected);
              return;
            }

            const accepted = await reopen(reopened);
            accepted.acceptAll();
            reviewer.acceptAll();
            expect(readerOrder(accepted)).toEqual(readerOrder(reviewer));
            for (const resolved of [accepted, reviewer]) {
              const acceptedBytes = await resolved.toBuffer();
              expect(await hasRevisions(acceptedBytes)).toBe(false);
              expect(await savedBody(acceptedBytes)).toEqual(expected);
            }

            const acceptedOneByOne = await reopen(reopened);
            acceptEachChange(acceptedOneByOne);
            expect(await savedBody(await acceptedOneByOne.toBuffer())).toEqual(expected);

            const rejected = await reopen(reopened);
            rejected.rejectAll();
            const rejectedBytes = await rejected.toBuffer();
            expect(await hasRevisions(rejectedBytes)).toBe(false);
            expect(await savedBody(rejectedBytes)).toEqual(expectedBody(ORIGINAL, object));
          });
        }
      }
    }
  }
});

describe("a paragraph holding a text box, first in the body", () => {
  const BOX = "Box words.";

  const buildTextBoxCarrierFirst = async (): Promise<ArrayBuffer> => {
    const document = createEmptyDocument();
    document.package.document.content = [
      {
        type: "paragraph",
        content: [
          { type: "run", content: [{ type: "text", text: CARRIER }] },
          { type: "run", content: [{ type: "shape", shape: textBoxShape(false) }] },
        ],
      },
      paragraph(TAIL),
    ];
    const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
    return toArrayBuffer(docx);
  };

  const bodyXml = async (bytes: ArrayBuffer): Promise<string> => {
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("word/document.xml")?.async("string");
    if (xml === undefined) {
      throw new Error("the package has no main document part");
    }
    return xml;
  };

  const texts = (reviewer: FolioDocxReviewer): string[] =>
    reviewer.getContent().map(({ text }) => text);

  test("deleting it directly deletes the box", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildTextBoxCarrierFirst());
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [{ id: "delete", type: "deleteBlock", blockId: blockIdOf(reviewer, CARRIER) }],
    });

    expect(result.applied.map(({ id }) => id)).toEqual(["delete"]);
    expect(texts(reviewer)).toEqual([TAIL]);
    const saved = await reviewer.toBuffer();
    const xml = await bodyXml(saved);
    expect(xml).not.toContain(BOX);
    expect(xml).not.toContain("<w:drawing");
    expect(texts(await FolioDocxReviewer.fromBuffer(saved))).toEqual([TAIL]);
  });

  test("deleting it tracked marks the box deleted with it", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildTextBoxCarrierFirst(), {
      author: "Reviewer",
    });
    reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [{ id: "delete", type: "deleteBlock", blockId: blockIdOf(reviewer, CARRIER) }],
    });
    const pending = await reviewer.toBuffer();
    // The drawing run is inside the paragraph's deletion.
    expect(await bodyXml(pending)).toMatch(/<w:del [^>]*>(?:(?!<\/w:del>).)*<w:drawing/s);

    const accepted = await FolioDocxReviewer.fromBuffer(pending);
    accepted.acceptAll();
    const acceptedXml = await bodyXml(await accepted.toBuffer());
    expect(acceptedXml).not.toContain(BOX);
    expect(acceptedXml).not.toContain(CARRIER);
    expect(texts(accepted)).toEqual([TAIL]);

    const rejected = await FolioDocxReviewer.fromBuffer(pending);
    rejected.rejectAll();
    expect(texts(rejected)).toEqual([CARRIER, BOX, TAIL]);
    expect(await savedBody(await rejected.toBuffer())).toEqual([
      { text: CARRIER, objects: [`box(${BOX})`] },
      { text: TAIL, objects: [] },
    ]);
  });

  for (const type of ["insertAfterBlock", "insertBeforeBlock"] as const) {
    test(`${type} keeps the reader's order across a save`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await buildTextBoxCarrierFirst());
      reviewer.applyDocumentOperations({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [{ id: "insert", type, blockId: blockIdOf(reviewer, CARRIER), text: INSERTED }],
      });

      expect(texts(reviewer)).toEqual(
        type === "insertAfterBlock"
          ? [CARRIER, BOX, INSERTED, TAIL]
          : [INSERTED, CARRIER, BOX, TAIL],
      );
      const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
      expect(readerOrder(reopened)).toEqual(readerOrder(reviewer));
    });
  }
});
