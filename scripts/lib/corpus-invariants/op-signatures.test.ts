import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../../packages/docx-core/src/model/document";
import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  type DocumentOp,
} from "../../../packages/docx-core/src/ops/documentOps";
import { RELATIONSHIP_TYPES } from "@stll/folio-core/docx/relsParser";
import { serializedInverseSequenceFailures, serializedInverseStepFailures } from "./op-inverse";
import { localityStepFailures, serializedLocalityStepFailures } from "./op-locality";
import { firstDifferingOpPart } from "./op-part-difference";
import {
  exactOpModel,
  serializeOpDocument,
  type OpSequence,
  type OpSequenceStep,
} from "./op-sequences";

const paragraph = (paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text: "text" }] }],
});

const fixture = () =>
  ({
    package: {
      document: { content: [paragraph("00000001")], background: { themeTint: "AA" } },
      headers: new Map([
        ["rIdHeader", { type: "header", hdrFtrType: "default", content: [paragraph("00000002")] }],
      ]),
      footers: new Map([
        ["rIdFooter", { type: "footer", hdrFtrType: "default", content: [paragraph("00000003")] }],
      ]),
      footnotes: [{ type: "footnote", id: 1, content: [paragraph("00000004")] }],
      endnotes: [{ type: "endnote", id: 1, content: [paragraph("00000005")] }],
      properties: { title: "retained" },
    },
  }) satisfies Document;

const stepFor = (op: DocumentOp, before: Document = fixture()) =>
  ({
    before,
    op,
    edit: applyDocumentOp(before, op).unwrap(),
  }) satisfies OpSequenceStep;

const insert = {
  type: DOCUMENT_OP_TYPES.INSERT_TEXT,
  at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
  text: "added",
  runProps: INHERIT_RUN_PROPS,
} as const satisfies DocumentOp;

const bytes = (text: string) => new TextEncoder().encode(text);
const parts = () =>
  new Map([
    ["custom/main.xml", bytes("body")],
    [
      "custom/_rels/main.xml.rels",
      bytes(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="${RELATIONSHIP_TYPES.header}" Target="../stories/header3.xml"/><Relationship Id="rIdFooter" Type="${RELATIONSHIP_TYPES.footer}" Target="../stories/footer4.xml"/><Relationship Id="rIdNotes" Type="${RELATIONSHIP_TYPES.footnotes}" Target="../stories/notes5.xml"/><Relationship Id="rIdEndnotes" Type="${RELATIONSHIP_TYPES.endnotes}" Target="../stories/endnotes6.xml"/></Relationships>`,
      ),
    ],
    [
      "[Content_Types].xml",
      bytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
    ],
    ["stories/header3.xml", bytes("header")],
    ["stories/footer4.xml", bytes("footer")],
    ["stories/notes5.xml", bytes("notes")],
    ["stories/endnotes6.xml", bytes("endnotes")],
  ]);

test("first differing part is deterministic and includes additions, removals and binary changes", () => {
  const control = new Map([
    ["word/styles.xml", bytes("styles")],
    ["word/media/image27.png", new Uint8Array([1, 2])],
  ]);
  expect(firstDifferingOpPart({ control, edited: new Map(control) })).toBeUndefined();
  for (const change of ["added", "removed", "changed"] as const) {
    const edited = new Map([...control].toReversed());
    if (change === "added") edited.set("word/header27.xml", bytes("header"));
    if (change === "removed") edited.delete("word/media/image27.png");
    if (change === "changed") edited.set("word/media/image27.png", new Uint8Array([1, 3]));
    expect(firstDifferingOpPart({ control, edited })).toBe(
      change === "added" ? "word/header27.xml" : "word/media/image27.png",
    );
  }
  const edited = new Map(control);
  edited.set("word/styles.xml", bytes("changed styles"));
  edited.set("word/header9.xml", bytes("header"));
  expect(firstDifferingOpPart({ control, edited })).toBe("word/header9.xml");
});

test("serialized inverse signatures retain the operation kind and first differing part", () => {
  const step = stepFor(insert);
  const control = parts();
  const restored = new Map(control);
  restored.set("stories/footer4.xml", bytes("changed"));
  expect(serializedInverseStepFailures({ step, control, restored })).toEqual([
    "insertText inverse changed the original serialized package parts: stories/footer4.xml",
  ]);
  expect(serializedInverseStepFailures({ step, control, restored: new Map(control) })).toEqual([]);
});

test("secondary story scope permits declared edits and rejects an unrelated touched identity", () => {
  const cases = [
    {
      story: { kind: "header", rId: "rIdHeader" },
      blockId: "00000002",
      part: "stories/header3.xml",
    },
    {
      story: { kind: "footer", rId: "rIdFooter" },
      blockId: "00000003",
      part: "stories/footer4.xml",
    },
    { story: { kind: "footnote", id: 1 }, blockId: "00000004", part: "stories/notes5.xml" },
    { story: { kind: "endnote", id: 1 }, blockId: "00000005", part: "stories/endnotes6.xml" },
  ] as const;
  for (const { story, blockId, part } of cases) {
    const step = stepFor({ ...insert, at: { story, blockId, offset: 1 } });
    expect(localityStepFailures(step)).toEqual([]);
    const control = parts();
    const edited = new Map(control);
    edited.set(part, bytes("declared change"));
    expect(
      serializedLocalityStepFailures({ step, control, edited, documentPart: "custom/main.xml" }),
    ).toEqual([]);
    const changedMain = new Map(edited);
    changedMain.set("custom/main.xml", bytes("unrelated main story"));
    expect(
      serializedLocalityStepFailures({
        step,
        control,
        edited: changedMain,
        documentPart: "custom/main.xml",
      }),
    ).toEqual(["insertText changed unrelated serialized part: custom/main.xml"]);
    edited.set("word/comments3.xml", bytes("unrelated change"));
    expect(
      serializedLocalityStepFailures({ step, control, edited, documentPart: "custom/main.xml" }),
    ).toEqual(["insertText changed unrelated serialized part: word/comments3.xml"]);
    const changed = structuredClone(step.edit.document);
    const main = changed.package.document.content.at(0);
    if (main?.type !== "paragraph") throw new Error("Missing main fixture paragraph");
    main.formatting = { alignment: "end" };
    const bad = {
      ...step,
      edit: {
        ...step.edit,
        document: changed,
        touched: { ...step.edit.touched, modified: [blockId, "00000001"] },
      },
    };
    expect(localityStepFailures(bad)).toContain(
      "insertText declared a touched block outside its addressed story",
    );
  }
});

test("restoration owns only supplied payload fields and collections", () => {
  const step = stepFor({
    type: DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
    expected: { body: { background: { themeTint: "AA" } } },
    parts: { body: { background: null } },
  });
  expect(localityStepFailures(step)).toEqual([]);
  const changed = structuredClone(step.edit.document);
  changed.package.properties = { title: "unowned" };
  expect(localityStepFailures({ ...step, edit: { ...step.edit, document: changed } })).toContain(
    "restoreStoryParts changed records outside its declared story and section fields",
  );
  const notes = fixture().package.footnotes;
  const restoreNotes = stepFor({
    type: DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
    expected: { footnotes: notes },
    parts: { footnotes: null },
  });
  expect(localityStepFailures(restoreNotes)).toEqual([]);
  const control = parts();
  const edited = new Map(control);
  edited.delete("stories/notes5.xml");
  expect(
    serializedLocalityStepFailures({
      step: restoreNotes,
      control,
      edited,
      documentPart: "custom/main.xml",
    }),
  ).toEqual([]);
  edited.set("stories/endnotes6.xml", bytes("unowned"));
  expect(
    serializedLocalityStepFailures({
      step: restoreNotes,
      control,
      edited,
      documentPart: "custom/main.xml",
    }),
  ).toEqual(["restoreStoryParts changed unrelated serialized part: stories/endnotes6.xml"]);
});

test("body-field and section-field ownership cannot conceal paragraph edits in touched sets", () => {
  const before: Document = fixture();
  const first = before.package.document.content.at(0);
  if (first?.type !== "paragraph") throw new Error("Missing main fixture paragraph");
  first.sectionProperties = { pageWidth: 12000 };
  before.package.document.sections = [
    { content: before.package.document.content, properties: first.sectionProperties },
  ];
  const section = stepFor(
    { type: DOCUMENT_OP_TYPES.SET_SECTION_PROPS, sectionIndex: 0, patch: { pageWidth: 13000 } },
    before,
  );
  const restoreSection = stepFor(
    {
      type: DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
      expected: { sections: [{ index: 0, properties: { pageWidth: 13000 } }] },
      parts: { sections: [{ index: 0, properties: { pageWidth: 13000 } }] },
    },
    section.edit.document,
  );
  const restoreBackground = stepFor({
    type: DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
    expected: { body: { background: { themeTint: "AA" } } },
    parts: { body: { background: null } },
  });
  for (const step of [section, restoreSection, restoreBackground]) {
    expect(localityStepFailures(step)).toEqual([]);
    const changed = structuredClone(step.edit.document);
    const target = changed.package.document.content.at(0);
    if (target?.type !== "paragraph") throw new Error("Missing main fixture paragraph");
    target.content = [{ type: "run", content: [{ type: "text", text: "unowned change" }] }];
    const corrupted = {
      ...step,
      edit: {
        ...step.edit,
        document: changed,
        touched: { ...step.edit.touched, modified: ["00000001"] },
      },
    };
    expect(localityStepFailures(corrupted)).toContain(
      `${step.op.type} changed records outside its declared story and section fields`,
    );
  }
});

test("compound inverse diagnosis identifies the first failing reverse undo instead of later cascades", async () => {
  const first = stepFor(insert);
  const second = stepFor(
    {
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      patch: { alignment: "end" },
    },
    first.edit.document,
  );
  const broken = { ...second, edit: { ...second.edit, inverse: [] } };
  const sequence = {
    original: first.before,
    originalModel: exactOpModel(first.before),
    originalXml: serializeOpDocument(first.before),
    document: second.edit.document,
    steps: [first, broken],
    inverse: [...first.edit.inverse],
    mutations: [],
    refusals: [],
  } satisfies OpSequence;
  const serializeParts = async (document: Document) =>
    new Map([["word/document.xml", bytes(serializeOpDocument(document))]]);
  const control = await serializeParts(first.before);
  const restored = await serializeParts(second.edit.document);
  expect(
    await serializedInverseSequenceFailures({ sequence, control, restored, serializeParts }),
  ).toEqual([
    "setParagraphProps inverse changed the original serialized package parts: word/document.xml",
  ]);
  let serializations = 0;
  expect(
    await serializedInverseSequenceFailures({
      sequence,
      control,
      restored: new Map(control),
      serializeParts: async (document) => {
        serializations += 1;
        return serializeParts(document);
      },
    }),
  ).toEqual([]);
  expect(serializations).toBe(0);
  expect(
    await serializedInverseSequenceFailures({
      sequence: { ...sequence, steps: [] },
      control,
      restored,
      serializeParts,
    }),
  ).toEqual([
    "empty sequence composition changed the original serialized package parts: word/document.xml",
  ]);
});

test("the first locality part is chosen across relationship, content type and byte failures", () => {
  const header = fixture().package.headers;
  const step = stepFor({
    type: DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
    expected: { headers: [...header] },
    parts: { headers: null },
  });
  const control = parts();
  const edited = new Map(control);
  edited.set(
    "custom/_rels/main.xml.rels",
    bytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFooter" Type="${RELATIONSHIP_TYPES.footer}" Target="../stories/changed-footer.xml"/></Relationships>`,
    ),
  );
  edited.set(
    "[Content_Types].xml",
    bytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/unowned.xml" ContentType="application/xml"/></Types>',
    ),
  );
  edited.set("stories/footer4.xml", bytes("changed"));
  expect(
    serializedLocalityStepFailures({ step, control, edited, documentPart: "custom/main.xml" }),
  ).toEqual(["restoreStoryParts changed unrelated package content types: [Content_Types].xml"]);
});
