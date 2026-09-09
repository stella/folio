import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import type { Document, Paragraph, SectionProperties, Table } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";

const OPTIONS = {
  author: "Section reviewer",
  timestamp: "2026-09-09T12:00:00.000Z",
} as const;

const SECTION_A = {
  sectionStart: "nextPage",
  pageWidth: 10_000,
  pageHeight: 14_000,
  marginLeft: 720,
} as const satisfies SectionProperties;

const SECTION_B = {
  sectionStart: "oddPage",
  pageWidth: 11_000,
  pageHeight: 15_000,
  columnCount: 2,
  columnSpace: 500,
} as const satisfies SectionProperties;

const FOLLOWING_SECTION = {
  sectionStart: "continuous",
  pageWidth: 12_000,
  pageHeight: 16_000,
  marginRight: 900,
  titlePg: true,
} as const satisfies SectionProperties;

const paragraph = (
  paraId: string,
  text: string,
  sectionProperties?: SectionProperties,
): Paragraph => ({
  type: "paragraph",
  paraId,
  textId: paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
  ...(sectionProperties === undefined ? {} : { sectionProperties }),
});

const table = (text: string): Table => ({
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [
        {
          type: "tableCell",
          content: [paragraph("10000001", text)],
        },
      ],
    },
  ],
});

const documentWith = (content: Document["package"]["document"]["content"]): Document => {
  const document = createEmptyDocument();
  return {
    ...document,
    package: {
      ...document.package,
      document: { ...document.package.document, content },
    },
  };
};

const paragraphSectionProjection = (reviewer: FolioDocxReviewer) =>
  reviewer
    .toDocument()
    .package.document.content.filter((block): block is Paragraph => block.type === "paragraph")
    .map(({ paraId, sectionProperties }) => ({
      paraId,
      sectionProperties:
        sectionProperties === undefined
          ? null
          : Object.fromEntries(Object.entries(sectionProperties)),
    }));

const storyProjection = (reviewer: FolioDocxReviewer) => ({
  blocks: reviewer.snapshot().blocks.map(({ id, kind, text }) => ({ id, kind, text })),
  sections: paragraphSectionProjection(reviewer),
  topLevelKinds: reviewer.toDocument().package.document.content.map(({ type }) => type),
});

const reopenResolved = async (
  buffer: ArrayBuffer,
  mode: "accept" | "reject",
): Promise<FolioDocxReviewer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  if (mode === "accept") {
    reviewer.acceptAll();
  } else {
    reviewer.rejectAll();
  }
  const first = await reviewer.toBuffer();
  const second = await reviewer.toBuffer();
  const [concurrentFirst, concurrentSecond] = await Promise.all([
    reviewer.toBuffer(),
    reviewer.toBuffer(),
  ]);
  const documentXml = async (saved: ArrayBuffer) =>
    (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text");
  const expectedXml = await documentXml(first);
  expect(await documentXml(second)).toBe(expectedXml);
  expect(await documentXml(concurrentFirst)).toBe(expectedXml);
  expect(await documentXml(concurrentSecond)).toBe(expectedXml);
  const [reopenedFirst, reopenedSecond, reopenedConcurrentFirst, reopenedConcurrentSecond] =
    await Promise.all([
      FolioDocxReviewer.fromBuffer(first),
      FolioDocxReviewer.fromBuffer(second),
      FolioDocxReviewer.fromBuffer(concurrentFirst),
      FolioDocxReviewer.fromBuffer(concurrentSecond),
    ]);
  const expectedProjection = storyProjection(reopenedFirst);
  expect(storyProjection(reopenedSecond)).toEqual(expectedProjection);
  expect(storyProjection(reopenedConcurrentFirst)).toEqual(expectedProjection);
  expect(storyProjection(reopenedConcurrentSecond)).toEqual(expectedProjection);
  return reopenedFirst;
};

const expectResolvedTerminals = async (baseDocument: Document, targetDocument: Document) => {
  const [base, target] = await Promise.all([createDocx(baseDocument), createDocx(targetDocument)]);
  const [baseReviewer, targetReviewer] = await Promise.all([
    FolioDocxReviewer.fromBuffer(base),
    FolioDocxReviewer.fromBuffer(target),
  ]);
  const result = await compareDocx(base, target, OPTIONS);
  if (result.isErr()) {
    throw result.error;
  }
  expect(result.value.verification).toEqual({ status: "verified" });
  const pendingXml = await (
    await JSZip.loadAsync(result.value.buffer)
  )
    .file("word/document.xml")
    ?.async("text");
  expect(pendingXml).toMatch(/<w:rPr><w:del\b[^>]*\/><\/w:rPr><w:sectPr>/u);

  const [accepted, rejected] = await Promise.all([
    reopenResolved(result.value.buffer, "accept"),
    reopenResolved(result.value.buffer, "reject"),
  ]);
  expect(storyProjection(accepted)).toEqual(storyProjection(targetReviewer));
  expect(storyProjection(rejected)).toEqual(storyProjection(baseReviewer));
  return { baseReviewer, targetReviewer };
};

const expectPendingTerminals = async ({
  pendingDocument,
  acceptedDocument,
  rejectedDocument,
}: {
  pendingDocument: Document;
  acceptedDocument: Document;
  rejectedDocument: Document;
}) => {
  const [pending, accepted, rejected] = await Promise.all([
    createDocx(pendingDocument),
    createDocx(acceptedDocument),
    createDocx(rejectedDocument),
  ]);
  const [pendingZip, acceptedReviewer, rejectedReviewer] = await Promise.all([
    JSZip.loadAsync(pending),
    FolioDocxReviewer.fromBuffer(accepted),
    FolioDocxReviewer.fromBuffer(rejected),
  ]);
  const pendingXml = await pendingZip.file("word/document.xml")?.async("text");
  expect(pendingXml).toMatch(/<w:pPr><w:rPr><w:ins\b[^>]*\/><\/w:rPr><w:sectPr>/u);

  const [resolvedAccepted, resolvedRejected] = await Promise.all([
    reopenResolved(pending, "accept"),
    reopenResolved(pending, "reject"),
  ]);
  expect(storyProjection(resolvedAccepted)).toEqual(storyProjection(acceptedReviewer));
  expect(storyProjection(resolvedRejected)).toEqual(storyProjection(rejectedReviewer));
};

describe("tracked section-boundary ownership", () => {
  test("accept removes a source endpoint before an ordinary paragraph", async () => {
    await expectResolvedTerminals(
      documentWith([paragraph("00000001", "Alpha", SECTION_A), paragraph("00000002", "Beta")]),
      documentWith([paragraph("00000001", "AlphaBeta")]),
    );
  });

  test("reject removes an inserted endpoint and restores the source paragraph", async () => {
    const insertedMark = {
      kind: "ins" as const,
      info: { id: 1, author: "Section reviewer", date: OPTIONS.timestamp },
    };
    await expectPendingTerminals({
      pendingDocument: documentWith([
        { ...paragraph("00000001", "Alpha", SECTION_A), pPrMark: insertedMark },
        paragraph("00000002", "Beta", FOLLOWING_SECTION),
      ]),
      acceptedDocument: documentWith([
        paragraph("00000001", "Alpha", SECTION_A),
        paragraph("00000002", "Beta", FOLLOWING_SECTION),
      ]),
      rejectedDocument: documentWith([paragraph("00000001", "AlphaBeta", FOLLOWING_SECTION)]),
    });
  });

  test("accept removes consecutive source section endpoints and reject restores them", async () => {
    const baseDocument = documentWith([
      paragraph("00000001", "Keep"),
      paragraph("00000002", "Delete section A", SECTION_A),
      paragraph("00000003", "Delete section B", SECTION_B),
      paragraph("00000004", "Tail", FOLLOWING_SECTION),
    ]);
    const targetDocument = documentWith([
      paragraph("00000001", "Keep"),
      paragraph("00000004", "Tail", FOLLOWING_SECTION),
    ]);
    const expectedBase = [
      { paraId: "00000001", sectionProperties: null },
      { paraId: "00000002", sectionProperties: SECTION_A },
      { paraId: "00000003", sectionProperties: SECTION_B },
      { paraId: "00000004", sectionProperties: FOLLOWING_SECTION },
    ];
    const expectedTarget = [
      { paraId: "00000001", sectionProperties: null },
      { paraId: "00000004", sectionProperties: FOLLOWING_SECTION },
    ];

    const { baseReviewer, targetReviewer } = await expectResolvedTerminals(
      baseDocument,
      targetDocument,
    );
    expect(paragraphSectionProjection(baseReviewer)).toEqual(expectedBase);
    expect(paragraphSectionProjection(targetReviewer)).toEqual(expectedTarget);
  });

  test("accept-all removes one section endpoint alongside ordinary deleted paragraphs", async () => {
    await expectResolvedTerminals(
      documentWith([
        paragraph("00000001", "Keep"),
        paragraph("00000002", "Remove section", SECTION_A),
        paragraph("00000003", "Remove ordinary"),
        paragraph("00000004", "Tail", FOLLOWING_SECTION),
      ]),
      documentWith([
        paragraph("00000001", "Keep"),
        paragraph("00000004", "Tail", FOLLOWING_SECTION),
      ]),
    );
  });

  test("an overlapping resolution cannot change the state captured by an in-flight save", async () => {
    const pending = await createDocx(
      documentWith([
        {
          ...paragraph("00000001", "Alpha", SECTION_A),
          pPrMark: {
            kind: "del",
            info: { id: 1, author: OPTIONS.author, date: OPTIONS.timestamp },
          },
        },
        {
          ...paragraph("00000002", "Beta", SECTION_B),
          pPrMark: {
            kind: "del",
            info: { id: 2, author: OPTIONS.author, date: OPTIONS.timestamp },
          },
        },
        paragraph("00000003", "Tail", FOLLOWING_SECTION),
      ]),
    );
    const reviewer = await FolioDocxReviewer.fromBuffer(pending);

    expect(reviewer.acceptChange(1)).toBe(true);
    const firstSave = reviewer.toBuffer();
    expect(reviewer.acceptChange(2)).toBe(true);
    const secondSave = reviewer.toBuffer();

    const [first, second] = await Promise.all([firstSave, secondSave]);
    const [reopenedFirst, reopenedSecond] = await Promise.all([
      FolioDocxReviewer.fromBuffer(first),
      FolioDocxReviewer.fromBuffer(second),
    ]);
    expect(storyProjection(reopenedFirst)).toEqual({
      blocks: [
        { id: "00000001", kind: "paragraph", text: "AlphaBeta" },
        { id: "00000003", kind: "paragraph", text: "Tail" },
      ],
      sections: [
        { paraId: "00000001", sectionProperties: SECTION_B },
        { paraId: "00000003", sectionProperties: FOLLOWING_SECTION },
      ],
      topLevelKinds: ["paragraph", "paragraph"],
    });
    expect(reopenedFirst.getChanges()).toHaveLength(1);
    expect(storyProjection(reopenedSecond)).toEqual({
      blocks: [{ id: "00000001", kind: "paragraph", text: "AlphaBetaTail" }],
      sections: [{ paraId: "00000001", sectionProperties: FOLLOWING_SECTION }],
      topLevelKinds: ["paragraph"],
    });
    expect(reopenedSecond.getChanges()).toHaveLength(0);
  });

  test("accept removes an emptied endpoint before a table and reject restores it", async () => {
    await expectResolvedTerminals(
      documentWith([
        paragraph("00000001", "Remove", SECTION_B),
        table("Cell"),
        paragraph("00000002", "Tail", FOLLOWING_SECTION),
      ]),
      documentWith([table("Cell"), paragraph("00000002", "Tail", FOLLOWING_SECTION)]),
    );
  });
});
