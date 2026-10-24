import { panic } from "better-result";
import JSZip from "jszip";
import assert from "node:assert/strict";
import { test } from "node:test";

import type {
  ComplexField,
  Document,
  HeaderFooter,
  ParagraphContent,
  Run,
  SimpleField,
} from "@stll/docx-core/model";
import { fromMarkdown } from "@stll/folio-core/markdown";
import { createDocx, paragraph } from "@stll/folio-core/server";

import { openReviewer } from "../support/documents.ts";
import {
  assertReviewRoundTrip,
  assertRevisionCoverage,
  canonicalResolvedReviewModel,
  canonicalReviewModel,
  revisionCoverage,
} from "../support/review-comparison.ts";

type Change = Parameters<typeof revisionCoverage>[0][number];
const revision = {
  id: 7,
  type: "deletion",
  author: "Reviewer",
  date: "2026-01-01T00:00:00Z",
  text: "A👍🏽B",
  blockId: "412DDC7F",
} as const satisfies Change;

const reviewDocumentWithContent = (content: ParagraphContent[]): Document => {
  const document = fromMarkdown("Alpha beta");
  const first = document.package.document.content.at(0);
  if (first?.type !== "paragraph") panic("Expected paragraph");
  first.content = content;
  return document;
};

const textRun = (text: string, fields: Partial<Run> = {}): Run => ({
  type: "run",
  ...fields,
  content: [{ type: "text", text }],
});

const documentWithSectionStories = (
  headerContent: ParagraphContent[],
  footerContent: ParagraphContent[],
): Document => {
  const document = reviewDocumentWithContent([textRun("Body")]);
  const packageHeader: HeaderFooter = {
    type: "header",
    hdrFtrType: "default",
    content: [paragraph("Package header")],
  };
  const packageFooter: HeaderFooter = {
    type: "footer",
    hdrFtrType: "default",
    content: [paragraph("Package footer")],
  };
  const sectionHeader: HeaderFooter = {
    type: "header",
    hdrFtrType: "default",
    content: [
      {
        type: "paragraph",
        content: headerContent,
      },
    ],
  };
  const sectionFooter: HeaderFooter = {
    type: "footer",
    hdrFtrType: "default",
    content: [
      {
        type: "paragraph",
        content: footerContent,
      },
    ],
  };
  document.package.headers = new Map([["rIdHeader", packageHeader]]);
  document.package.footers = new Map([["rIdFooter", packageFooter]]);
  document.package.document.sections = [
    {
      properties: {
        headerReferences: [{ type: "default", rId: "rIdHeader" }],
        footerReferences: [{ type: "default", rId: "rIdFooter" }],
      },
      content: [],
      headers: new Map([["default", sectionHeader]]),
      footers: new Map([["default", sectionFooter]]),
    },
  ];
  return document;
};

test("model comparison normalizes only safely mergeable adjacent runs", () => {
  const merged = reviewDocumentWithContent([textRun("Alpha beta")]);
  const segmented = reviewDocumentWithContent([
    textRun("Alpha "),
    textRun("beta"),
  ]);

  assert.notDeepEqual(
    canonicalReviewModel(segmented),
    canonicalReviewModel(merged),
  );
  assert.deepEqual(
    canonicalResolvedReviewModel(segmented),
    canonicalResolvedReviewModel(merged),
  );

  const controls: ParagraphContent[][] = [
    [textRun("Alpha "), textRun("gamma")],
    [textRun("Alpha "), textRun("beta", { formatting: { bold: true } })],
    [
      textRun("Alpha "),
      textRun("beta", {
        propertyChanges: [
          {
            type: "runPropertyChange",
            info: { id: 8, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
            previousFormatting: { italic: true },
          },
        ],
      }),
    ],
    [
      textRun("Alpha "),
      textRun("beta", {
        preservedAttributes: [
          {
            name: "rsidR",
            namespace:
              "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            value: "00ABCDEF",
          },
        ],
      }),
    ],
    [textRun("Alpha "), textRun(""), textRun("beta")],
    [
      textRun("Alpha "),
      textRun("beta", {
        formatting: {
          preserved: { children: [{ index: 1, xml: "<w:spacing/>" }] },
        },
      }),
    ],
    [
      textRun("Alpha "),
      { type: "run", content: [{ type: "tab" }] },
      textRun("beta"),
    ],
    [
      textRun("Alpha "),
      { type: "bookmarkStart", id: 9, name: "range" },
      textRun("beta"),
    ],
    [
      textRun("Alpha "),
      { type: "bookmarkEnd", id: 9, copyOriginAncestors: [0] },
      textRun("beta"),
    ],
    [
      textRun("Alpha "),
      {
        type: "moveToRangeStart",
        id: 9,
        name: "move",
        author: "Reviewer",
      },
      textRun("beta"),
    ],
    [
      textRun("Alpha "),
      {
        type: "insertion",
        info: { id: 8, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
        content: [textRun("beta")],
      },
    ],
  ];

  for (const content of controls) {
    assert.notDeepEqual(
      canonicalResolvedReviewModel(reviewDocumentWithContent(content)),
      canonicalResolvedReviewModel(merged),
    );
  }
});

test("resolved projection visits section-local headers and footers without relaxing run formatting or metadata", () => {
  const segmented = documentWithSectionStories(
    [textRun("Section "), textRun("header")],
    [textRun("Section "), textRun("footer")],
  );
  const merged = documentWithSectionStories(
    [textRun("Section header")],
    [textRun("Section footer")],
  );

  assert.notDeepEqual(
    canonicalReviewModel(segmented),
    canonicalReviewModel(merged),
  );
  assert.deepEqual(
    canonicalResolvedReviewModel(segmented),
    canonicalResolvedReviewModel(merged),
  );

  const formatted = documentWithSectionStories(
    [textRun("Section "), textRun("header", { formatting: { bold: true } })],
    [textRun("Section footer")],
  );
  assert.notDeepEqual(
    canonicalResolvedReviewModel(formatted),
    canonicalResolvedReviewModel(merged),
  );

  const metadata = documentWithSectionStories(
    [textRun("Section header")],
    [
      textRun("Section "),
      textRun("footer", {
        preservedAttributes: [
          {
            name: "rsidR",
            namespace:
              "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            value: "00ABCDEF",
          },
        ],
      }),
    ],
  );
  assert.notDeepEqual(
    canonicalResolvedReviewModel(metadata),
    canonicalResolvedReviewModel(merged),
  );
});

test("revision comparison preserves character order across adjacent run groups", () => {
  const fragments = [...revision.text].map((text, index) =>
    Object.assign({}, revision, { id: index + 10, text }),
  );
  assertRevisionCoverage([revision], fragments);
  const formatting = { ...revision, type: "formatting" as const };
  assertRevisionCoverage(
    [revision, formatting],
    fragments.flatMap((entry) => [
      entry,
      { ...entry, type: "formatting" as const },
    ]),
  );
  for (const changed of [
    fragments.slice(1),
    [...fragments].reverse(),
    fragments.map((entry, index) =>
      index === 1 ? Object.assign({}, entry, { text: "X" }) : entry,
    ),
    [{ ...revision, type: "insertion" as const }],
    [{ ...revision, author: "Other" }],
    [{ ...revision, date: "2026-01-02T00:00:00Z" }],
    [{ ...revision, blockId: "31617F9A" }],
  ])
    assert.throws(() => assertRevisionCoverage([revision], changed));
  assert.throws(() => assertRevisionCoverage([{ ...revision, text: "" }], []));
  assert.throws(() =>
    assertRevisionCoverage(
      [
        { ...revision, text: "" },
        { ...revision, text: "" },
      ],
      [{ ...revision, text: "" }],
    ),
  );
});

test("revision comparison changes only explicitly paired positional identities", () => {
  const alias = "\u0000position:0";
  const actual = { ...revision, blockId: "31617F9A" };
  assert.deepEqual(
    revisionCoverage([revision], new Map([[revision.blockId, alias]])),
    revisionCoverage([actual], new Map([[actual.blockId, alias]])),
  );
  assert.notDeepEqual(revisionCoverage([revision]), revisionCoverage([actual]));
});

test("round-trip comparison checks both full resolution documents", async () => {
  const model = fromMarkdown("Alpha beta");
  const reviewer = await openReviewer(new Uint8Array(await createDocx(model)));
  const block = reviewer.getContent().at(0);
  if (!block) throw new Error("Expected paragraph");
  reviewer.applyDocumentOperations(
    {
      version: 1,
      mode: "tracked-changes",
      operations: [
        {
          id: "replace",
          type: "replaceInBlock",
          blockId: block.id,
          find: "beta",
          replace: "gamma",
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  const reopened = await openReviewer(
    new Uint8Array(await reviewer.toBuffer()),
  );
  await assertReviewRoundTrip({
    live: reviewer,
    reopened,
    context: "round trip",
  });
  const changed = reopened.toDocument();
  changed.package.document.content.push(paragraph("Different"));
  const changedSaved = await openReviewer(
    new Uint8Array(await createDocx(changed)),
  );
  await assert.rejects(
    assertReviewRoundTrip({
      live: reviewer,
      reopened: {
        getContent: () => reopened.getContent(),
        getChanges: () => reopened.getChanges(),
        toResolvedDocument: (decision) =>
          changedSaved.toResolvedDocument(decision),
      },
      context: "changed resolution",
    }),
    /accept-all document differs/,
  );
  const rejected = reopened.toDocument();
  const first = rejected.package.document.content.at(0);
  if (first?.type !== "paragraph") throw new Error("Expected paragraph");
  const deletion = first.content.find(({ type }) => type === "deletion");
  if (deletion?.type !== "deletion") throw new Error("Expected deleted text");
  let altered = 0;
  for (const entry of deletion.content) {
    if (entry.type !== "run") continue;
    for (const content of entry.content) {
      if (content.type !== "text") continue;
      content.text = `${content.text} changed`;
      altered++;
    }
  }
  assert.ok(altered > 0);
  const rejectedSaved = await openReviewer(
    new Uint8Array(await createDocx(rejected)),
  );
  await assert.rejects(
    assertReviewRoundTrip({
      live: reviewer,
      reopened: {
        getContent: () => reopened.getContent(),
        getChanges: () => reopened.getChanges(),
        toResolvedDocument: (decision) =>
          rejectedSaved.toResolvedDocument(decision),
      },
      context: "changed rejected content",
    }),
    /reject-all document differs/,
  );
});

test("round-trip comparison preserves table formatting in resolved models", async () => {
  const model = fromMarkdown(
    "| First | Second |\n| --- | --- |\n| Alpha beta | Body |",
  );
  const reviewer = await openReviewer(new Uint8Array(await createDocx(model)));
  const block = reviewer.getContent().find(({ text }) => text.includes("beta"));
  if (!block) throw new Error("Expected table cell paragraph");
  reviewer.applyDocumentOperations(
    {
      version: 1,
      mode: "tracked-changes",
      operations: [
        {
          id: "replace",
          type: "replaceInBlock",
          blockId: block.id,
          find: "beta",
          replace: "gamma",
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  const reopened = await openReviewer(
    new Uint8Array(await reviewer.toBuffer()),
  );
  await assertReviewRoundTrip({
    live: reviewer,
    reopened,
    context: "table round trip",
  });
  const changed = reopened.toDocument();
  const table = changed.package.document.content.find(
    ({ type }) => type === "table",
  );
  if (table?.type !== "table") throw new Error("Expected table");
  const cell = table.rows.at(0)?.cells.at(0);
  if (!cell) throw new Error("Expected table cell");
  cell.formatting = { ...cell.formatting, textDirection: "tbRl" };
  const changedSaved = await openReviewer(
    new Uint8Array(await createDocx(changed)),
  );
  await assert.rejects(
    assertReviewRoundTrip({
      live: reviewer,
      reopened: {
        getContent: () => reopened.getContent(),
        getChanges: () => reopened.getChanges(),
        toResolvedDocument: (decision) =>
          changedSaved.toResolvedDocument(decision),
      },
      context: "changed table formatting",
    }),
    /accept-all document differs/,
  );
});

test("resolved projection preserves source metadata checks without mutation", () => {
  const model = reviewDocumentWithContent([textRun("Alpha "), textRun("beta")]);
  const sourceParagraph = model.package.document.content.at(0);
  if (sourceParagraph?.type !== "paragraph") panic("Expected paragraph");
  const originalProjection = canonicalReviewModel(model);
  const firstProjection = canonicalResolvedReviewModel(model);

  assert.deepEqual(canonicalResolvedReviewModel(model), firstProjection);
  assert.deepEqual(canonicalReviewModel(model), originalProjection);

  Object.defineProperty(sourceParagraph, Symbol("unclassifiedCapture"), {
    enumerable: true,
    value: {},
  });
  assert.throws(
    () => canonicalResolvedReviewModel(model),
    /Unclassified model capture symbol/,
  );

  const unknownFieldModel = reviewDocumentWithContent([textRun("Alpha")]);
  const unknownFieldParagraph =
    unknownFieldModel.package.document.content.at(0);
  if (unknownFieldParagraph?.type !== "paragraph") panic("Expected paragraph");
  Object.defineProperty(unknownFieldParagraph, "sourceCaptureNew", {
    enumerable: true,
    value: {},
  });
  assert.throws(
    () => canonicalResolvedReviewModel(unknownFieldModel),
    /Unclassified model capture field/,
  );

  const sourceHandleModel = reviewDocumentWithContent([textRun("Alpha")]);
  const sourceHandleParagraph =
    sourceHandleModel.package.document.content.at(0);
  if (sourceHandleParagraph?.type !== "paragraph") panic("Expected paragraph");
  Object.defineProperty(
    sourceHandleParagraph,
    Symbol("paragraphPropertyCapture"),
    { enumerable: true, value: {} },
  );
  assert.deepEqual(
    canonicalResolvedReviewModel(sourceHandleModel),
    canonicalResolvedReviewModel(reviewDocumentWithContent([textRun("Alpha")])),
  );
});

test("resolved projection rejects metadata on runs removed by consolidation", () => {
  for (const enumerable of [true, false]) {
    const model = reviewDocumentWithContent([
      textRun("Alpha "),
      textRun("beta"),
    ]);
    const sourceParagraph = model.package.document.content.at(0);
    if (sourceParagraph?.type !== "paragraph") panic("Expected paragraph");
    const secondRun = sourceParagraph.content.at(1);
    if (secondRun?.type !== "run") panic("Expected second run");
    Object.defineProperty(secondRun, Symbol("unclassifiedCapture"), {
      enumerable,
      value: {},
    });
    assert.throws(
      () => canonicalResolvedReviewModel(model),
      /Unclassified model capture symbol/,
    );
  }

  const model = reviewDocumentWithContent([textRun("Alpha "), textRun("beta")]);
  const remainingParagraph = model.package.document.content.at(0);
  if (remainingParagraph?.type !== "paragraph") panic("Expected paragraph");
  const remainingSecondRun = remainingParagraph.content.at(1);
  if (remainingSecondRun?.type !== "run") panic("Expected second run");
  Object.defineProperty(remainingSecondRun, "sourceCaptureNew", {
    enumerable: true,
    value: {},
  });
  assert.throws(
    () => canonicalResolvedReviewModel(model),
    /Unclassified model capture field/,
  );
});

// Coverage alone missed idempotent loss of a revision's original formatting.
test("round-trip comparison detects lossy revision formatting with unchanged coverage", async () => {
  const model = fromMarkdown("Alpha");
  const sourceParagraph = model.package.document.content.at(0);
  if (sourceParagraph?.type !== "paragraph")
    throw new Error("Expected paragraph");
  const run = sourceParagraph.content.find(({ type }) => type === "run");
  if (run?.type !== "run") throw new Error("Expected run");
  run.formatting = { bold: true };
  run.propertyChanges = [
    {
      type: "runPropertyChange",
      info: { id: 7, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
      previousFormatting: { italic: true },
    },
  ];
  const reviewer = await openReviewer(new Uint8Array(await createDocx(model)));
  const original = structuredClone(reviewer.toDocument());
  const lossy = structuredClone(reviewer.toDocument());
  const lossyParagraph = lossy.package.document.content.at(0);
  if (lossyParagraph?.type !== "paragraph")
    throw new Error("Expected paragraph");
  const lossyRun = lossyParagraph.content.find(({ type }) => type === "run");
  const change =
    lossyRun?.type === "run" ? lossyRun.propertyChanges?.at(0) : undefined;
  if (!change) throw new Error("Expected formatting revision");
  delete change.previousFormatting;
  // Simulate an exporter that drops the same field on every save. The live
  // in-memory resolution remains independent of that lossy export boundary.
  const live = {
    getContent: () => reviewer.getContent(),
    getChanges: () => reviewer.getChanges(),
    toDocument: () => lossy,
    toResolvedDocument: (decision: "accept" | "reject") =>
      reviewer.toResolvedDocument(decision),
  };
  const reopened = await openReviewer(
    new Uint8Array(await createDocx(live.toDocument())),
  );
  assertRevisionCoverage(live.getChanges(), reopened.getChanges());
  await assert.rejects(
    assertReviewRoundTrip({ live, reopened, context: "lossy formatting save" }),
    /reject-all document differs/,
  );
  assert.deepEqual(structuredClone(reviewer.toDocument()), original);
});

test("model projection classifies capture metadata and recurses through collections", () => {
  const text = { content: [{ type: "text", text: "Header" }] };
  const captured = {
    ...text,
    verbatimXml: "<w:p/>",
    namespaceDeclarations: { w: "uri" },
    [Symbol("paragraphPropertyCapture")]: {},
  };
  assert.deepEqual(
    canonicalReviewModel(new Map([["header", new Set([captured])]])),
    new Map([["header", new Set([text])]]),
  );
  assert.throws(
    () => canonicalReviewModel({ sourceCaptureNew: {} }),
    /Unclassified model capture field/,
  );
  assert.throws(
    () => canonicalReviewModel({ [Symbol("newCapture")]: {} }),
    /Unclassified model capture symbol/,
  );
  assert.notDeepEqual(
    canonicalReviewModel({ rawImageXml: "<w:unknown/>" }),
    canonicalReviewModel({}),
  );
});

test("round-trip comparison detects changed header text with unchanged revision coverage", async () => {
  const model = fromMarkdown("Alpha");
  const body = model.package.document.content.at(0);
  const run = body?.type === "paragraph" ? body.content.at(0) : undefined;
  if (run?.type !== "run") throw new Error("Expected run");
  run.formatting = { bold: true };
  run.propertyChanges = [
    {
      type: "runPropertyChange",
      info: { id: 7, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
      previousFormatting: { italic: true },
    },
  ];
  const reviewer = await openReviewer(new Uint8Array(await createDocx(model)));
  const reopened = await openReviewer(
    new Uint8Array(await reviewer.toBuffer()),
  );
  const withHeader = (decision: "accept" | "reject", text: string) => {
    const document = reviewer.toResolvedDocument(decision);
    document.package.headers = new Map([
      [
        "rId8",
        { type: "header", hdrFtrType: "default", content: [paragraph(text)] },
      ],
    ]);
    return document;
  };
  await assert.rejects(
    assertReviewRoundTrip({
      live: {
        getContent: () => reviewer.getContent(),
        getChanges: () => reviewer.getChanges(),
        toResolvedDocument: (decision) => withHeader(decision, "Header"),
      },
      reopened: {
        getContent: () => reopened.getContent(),
        getChanges: () => reopened.getChanges(),
        toResolvedDocument: (decision) => withHeader(decision, "Changed"),
      },
      context: "changed header",
    }),
    /accept-all document differs/,
  );
});

test("model projection keeps authored identities and revision dates strict", () => {
  const absent = [{ type: "paragraph", content: [] }];
  const generated = [{ type: "paragraph", paraId: "412DDC7F", content: [] }];
  assert.deepEqual(
    canonicalReviewModel(absent),
    canonicalReviewModel(
      generated,
      new Map([["412DDC7F", "\u0000position:0"]]),
    ),
  );
  assert.notDeepEqual(
    canonicalReviewModel(absent),
    canonicalReviewModel(generated),
  );
  assert.notDeepEqual(
    canonicalReviewModel({ info: { date: "2026-01-01" } }),
    canonicalReviewModel({ info: { date: "2026-01-02" } }),
  );
});

test("capture-only formatting holders disappear while authored properties stay strict", () => {
  for (const key of ["formatting", "previousFormatting", "currentFormatting"]) {
    const captures = { sourceXml: "<w:tcPr/>", gridSourceXml: "<w:tblGrid/>" };
    assert.deepEqual(
      canonicalReviewModel({ [key]: captures }),
      canonicalReviewModel({}),
    );
    assert.throws(() =>
      assert.deepEqual(
        canonicalReviewModel({ [key]: { ...captures, bold: true } }),
        canonicalReviewModel({ [key]: { ...captures, bold: false } }),
      ),
    );
    assert.notDeepEqual(
      canonicalReviewModel({ [key]: {} }),
      canonicalReviewModel({}),
    );
    assert.notDeepEqual(
      canonicalReviewModel({ [key]: { ...captures, bold: true } }),
      canonicalReviewModel({}),
    );
    assert.throws(
      () => canonicalReviewModel({ [key]: { sourceCaptureNew: {} } }),
      /Unclassified model capture field/,
    );
  }
});

test("field model comparison preserves meaning across simple and complex forms", () => {
  const simple = {
    type: "simpleField",
    instruction: " DATE ",
    fieldType: "DATE",
    fldLock: false,
    dirty: true,
    content: [
      {
        type: "run",
        formatting: { bold: true },
        content: [{ type: "text", text: "1 May 2026" }],
      },
    ],
  } satisfies SimpleField;
  const complex = {
    type: "complexField",
    instruction: "DATE",
    fieldType: "DATE",
    fldLock: false,
    dirty: true,
    formatting: { bold: true },
    fieldCode: [
      {
        type: "run",
        formatting: { bold: true },
        content: [{ type: "instrText", text: "DATE" }],
      },
    ],
    fieldResult: simple.content,
  } satisfies ComplexField;
  const expected = canonicalReviewModel(simple);
  assert.deepEqual(canonicalReviewModel(complex), expected);
  for (const changed of [
    { ...complex, instruction: "PAGE" },
    { ...complex, fldLock: true },
    { ...complex, dirty: false },
    { ...complex, fieldResult: [] },
    {
      ...complex,
      fieldCode: [
        { type: "run", content: [{ type: "instrText", text: "PAGE" }] },
      ],
    },
    {
      ...complex,
      fieldResult: [
        {
          type: "run",
          content: [{ type: "text", text: "1 May 2026" }],
          propertyChanges: [
            {
              type: "runPropertyChange",
              info: { id: 9, author: "Reviewer" },
              previousFormatting: { italic: true },
            },
          ],
        },
      ],
    },
  ] satisfies ComplexField[])
    assert.notDeepEqual(canonicalReviewModel(changed), expected);
});

test("a simple DATE field in a deleted paragraph keeps its resolved model and untouched XML", async () => {
  const model = fromMarkdown("Field paragraph");
  const first = model.package.document.content.at(0);
  if (first?.type !== "paragraph") throw new Error("Expected field paragraph");
  first.content = [
    {
      type: "simpleField",
      instruction: " DATE ",
      fieldType: "DATE",
      fldLock: true,
      dirty: false,
      content: [
        {
          type: "run",
          formatting: { italic: true },
          content: [{ type: "text", text: "1 May 2026" }],
        },
      ],
    },
  ];
  const bytes = new Uint8Array(await createDocx(model));
  const reviewer = await openReviewer(bytes);
  const originalXml = await (
    await JSZip.loadAsync(bytes)
  )
    .file("word/document.xml")
    ?.async("string");
  assert.ok(originalXml?.includes("fldSimple"));
  const unchangedXml = await (
    await JSZip.loadAsync(await reviewer.toBuffer())
  )
    .file("word/document.xml")
    ?.async("string");
  assert.equal(unchangedXml, originalXml);
  const block = reviewer.getContent().at(0);
  if (!block) throw new Error("Expected field block");
  const result = reviewer.applyDocumentOperations(
    {
      version: 1,
      mode: "tracked-changes",
      operations: [{ id: "delete", type: "deleteBlock", blockId: block.id }],
    },
    { undefinedReferences: "refuse" },
  );
  assert.equal(result.applied.length, 1);
  assert.deepEqual(result.skipped, []);
  const reopened = await openReviewer(
    new Uint8Array(await reviewer.toBuffer()),
  );
  await assertReviewRoundTrip({
    live: reviewer,
    reopened,
    context: "tracked DATE field",
  });
});
