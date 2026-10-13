import assert from "node:assert/strict";
import { test } from "node:test";
import { createDocx, paragraph } from "@stll/folio-core/server";
import { fromMarkdown } from "@stll/folio-core/markdown";
import { openReviewer } from "../support/documents.ts";
import {
  canonicalReviewModel,
  assertRevisionCoverage,
  assertReviewRoundTrip,
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

test("revision comparison preserves character order across adjacent run groups", () => {
  const fragments = [...revision.text].map((text, index) =>
    Object.assign({}, revision, { id: index + 10, text }),
  );
  assertRevisionCoverage([revision], fragments);
  const formatting = { ...revision, type: "formatting" as const };
  assertRevisionCoverage(
    [revision, formatting],
    fragments.flatMap((entry) => [entry, { ...entry, type: "formatting" as const }]),
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
  const reopened = await openReviewer(new Uint8Array(await reviewer.toBuffer()));
  await assertReviewRoundTrip({ live: reviewer, reopened, context: "round trip" });
  const changed = reopened.toDocument();
  changed.package.document.content.push(paragraph("Different"));
  const changedSaved = await openReviewer(new Uint8Array(await createDocx(changed)));
  await assert.rejects(
    assertReviewRoundTrip({
      live: reviewer,
      reopened: {
        getContent: () => reopened.getContent(),
        getChanges: () => reopened.getChanges(),
        toResolvedDocument: (decision) => changedSaved.toResolvedDocument(decision),
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
  const rejectedSaved = await openReviewer(new Uint8Array(await createDocx(rejected)));
  await assert.rejects(
    assertReviewRoundTrip({
      live: reviewer,
      reopened: {
        getContent: () => reopened.getContent(),
        getChanges: () => reopened.getChanges(),
        toResolvedDocument: (decision) => rejectedSaved.toResolvedDocument(decision),
      },
      context: "changed rejected content",
    }),
    /reject-all document differs/,
  );
});

test("round-trip comparison preserves table formatting in resolved models", async () => {
  const model = fromMarkdown("| First | Second |\n| --- | --- |\n| Alpha beta | Body |");
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
  const reopened = await openReviewer(new Uint8Array(await reviewer.toBuffer()));
  await assertReviewRoundTrip({ live: reviewer, reopened, context: "table round trip" });
  const changed = reopened.toDocument();
  const table = changed.package.document.content.find(({ type }) => type === "table");
  if (table?.type !== "table") throw new Error("Expected table");
  const cell = table.rows.at(0)?.cells.at(0);
  if (!cell) throw new Error("Expected table cell");
  cell.formatting = { ...cell.formatting, textDirection: "tbRl" };
  const changedSaved = await openReviewer(new Uint8Array(await createDocx(changed)));
  await assert.rejects(
    assertReviewRoundTrip({
      live: reviewer,
      reopened: {
        getContent: () => reopened.getContent(),
        getChanges: () => reopened.getChanges(),
        toResolvedDocument: (decision) => changedSaved.toResolvedDocument(decision),
      },
      context: "changed table formatting",
    }),
    /accept-all document differs/,
  );
});

// Coverage alone missed idempotent loss of a revision's original formatting.
test("round-trip comparison detects lossy revision formatting with unchanged coverage", async () => {
  const model = fromMarkdown("Alpha");
  const sourceParagraph = model.package.document.content.at(0);
  if (sourceParagraph?.type !== "paragraph") throw new Error("Expected paragraph");
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
  if (lossyParagraph?.type !== "paragraph") throw new Error("Expected paragraph");
  const lossyRun = lossyParagraph.content.find(({ type }) => type === "run");
  const change = lossyRun?.type === "run" ? lossyRun.propertyChanges?.at(0) : undefined;
  if (!change) throw new Error("Expected formatting revision");
  delete change.previousFormatting;
  // Simulate an exporter that drops the same field on every save. The live
  // in-memory resolution remains independent of that lossy export boundary.
  const live = {
    getContent: () => reviewer.getContent(),
    getChanges: () => reviewer.getChanges(),
    toDocument: () => lossy,
    toResolvedDocument: (decision: "accept" | "reject") => reviewer.toResolvedDocument(decision),
  };
  const reopened = await openReviewer(new Uint8Array(await createDocx(live.toDocument())));
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
    canonicalReviewModel({ sourceXml: "<w:unknown/>" }),
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
  const reopened = await openReviewer(new Uint8Array(await reviewer.toBuffer()));
  const withHeader = (decision: "accept" | "reject", text: string) => {
    const document = reviewer.toResolvedDocument(decision);
    document.package.headers = new Map([
      ["rId8", { type: "header", hdrFtrType: "default", content: [paragraph(text)] }],
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
