import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { paragraph, run, table } from "../docx/server/build";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

type FixtureOptions = {
  kind: "ins" | "moveTo";
  following: "paragraph" | "table" | "end";
  anchors: "none" | "bookmark";
  text: string;
};

const fixture = async ({ kind, following, anchors, text }: FixtureOptions) => {
  const document = createEmptyDocument();
  const info = { author: "Earlier Reviewer", date: "2026-01-01T00:00:00Z" };
  document.package.document.content = [
    paragraph("Opening clause."),
    ...(kind === "moveTo"
      ? [
          {
            ...paragraph([
              {
                type: "moveFrom",
                info: { ...info, id: 2 },
                content: [run(text)],
              },
            ]),
            pPrMark: { kind: "moveFrom" as const, info: { ...info, id: 1 } },
          },
        ]
      : []),
    {
      ...paragraph("Middle clause.", { alignment: "right" }),
      ...(following === "end" && {
        pPrMark: { kind, info: { ...info, id: 3 } },
      }),
    },
    {
      ...paragraph(
        [
          ...(anchors === "bookmark"
            ? [{ type: "bookmarkStart" as const, id: 11, name: "Clause" }]
            : []),
          {
            type: kind === "moveTo" ? "moveTo" : "insertion",
            info: { ...info, id: 4 },
            content: [run(text)],
          },
        ],
        { alignment: "right" },
      ),
      paraId: "10000001",
      ...(following !== "end" && {
        pPrMark: { kind, info: { ...info, id: 3 } },
      }),
    },
    ...(anchors === "bookmark"
      ? [{ type: "bookmarkEnd" as const, id: 11 }]
      : []),
    ...(following === "table" ? [table({ rows: [["Cell."]] })] : []),
    ...(following === "end"
      ? []
      : [paragraph("Closing clause.", { alignment: "center" })]),
  ];
  return createDocx(document);
};

const open = (bytes: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });
const reopen = async (reviewer: FolioDocxReviewer) =>
  open(await reviewer.toBuffer());
const resolved = async (
  reviewer: FolioDocxReviewer,
  resolution: "accept" | "reject",
) => {
  const copy = await reopen(reviewer);
  if (resolution === "accept") copy.acceptAll();
  else copy.rejectAll();
  const saved = await copy.toBuffer();
  const blocks = (await open(saved)).getContent().map((block) => {
    if (block.kind === "diagnostic")
      panic("The added-break fixture contains an opaque block");
    const { text, kind, directAlignment } = block;
    return { text, kind, directAlignment };
  });
  // A paragraph that goes takes no bookmark with it: both ends stay, in place.
  const xml =
    (await (
      await JSZip.loadAsync(saved)
    )
      .file("word/document.xml")
      ?.async("string")) ?? "";
  const bookmarks = [
    ...xml.matchAll(/<w:(bookmarkStart|bookmarkEnd)\b|<w:(p|tbl)\b/gu),
  ].map((match) => match[1] ?? match[2]);
  return { blocks, bookmarks };
};

const checkDeletion = async (options: FixtureOptions) => {
  const bytes = await fixture(options);
  const baseline = await open(bytes);
  const expectedRejected = await resolved(baseline, "reject");
  const direct = await open(bytes);
  const tracked = await open(bytes);
  for (const [mode, reviewer] of [
    ["direct", direct],
    ["tracked-changes", tracked],
  ] as const) {
    const target = reviewer.getContent().find(({ id }) => id === "10000001");
    if (!target) panic("The added-break fixture lost its deletion target");
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [{ id: "delete", type: "deleteBlock", blockId: target.id }],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.applied).toHaveLength(1);
    expect(result.skipped).toEqual([]);
  }
  if (options.kind === "moveTo" && options.following === "paragraph") {
    const xml = await (
      await JSZip.loadAsync(await tracked.toBuffer())
    )
      .file("word/document.xml")
      ?.async("string");
    expect(xml).toMatch(/<w:moveTo\b[^>]*><w:del\b/u);
  }
  expect(await resolved(tracked, "accept")).toEqual(
    await resolved(direct, "accept"),
  );
  expect(await resolved(tracked, "reject")).toEqual(expectedRejected);
};

// The sequence generator previously started with ordinary revisions only.
// Exercise imported relocation breaks as well as insertion breaks on every run.
const CASES = (["ins", "moveTo"] as const).flatMap((kind) =>
  (["paragraph", "table", "end"] as const).flatMap((following) =>
    (["none", "bookmark"] as const).map((anchors) => ({
      kind,
      following,
      anchors,
    })),
  ),
);

test.each(CASES)(
  "deleting an imported $kind paragraph before $following ($anchors) resolves alike in both modes",
  async ({ kind, following, anchors }) => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom("Moved clause.", "Approved 👍🏽.", "Článek 日本語."),
        async (text) => checkDeletion({ kind, following, anchors, text }),
      ),
      {
        id: "deleting an imported $kind paragraph before $following",
        numRuns: 3,
      },
    );
  },
  propertyTestTimeout(60_000),
);
