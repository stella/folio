/**
 * A bookmark boundary without a partner, as real documents hold one: a lone
 * `_GoBack` start between blocks, a start whose end a template dropped, an end
 * whose start is gone. Each placement (between body blocks, in a paragraph, in
 * a hyperlink, in a table cell, in a header, in a footnote) and each kind (a
 * lone start, a lone end) is opened by both reviewer engines, and the
 * canonical one keeps the boundary exactly where the source has it:
 *
 * - it opens wherever the legacy engine opens;
 * - saving, with no edit, with an edit elsewhere (direct, or tracked and then
 *   pending, accepted or rejected), and saving again after reopening, writes
 *   the boundaries of the source part, no more and no fewer;
 * - the text it saves is the legacy engine's for the same edits;
 * - an edit beside the boundary, or over the text around it, keeps or drops it
 *   as the legacy engine does.
 */

import { describe, expect, mock, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { readBookmarkBoundaryAttrs } from "../prosemirror/bookmarkBoundaryAttrs";
import * as validation from "../prosemirror/validation";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml";

type Kind = "start" | "end";

const markerXml = (kind: Kind): string =>
  kind === "start"
    ? '<w:bookmarkStart w:id="90" w:name="_GoBack"/>'
    : '<w:bookmarkEnd w:id="90"/>';

const run = (text: string): string =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** The paragraph a placement puts the marker in or beside: "Alpha beta". */
const inParagraph = (marker: string): string =>
  `<w:p>${run("Alpha")}${marker}${run(" beta")}</w:p>`;
const inHyperlink = (marker: string): string =>
  `<w:p><w:hyperlink w:anchor="target">${run("Alpha")}${marker}</w:hyperlink>${run(" beta")}</w:p>`;
const beforeParagraph = (marker: string): string =>
  `${marker}<w:p>${run("Alpha beta")}</w:p>`;
const inTableCell = (marker: string): string =>
  '<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${beforeParagraph(marker)}</w:tc></w:tr></w:tbl>`;

type Placement = {
  name: string;
  /** The part that holds the marker. */
  part: "word/document.xml" | "word/header1.xml" | "word/footnotes.xml";
  /** The body's blocks before the closing paragraph. */
  body: (marker: string) => string;
  header?: (marker: string) => string;
  footnote?: (marker: string) => string;
  /** Whether the "Alpha beta" paragraph is a body block a reviewer edits. */
  editsBeside: boolean;
};

const PLACEMENTS: readonly Placement[] = [
  {
    name: "between body blocks",
    part: "word/document.xml",
    body: beforeParagraph,
    editsBeside: true,
  },
  {
    name: "in a paragraph",
    part: "word/document.xml",
    body: inParagraph,
    editsBeside: true,
  },
  {
    name: "in a hyperlink",
    part: "word/document.xml",
    body: inHyperlink,
    editsBeside: true,
  },
  {
    name: "in a table cell",
    part: "word/document.xml",
    body: inTableCell,
    editsBeside: false,
  },
  {
    name: "in a header",
    part: "word/header1.xml",
    body: () => `<w:p>${run("Alpha beta")}</w:p>`,
    header: inHyperlink,
    editsBeside: false,
  },
  {
    name: "in a footnote",
    part: "word/footnotes.xml",
    body: () =>
      `<w:p>${run("Alpha beta")}<w:r><w:footnoteReference w:id="1"/></w:r></w:p>`,
    footnote: inHyperlink,
    editsBeside: false,
  },
];

const packageWith = async (
  placement: Placement,
  kind: Kind,
): Promise<ArrayBuffer> => {
  const marker = markerXml(kind);
  const zip = new ZipArchive();
  const overrides = [
    `<Override PartName="/word/document.xml" ContentType="${MAIN}.document.main+xml"/>`,
    placement.header
      ? `<Override PartName="/word/header1.xml" ContentType="${MAIN}.header+xml"/>`
      : "",
    placement.footnote
      ? `<Override PartName="/word/footnotes.xml" ContentType="${MAIN}.footnotes+xml"/>`
      : "",
  ].join("");
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  const relationships = [
    placement.header
      ? `<Relationship Id="rIdH" Type="${R}/header" Target="header1.xml"/>`
      : "",
    placement.footnote
      ? `<Relationship Id="rIdF" Type="${R}/footnotes" Target="footnotes.xml"/>`
      : "",
  ].join("");
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`,
  );
  const sectPr = placement.header
    ? '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>'
    : "<w:sectPr/>";
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${placement.body(marker)}<w:p>${run("Tail")}</w:p>${sectPr}</w:body></w:document>`,
  );
  if (placement.header)
    zip.set(
      "word/header1.xml",
      `${HEAD}<w:hdr xmlns:w="${W}" xmlns:r="${R}">${placement.header(marker)}</w:hdr>`,
    );
  if (placement.footnote)
    zip.set(
      "word/footnotes.xml",
      `${HEAD}<w:footnotes xmlns:w="${W}" xmlns:r="${R}"><w:footnote w:id="1">${placement.footnote(marker)}</w:footnote></w:footnotes>`,
    );
  const bytes = await zip.generate({ compression: "store" });
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
};

/** A part's bookmark boundaries in document order, and its words. */
const partOf = async (buffer: ArrayBuffer, name: Placement["part"]) => {
  const zip = await ZipArchive.read(buffer);
  const xml = (await zip.entry(name)?.text()) ?? "";
  return {
    markers: [
      ...xml.matchAll(/<w:bookmark(Start|End)\b[^>]*\bw:id="(\d+)"/gu),
    ].map(([, kind, id]) => `${kind}:${id}`),
    text: [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)</gu)]
      .map(([, text]) => text)
      .join(""),
  };
};

const open = (buffer: ArrayBuffer, session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Reviewer",
    session,
  });

const blockIdWithText = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer
    .snapshot()
    .blocks.find((candidate) => candidate.text === text);
  if (block === undefined) throw new Error(`no block reads "${text}"`);
  return block.id;
};

type Mode = "direct" | "tracked-changes";
type Resolution = "pending" | "accepted" | "rejected";

/** Whether the reviewer applied the one operation; a refusal is `false`. */
const applies = (
  reviewer: FolioDocxReviewer,
  operation: FolioDocumentOperation,
  mode: Mode,
): boolean => {
  try {
    return (
      reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          operations: [operation],
          mode,
        },
        { undefinedReferences: "refuse" },
      ).applied.length === 1
    );
  } catch {
    return false;
  }
};

type Flow = {
  find: string;
  replace: string;
  /** The text of the block the edit targets. */
  block: string;
  mode: Mode;
  resolution: Resolution;
};

/** The flow on one engine: whether it applied, and the saved part. */
const runFlow = async (
  buffer: ArrayBuffer,
  session: HeadlessReviewerSession,
  placement: Placement,
  flow: Flow | null,
) => {
  const reviewer = await open(buffer, session);
  let applied = true;
  if (flow !== null) {
    applied = applies(
      reviewer,
      {
        id: "edit",
        type: "replaceInBlock",
        blockId: blockIdWithText(reviewer, flow.block),
        find: flow.find,
        replace: flow.replace,
      },
      flow.mode,
    );
    if (flow.resolution === "accepted") reviewer.acceptAll();
    if (flow.resolution === "rejected") reviewer.rejectAll();
  }
  const saved = await reviewer.toBuffer();
  return { applied, saved, part: await partOf(saved, placement.part) };
};

const elsewhere = (mode: Mode, resolution: Resolution): Flow => ({
  mode,
  resolution,
  block: "Tail",
  find: "Tail",
  replace: "End",
});

const FLOWS_ELSEWHERE: readonly Flow[] = [
  elsewhere("direct", "pending"),
  elsewhere("tracked-changes", "pending"),
  elsewhere("tracked-changes", "accepted"),
  elsewhere("tracked-changes", "rejected"),
];

const beside = (mode: Mode, find: string, replace: string): Flow => ({
  mode,
  resolution: "pending",
  block: "Alpha beta",
  find,
  replace,
});

const FLOWS_BESIDE: readonly Flow[] = [
  beside("direct", "beta", "gamma"),
  beside("direct", "Alpha beta", "Omega"),
  beside("tracked-changes", "beta", "gamma"),
  beside("tracked-changes", "Alpha beta", "Omega"),
];

const KINDS: readonly Kind[] = ["start", "end"];

const loneMarker = (kind: Kind): string =>
  kind === "start" ? "Start:90" : "End:90";

describe("a bookmark boundary without a partner", () => {
  for (const placement of PLACEMENTS)
    for (const kind of KINDS) {
      test(`a lone ${kind} ${placement.name} opens, saves and reopens as the source has it`, async () => {
        const buffer = await packageWith(placement, kind);
        const source = await partOf(buffer, placement.part);
        expect(source.markers).toEqual([loneMarker(kind)]);
        // Both engines open it: the canonical one wherever the legacy one does.
        await open(buffer, "legacy");
        await open(buffer, "canonical");

        for (const flow of [null, ...FLOWS_ELSEWHERE]) {
          const legacy = await runFlow(buffer, "legacy", placement, flow);
          const canonical = await runFlow(buffer, "canonical", placement, flow);
          expect(canonical.applied).toBe(legacy.applied);
          expect(canonical.part.markers).toEqual(source.markers);
          expect(canonical.part.text).toBe(legacy.part.text);
          // Reopened and saved again, the part is a fixed point.
          const reopened = await runFlow(
            canonical.saved,
            "canonical",
            placement,
            null,
          );
          expect(reopened.part).toEqual(canonical.part);
        }
      }, 60_000);

      if (!placement.editsBeside) continue;
      test(`an edit beside a lone ${kind} ${placement.name} keeps it as the legacy engine does`, async () => {
        const buffer = await packageWith(placement, kind);
        for (const flow of FLOWS_BESIDE) {
          const legacy = await runFlow(buffer, "legacy", placement, flow);
          const canonical = await runFlow(buffer, "canonical", placement, flow);
          expect(canonical.applied).toBe(legacy.applied);
          expect(canonical.part.text).toBe(legacy.part.text);
          // The legacy engine writes a paragraph's lone start back with an end
          // of its own; whether the source's boundary survived is the same.
          expect(canonical.part.markers.includes(loneMarker(kind))).toBe(
            legacy.part.markers.includes(loneMarker(kind)),
          );
          // What the canonical engine keeps, it keeps as the source has it.
          expect(
            canonical.part.markers.every(
              (marker) => marker === loneMarker(kind),
            ),
          ).toBe(true);
        }
      }, 60_000);
    }
});

/** Every boundary node id with one start and one end, as the check used to ask. */
const everyBoundaryPaired = (doc: PMNode): boolean => {
  const counts = new Map<number, number>();
  let paired = true;
  doc.descendants((node) => {
    if (
      node.type.name !== "bookmarkBoundary" &&
      node.type.name !== "blockBookmarkBoundary"
    )
      return true;
    const attrs = readBookmarkBoundaryAttrs(node);
    if (!attrs.ok) return false;
    counts.set(attrs.value.id, (counts.get(attrs.value.id) ?? 0) + 1);
    return false;
  });
  for (const count of counts.values()) if (count !== 2) paired = false;
  return paired;
};

test("the open law catches a projection check that refuses a lone boundary", async () => {
  // The test-only break: the projection check refuses a boundary node whose
  // partner is missing, as it did before such a boundary was source content.
  const original = { ...validation };
  const strict = (doc: PMNode, context: string): void => {
    if (!everyBoundaryPaired(doc))
      throw new validation.ProseMirrorDocumentValidationError(context, [
        { path: "doc", message: "A bookmark boundary has no partner." },
      ]);
  };
  mock.module("../prosemirror/validation", () => ({
    ...original,
    assertValidProseMirrorDocument: (doc: PMNode, context: string) => {
      strict(doc, context);
      original.assertValidProseMirrorDocument(doc, context);
    },
    assertValidProseMirrorDocumentKeeping: (
      doc: PMNode,
      context: string,
      previous: Parameters<
        typeof original.assertValidProseMirrorDocumentKeeping
      >[2],
    ) => {
      strict(doc, context);
      return original.assertValidProseMirrorDocumentKeeping(
        doc,
        context,
        previous,
      );
    },
  }));
  const hyperlink = PLACEMENTS.find(({ name }) => name === "in a hyperlink");
  if (hyperlink === undefined) throw new Error("no hyperlink placement");
  const buffer = await packageWith(hyperlink, "start");
  try {
    await expect(open(buffer, "canonical")).rejects.toThrow();
  } finally {
    mock.module("../prosemirror/validation", () => original);
  }
  await open(buffer, "canonical");
}, 60_000);
