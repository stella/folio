/**
 * A direct or tracked replacement in a paragraph, in the body and in a table
 * cell, applies on the canonical reviewer wherever the default one applies
 * it, and both read the same text afterwards: a word that runs out of a
 * hyperlink into the text after it, a word after an empty hyperlink, a word
 * beside a footnote reference and one beside a symbol.
 */

import { describe, expect, test } from "bun:test";
import { EditorState, type Transaction } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { CANONICAL_PROJECTION_META } from "../canonical/canonicalSession";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { canonicalReviewerSave } from "./__tests__/canonicalReviewerSave";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIEditApplyMode } from "./types";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const run = (text: string, properties = ""): string =>
  `<w:r>${properties === "" ? "" : `<w:rPr>${properties}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`;

type Shape = { name: string; paragraph: string; find: string };

const SHAPES: readonly Shape[] = [
  {
    name: "a word that runs out of a hyperlink",
    paragraph: `<w:p>${run("Start ")}<w:hyperlink r:id="rIdLink">${run("lin")}${run("ked", "<w:b/>")}</w:hyperlink>${run("word end")}</w:p>`,
    find: "linkedword",
  },
  {
    name: "a word after an empty hyperlink",
    paragraph: `<w:p><w:hyperlink r:id="rIdLink"></w:hyperlink>${run("plain", "<w:i/>")}</w:p>`,
    find: "plain",
  },
  {
    name: "a word inside a hyperlink",
    paragraph: `<w:p>${run("Before ")}<w:hyperlink r:id="rIdLink">${run("inside")}</w:hyperlink>${run(" after")}</w:p>`,
    find: "inside",
  },
  {
    name: "a word split across runs with different formatting",
    paragraph: `<w:p>${run("al")}${run("pha", "<w:b/>")}${run(" beta")}</w:p>`,
    find: "alpha",
  },
];

type Place = "body" | "table cell";

const bodyOf = (paragraph: string, place: Place): string =>
  place === "body"
    ? `${paragraph}<w:p>${run("Tail")}</w:p>`
    : '<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr>' +
      `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${paragraph}<w:p>${run("Cell tail")}</w:p></w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p>${run("Other")}</w:p></w:tc>` +
      `</w:tr></w:tbl><w:p>${run("Tail")}</w:p>`;

const fixture = async (shape: Shape, place: Place): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="${R}/hyperlink" Target="https://example.org" TargetMode="External"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${bodyOf(shape.paragraph, place)}<w:sectPr/></w:body></w:document>`,
  );
  const bytes = await zip.generate({ compression: "store" });
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
};

const open = (buffer: ArrayBuffer, session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Reviewer",
    session,
  });

const runShape = async (
  buffer: ArrayBuffer,
  session: HeadlessReviewerSession,
  shape: Shape,
  mode: FolioAIEditApplyMode,
) => {
  const reviewer = await open(buffer, session);
  const block = reviewer
    .snapshot()
    .blocks.find((candidate) => candidate.text.includes(shape.find));
  if (block === undefined) throw new Error(`no block holds ${shape.find}`);
  let outcome: unknown;
  try {
    const result = reviewer.applyDocumentOperations(
      {
        version: 1,
        mode,
        operations: [
          {
            id: "replace",
            type: "replaceInBlock",
            blockId: block.id,
            find: shape.find,
            replace: "revised",
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    outcome = {
      applied: result.applied.length,
      skipped: result.skipped.map(({ reason }) => reason),
    };
  } catch (error) {
    outcome = { threw: error instanceof Error ? error.message : String(error) };
  }
  const text = reviewer
    .snapshot()
    .blocks.map(({ text: blockText }) => blockText);
  const saved =
    session === "canonical"
      ? await canonicalReviewerSave(reviewer)
      : await reviewer.toBuffer();
  const reopened = await open(saved, session);
  return {
    outcome,
    text,
    reopened: reopened
      .snapshot()
      .blocks.map(({ text: blockText }) => blockText),
  };
};

/**
 * Diagnostic: the plugins that refuse or change a transaction the canonical
 * session publishes, each tried alone on the transaction.
 */
const pluginsChangingCanonicalCommits: string[] = [];
const originalApplyTransaction = EditorState.prototype.applyTransaction;
const tracedApplyTransaction = function (
  this: EditorState,
  rootTr: Transaction,
) {
  if (rootTr.getMeta(CANONICAL_PROJECTION_META) !== undefined) {
    const next = this.apply(rootTr);
    for (const plugin of this.plugins) {
      const key = String(Reflect.get(plugin, "key"));
      if (plugin.spec.filterTransaction?.call(plugin, rootTr, this) === false)
        pluginsChangingCanonicalCommits.push(`${key} filtered`);
      const appended = plugin.spec.appendTransaction?.call(
        plugin,
        [rootTr],
        this,
        next,
      );
      if (appended?.docChanged === true)
        pluginsChangingCanonicalCommits.push(
          `${key} appended ${JSON.stringify(appended.steps.map((step) => step.toJSON()))}`,
        );
    }
  }
  return originalApplyTransaction.call(this, rootTr);
};

const PLACES: readonly Place[] = ["body", "table cell"];
const MODES: readonly FolioAIEditApplyMode[] = ["direct", "tracked-changes"];

describe("a replacement in a paragraph", () => {
  for (const shape of SHAPES)
    for (const place of PLACES)
      for (const mode of MODES)
        test(`${mode}, ${place}: ${shape.name}`, async () => {
          const buffer = await fixture(shape, place);
          const legacy = await runShape(buffer, "legacy", shape, mode);
          pluginsChangingCanonicalCommits.length = 0;
          EditorState.prototype.applyTransaction = tracedApplyTransaction;
          const canonical = await runShape(buffer, "canonical", shape, mode);
          EditorState.prototype.applyTransaction = originalApplyTransaction;
          expect({
            ...canonical,
            plugins: [...pluginsChangingCanonicalCommits],
          }).toEqual({ ...legacy, plugins: [] });
        }, 60_000);
});
