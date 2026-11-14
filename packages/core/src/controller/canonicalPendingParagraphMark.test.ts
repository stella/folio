/**
 * A public operation that edits the properties of a paragraph whose mark is
 * pending deletion waits for that change in both engines, under the same
 * reason and without naming a canonical gap; once the deletion is accepted
 * or rejected, both engines apply it and give the same package.
 */

import { panic } from "better-result";
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  FOLIO_PARAGRAPH_ALIGNMENT_VALUES,
  type FolioDocumentOperation,
} from "../ai-edits";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { createEmptyDocx } from "../docx/rezip";
import { schema } from "../prosemirror/schema";
import { CanonicalPublicOperations } from "./canonicalPublicOperations";

setDefaultTimeout(propertyTestTimeout(240_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const JOINED = "10000001";

const paragraph = (id: string, text: string, markDeleted = false) =>
  `<w:p w14:paraId="${id}">${markDeleted ? '<w:pPr><w:rPr><w:del w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr>' : ""}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const STYLES = `${HEAD}<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>`;

/** "Alpha", whose mark is pending deletion, then "Beta" and "Gamma". */
const source = async () => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set("word/styles.xml", STYLES);
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${paragraph(JOINED, "Alpha", true)}${paragraph("10000002", "Beta")}${paragraph("10000003", "Gamma")}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

type Mode = "direct" | "tracked-changes";

const operationArbitrary = fc.oneof(
  fc
    .record({
      alignment: fc.constantFrom(...FOLIO_PARAGRAPH_ALIGNMENT_VALUES),
      styleId: fc.constantFrom(undefined, "Heading1", null),
    })
    .map(
      ({ alignment, styleId }) =>
        (blockId: string, _text: string): FolioDocumentOperation => ({
          id: "op",
          type: "setBlockParagraphProperties",
          blockId,
          properties: {
            alignment,
            ...(styleId === undefined ? {} : { styleId }),
          },
        }),
    ),
  fc.constant((blockId: string, text: string): FolioDocumentOperation => ({
    id: "op",
    type: "replaceBlock",
    blockId,
    text,
    styleId: "Heading1",
  })),
);

/** Each body paragraph's text and the properties these operations set. */
const paragraphsOf = async (bytes: ArrayBuffer) => {
  const document = await parseDocx(bytes.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  return document.package.document.content.flatMap((block) =>
    block.type === "paragraph"
      ? [
          {
            text: block.content
              .flatMap((item) =>
                item.type === "run"
                  ? item.content.flatMap((content) =>
                      content.type === "text" ? [content.text] : [],
                    )
                  : [],
              )
              .join(""),
            alignment: block.formatting?.alignment,
            styleId: block.formatting?.styleId,
          },
        ]
      : [],
  );
};

/** The package pending, with every change accepted, and with every change rejected. */
const outcomes = async (bytes: ArrayBuffer) => {
  const resolved = async (decision: "accept" | "reject") => {
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
    if (decision === "accept") reviewer.acceptAll();
    else reviewer.rejectAll();
    return paragraphsOf(await reviewer.toBuffer());
  };
  return {
    pending: await paragraphsOf(bytes),
    accepted: await resolved("accept"),
    rejected: await resolved("reject"),
  };
};

const viaDefaultEngine = async (
  bytes: ArrayBuffer,
  operation: FolioDocumentOperation,
  mode: Mode,
) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
  const result = reviewer.applyOperations([operation], { mode });
  return {
    skipped: result.skipped.map(({ reason }) => reason),
    refusals: result.skipped.flatMap(({ canonicalRefusal }) =>
      canonicalRefusal === undefined ? [] : [canonicalRefusal],
    ),
    bytes: await reviewer.toBuffer(),
  };
};

const viaCanonicalEngine = async (
  bytes: ArrayBuffer,
  operation: FolioDocumentOperation,
  mode: Mode,
) => {
  const document = await parseDocx(bytes.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  const session = createCanonicalSession(
    document,
    document.package.styles,
  ).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const executor = new CanonicalPublicOperations({
    session,
    getState: () => state,
    publish: (commit: CanonicalCommit) => {
      const published = publishCanonicalProjection({ state, commit, session });
      if (published.isErr()) return false;
      state = published.value.state;
      return true;
    },
  });
  const result = executor.apply({
    snapshot: createFolioAIEditSnapshot(state.doc),
    batch: { version: 1, mode, operations: [operation] },
  });
  return {
    skipped: result.skipped.map(({ reason }) => reason),
    refusals: result.issues.flatMap(({ canonicalRefusal }) =>
      canonicalRefusal === undefined ? [] : [canonicalRefusal],
    ),
    bytes: (
      await serializeCanonicalSave({ snapshot: session.captureSaveSnapshot() })
    ).buffer,
  };
};

/** The paragraph holding "Alpha" once the join is resolved, and its words. */
const alphaParagraph = async (bytes: ArrayBuffer) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
  return (
    reviewer.getContent().find(({ text }) => text.startsWith("Alpha")) ??
    panic("The resolved package lost the paragraph holding Alpha.")
  );
};

describe("a paragraph whose mark is pending deletion", () => {
  test("both engines refuse a property edit alike, and apply it alike once the deletion is accepted or rejected", async () => {
    const sourceBytes = await source();
    await assertProperty(
      fc.asyncProperty(
        operationArbitrary,
        fc.constantFrom<Mode>("direct", "tracked-changes"),
        fc.constantFrom("pending", "accept", "reject"),
        async (operationFor, mode, resolution) => {
          let bytes = sourceBytes;
          if (resolution !== "pending") {
            const reviewer = await FolioDocxReviewer.fromBuffer(
              sourceBytes.slice(0),
            );
            if (resolution === "accept") reviewer.acceptAll();
            else reviewer.rejectAll();
            bytes = await reviewer.toBuffer();
          }
          const target =
            resolution === "pending"
              ? { id: JOINED, text: "Alpha" }
              : await alphaParagraph(bytes);
          const operation = operationFor(target.id, target.text);
          const legacy = await viaDefaultEngine(bytes, operation, mode);
          const canonical = await viaCanonicalEngine(bytes, operation, mode);
          expect(legacy.refusals).toEqual([]);
          expect(canonical.refusals).toEqual([]);
          expect(canonical.skipped).toEqual(legacy.skipped);
          if (resolution === "pending") {
            expect(legacy.skipped).toEqual(["pendingParagraphMarkDeletion"]);
            return;
          }
          expect(legacy.skipped).toEqual([]);
          expect(await outcomes(canonical.bytes)).toEqual(
            await outcomes(legacy.bytes),
          );
        },
      ),
      {
        id: "both engines refuse a property edit alike, and apply it alike once the deletion is accepted or rejected",
        numRuns: 60,
      },
    );
  });
});
