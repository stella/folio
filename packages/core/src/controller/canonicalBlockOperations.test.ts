/**
 * Block-level public operations in the canonical session give the document
 * the default engine gives: the same blocks and paragraph properties pending,
 * once every change is accepted, and once every change is rejected.
 */

import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import type { FolioDocumentOperation } from "../ai-edits";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { schema } from "../prosemirror/schema";
import { CanonicalPublicOperations } from "./canonicalPublicOperations";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const paragraph = (id: string, text: string, properties = "") =>
  `<w:p w14:paraId="${id}">${properties ? `<w:pPr>${properties}</w:pPr>` : ""}${text ? `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>` : ""}</w:p>`;
const table = (id: string, text: string) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${paragraph(id, text)}</w:tc></w:tr></w:tbl>`;

/** Bodies, each holding the paragraph `10000002` the operations address. */
const BODIES = {
  "between two paragraphs": [
    paragraph("10000001", "Alpha", '<w:jc w:val="center"/>'),
    paragraph("10000002", "Beta", '<w:ind w:left="720"/>'),
    paragraph("10000003", "Gamma", '<w:spacing w:after="240"/>'),
  ],
  "ending the body": [
    paragraph("10000001", "Alpha", '<w:jc w:val="center"/>'),
    paragraph("10000002", "Beta", '<w:ind w:left="720"/>'),
  ],
  "before a table": [
    paragraph("10000001", "Alpha"),
    paragraph("10000002", "Beta", '<w:ind w:left="720"/>'),
    table("10000004", "Cell"),
    paragraph("10000003", "Gamma"),
  ],
} as const;

const packageOf = async (body: readonly string[]) => {
  const zip = await ZipArchive.read(
    await (await import("../docx/rezip")).createEmptyDocx(),
  );
  const styles = (await zip.entry("word/styles.xml")?.text()) ?? "";
  // A heading style with its own alignment and spacing, for a restyle.
  zip.set(
    "word/styles.xml",
    styles.replace(
      "</w:styles>",
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/><w:spacing w:before="480"/><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>',
    ),
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body.join("")}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

/** Each body block's text and paragraph properties, as a package reads. */
const blocksOf = async (bytes: ArrayBuffer) => {
  const document = await parseDocx(bytes.slice(0), {
    preloadFonts: false,
    detectVariables: false,
  });
  return document.package.document.content.map((block) =>
    block.type === "paragraph"
      ? {
          text: block.content
            .flatMap((item) =>
              item.type === "run"
                ? item.content.flatMap((leaf) =>
                    leaf.type === "text" ? [leaf.text] : [],
                  )
                : [],
            )
            .join(""),
          // Source paragraphs keep their ids; a new one's is the engine's own.
          id: block.paraId?.startsWith("1000000") ? block.paraId : "new",
          pageBreakBefore: block.formatting?.pageBreakBefore,
          runFormatting: block.content.flatMap((item) =>
            item.type === "run" && item.formatting !== undefined
              ? [item.formatting]
              : [],
          ),
          styleId: block.formatting?.styleId,
          outlineLevel: block.formatting?.outlineLevel,
          alignment: block.formatting?.alignment,
          indentLeft: block.formatting?.indentLeft,
          indentRight: block.formatting?.indentRight,
          spaceBefore: block.formatting?.spaceBefore,
          spaceAfter: block.formatting?.spaceAfter,
        }
      : { type: block.type },
  );
};

/** The package pending, with every change accepted, and with every change rejected. */
const outcomes = async (bytes: ArrayBuffer) => {
  const resolved = async (decision: "accept" | "reject") => {
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0));
    if (decision === "accept") reviewer.acceptAll();
    else reviewer.rejectAll();
    return blocksOf(await reviewer.toBuffer());
  };
  return {
    pending: await blocksOf(bytes),
    accepted: await resolved("accept"),
    rejected: await resolved("reject"),
  };
};

type Mode = "direct" | "tracked-changes";
type Suggestion = "accept" | "reject";

const viaDefaultEngine = async (
  source: ArrayBuffer,
  operation: FolioDocumentOperation,
  mode: Mode,
) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source.slice(0));
  const result = reviewer.applyOperations([operation], {
    mode,
    undefinedReferences: "refuse",
  });
  expect(result.skipped).toEqual([]);
  return await reviewer.toBuffer();
};

/**
 * The canonical session's saved package after the batch; a suggested batch
 * is resolved as `suggestion` says before the save.
 */
const canonicalRun = async (
  source: ArrayBuffer,
  operation: FolioDocumentOperation,
  mode: Mode | "suggested",
) => {
  const document = await parseDocx(source.slice(0), {
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
    undefinedReferences: "refuse",
  });
  return { session, state: () => state, skipped: result.skipped };
};

const viaCanonicalEngine = async (
  source: ArrayBuffer,
  operation: FolioDocumentOperation,
  mode: Mode | "suggested",
  suggestion?: Suggestion,
) => {
  const run = await canonicalRun(source, operation, mode);
  expect(run.skipped).toEqual([]);
  const { session } = run;
  let state = run.state();
  if (mode === "suggested") {
    const commit = session
      .prepareSuggestion(state, {
        suggestionId: operation.id,
        resolution: suggestion ?? "reject",
        author: "Adopter",
        date: "2026-02-02T00:00:00Z",
      })
      .unwrap();
    const published = publishCanonicalProjection({ state, commit, session });
    if (published.isErr()) throw published.error;
    state = published.value.state;
  }
  return (
    await serializeCanonicalSave({ snapshot: session.captureSaveSnapshot() })
  ).buffer;
};

/** Cases the canonical session declares unavailable, with its reason. */
const REFUSED: Record<string, string> = {
  "deleteBlock before a table, tracked-changes":
    "Canonical tracked deletion of a paragraph before a table is unavailable.",
  "deleteBlock ending the body, direct":
    "Canonical direct deletion of the paragraph that ends its block list is unavailable.",
  "deleteBlock before a table, direct":
    "Canonical direct deletion of the paragraph that ends its block list is unavailable.",
  "insertAfterBlock ending the body, direct":
    "Canonical insertion after a paragraph no paragraph follows is unavailable.",
  "insertAfterBlock ending the body, tracked-changes":
    "Canonical insertion after a paragraph no paragraph follows is unavailable.",
  "insertAfterBlock before a table, direct":
    "Canonical insertion after a paragraph no paragraph follows is unavailable.",
  "insertAfterBlock before a table, tracked-changes":
    "Canonical insertion after a paragraph no paragraph follows is unavailable.",
  "insertAfterBlock blank ending the body, direct":
    "Canonical insertion after a paragraph no paragraph follows is unavailable.",
  "insertAfterBlock blank ending the body, tracked-changes":
    "Canonical insertion after a paragraph no paragraph follows is unavailable.",
  "insertAfterBlock blank before a table, direct":
    "Canonical insertion after a paragraph no paragraph follows is unavailable.",
  "insertAfterBlock blank before a table, tracked-changes":
    "Canonical insertion after a paragraph no paragraph follows is unavailable.",
  "insertAfterBlock without inheritance ending the body, direct":
    "Canonical insertion after a paragraph no paragraph follows is unavailable.",
  "insertAfterBlock without inheritance ending the body, tracked-changes":
    "Canonical insertion after a paragraph no paragraph follows is unavailable.",
  "insertAfterBlock without inheritance before a table, direct":
    "Canonical insertion after a paragraph no paragraph follows is unavailable.",
  "insertAfterBlock without inheritance before a table, tracked-changes":
    "Canonical insertion after a paragraph no paragraph follows is unavailable.",
};

const OPERATIONS = {
  deleteBlock: { id: "op", type: "deleteBlock", blockId: "10000002" },
  insertAfterBlock: {
    id: "op",
    type: "insertAfterBlock",
    blockId: "10000002",
    text: "Inserted",
  },
  insertBeforeBlock: {
    id: "op",
    type: "insertBeforeBlock",
    blockId: "10000002",
    text: "Inserted",
  },
  "insertAfterBlock blank": {
    id: "op",
    type: "insertAfterBlock",
    blockId: "10000002",
    text: "",
  },
  "insertAfterBlock without inheritance": {
    id: "op",
    type: "insertAfterBlock",
    blockId: "10000002",
    text: "Inserted",
    inheritFormatting: false,
  },
  "insertBeforeBlock restyled": {
    id: "op",
    type: "insertBeforeBlock",
    blockId: "10000002",
    text: "Inserted",
    styleId: "Heading1",
    alignment: null,
    spacing: { spaceAfter: 60 },
    pageBreakBefore: true,
  },
  "setBlockParagraphProperties alignment": {
    id: "op",
    type: "setBlockParagraphProperties",
    blockId: "10000002",
    properties: { alignment: "right" },
  },
  "setBlockParagraphProperties spacing and cleared indentation": {
    id: "op",
    type: "setBlockParagraphProperties",
    blockId: "10000002",
    properties: { spacing: { spaceBefore: 120 }, indentation: null },
  },
  "setBlockParagraphProperties partial indentation": {
    id: "op",
    type: "setBlockParagraphProperties",
    blockId: "10000002",
    properties: {
      indentation: { indentRight: 360 },
      outlineLevel: { kind: "heading", level: 1 },
    },
  },
  "setBlockParagraphProperties style": {
    id: "op",
    type: "setBlockParagraphProperties",
    blockId: "10000002",
    properties: { styleId: "Heading1" },
  },
} as const satisfies Record<string, FolioDocumentOperation>;

describe("block-level public operations match the default engine", () => {
  for (const [kind, operation] of Object.entries(OPERATIONS))
    for (const [shape, body] of Object.entries(BODIES))
      for (const mode of ["direct", "tracked-changes"] as const) {
        const name = `${kind} ${shape}, ${mode}`;
        const refused = REFUSED[name];
        test(name, async () => {
          const source = await packageOf(body);
          if (refused !== undefined) {
            const { skipped } = await canonicalRun(source, operation, mode);
            expect(skipped.map(({ message }) => message)).toEqual([refused]);
            return;
          }
          const expected = await outcomes(
            await viaDefaultEngine(source, operation, mode),
          );
          expect(
            await outcomes(await viaCanonicalEngine(source, operation, mode)),
          ).toEqual(expected);
          if (mode !== "tracked-changes") return;
          // Suggested, the edit saves as nothing until it is resolved:
          // rejecting gives the source back, accepting the tracked edit.
          expect(
            await blocksOf(
              await viaCanonicalEngine(
                source,
                operation,
                "suggested",
                "reject",
              ),
            ),
          ).toEqual(await blocksOf(source));
          expect(
            await outcomes(
              await viaCanonicalEngine(
                source,
                operation,
                "suggested",
                "accept",
              ),
            ),
          ).toEqual(expected);
        });
      }
});
