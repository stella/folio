/**
 * A suggested paste whose copied deletion of a break holds a cut of its own
 * saves the cut on the pasted break, with the copied deletion beside it, and
 * the reopened suggestion resolves to the documents the unsaved one does.
 */
import { describe, expect, test } from "bun:test";

import {
  applyDocumentOps,
  compileEditorIntent,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import type { Document, Paragraph, Run } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});

const opened = async (): Promise<Document> => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
      `<w:p w14:paraId="00000001"><w:r><w:t>Here.</w:t></w:r></w:p>` +
      `<w:sectPr/></w:body></w:document>`,
  );
  return parseDocx(
    await zip.generate({ compression: "store" }).then((bytes) => bytes.buffer),
    { preloadFonts: false },
  );
};

const apply = (document: Document, ops: readonly DocumentOp[]) => {
  const applied = applyDocumentOps(document, ops);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const paragraphsOf = (document: Document): Paragraph[] =>
  document.package.document.content.flatMap((block) =>
    block.type === "paragraph" ? [block] : [],
  );

const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
) =>
  apply(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds,
      decision,
    },
  ]);

/** What each decision pair leaves, the copied deletion resolved first. */
const outcomes = (document: Document): unknown[] => {
  const mark = paragraphsOf(document).find(
    ({ pPrMark }) => pPrMark?.deletion !== undefined,
  )?.pPrMark;
  const copied = mark?.deletion?.info.id;
  if (copied === undefined) throw new TypeError("No copied deletion.");
  const rest = Array.from({ length: 400 }, (_, index) => index).filter(
    (id) => id !== copied,
  );
  const out: unknown[] = [];
  for (const first of [REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT])
    for (const second of [REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT])
      out.push(
        paragraphsOf(
          resolve(resolve(document, [copied], first), rest, second),
        ).map(({ content }) => content),
      );
  return out;
};

describe("a copied deletion holding a cut, suggested beside a pasted break", () => {
  test("saves its cut and reopens to resolve alike", async () => {
    const document = await opened();
    const link = (text: string) => ({
      type: "hyperlink" as const,
      rId: "rId9",
      href: "https://example.org",
      children: [run(text)],
    });
    const compiled = compileEditorIntent(document, {
      intent: {
        type: "replaceFragment",
        from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 5 },
        to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 5 },
        openStart: 0,
        openEnd: 0,
        paragraphs: [
          {
            type: "paragraph",
            paraId: "0000000D",
            pPrMark: {
              kind: "del",
              info: {
                id: 93,
                author: "Earlier",
                date: "2026-01-01T00:00:00Z",
              },
              resolutionJoin: 1,
              splitDate: "2025-12-01T00:00:00Z",
            },
            content: [link("Cop")],
          },
          { type: "paragraph", paraId: "0000000E", content: [link("ied")] },
          { type: "paragraph", content: [run("tail")] },
        ],
      },
      mode: {
        type: "suggesting",
        revision: { id: 300, author: "Reviewer", date: "2026-02-03T00:00:00Z" },
        newIds: {
          revision: Array.from({ length: 30 }, (_, index) => 301 + index),
        },
      },
    });
    if (compiled.isErr()) throw compiled.error;
    const suggested = apply(document, compiled.value.ops);
    const reopened = await parseDocx(
      await repackDocx(suggested, { updateModifiedDate: false }),
      { preloadFonts: false },
    );
    const markOf = (from: Document) => {
      const mark = paragraphsOf(from).find(
        ({ pPrMark }) => pPrMark?.deletion !== undefined,
      )?.pPrMark;
      if (mark === undefined) throw new TypeError("No pasted break.");
      return {
        kind: mark.kind,
        resolutionJoin: mark.resolutionJoin,
        splitDate: mark.splitDate,
        deletion: mark.deletion?.info.author,
      };
    };
    expect(markOf(reopened)).toEqual({
      kind: "ins",
      resolutionJoin: 2,
      splitDate: "2025-12-01T00:00:00Z",
      deletion: "Earlier",
    });
    expect(markOf(reopened)).toEqual(markOf(suggested));
    expect(outcomes(reopened)).toEqual(outcomes(suggested));
  });
});
