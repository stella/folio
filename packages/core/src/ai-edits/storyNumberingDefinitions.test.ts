/**
 * A note starts its lists against the package alone, so the numbering it
 * started is read in the note's own definitions, which label its paragraphs,
 * and not in the main story's. In every mode, a list a note operation starts
 * is defined where the note's snapshot points.
 */
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createNote } from "../docx/noteReferenceMark";
import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const NOTE = { type: "footnote", noteId: 1 } as const;

const fixture = async (): Promise<ArrayBuffer> => {
  const document = createEmptyDocument({ initialText: "Body" });
  const body = document.package.document.content.at(0);
  if (body?.type !== "paragraph") throw new Error("expected a body paragraph");
  body.content.push({
    type: "run",
    content: [{ type: "footnoteRef", id: 1 }],
  });
  document.package.footnotes = [
    createNote("footnote", 1, [
      {
        type: "paragraph",
        paraId: "10000001",
        content: [{ type: "run", content: [{ type: "text", text: "Note" }] }],
      },
    ]),
  ];
  return createDocx(document);
};

/** The saved document and notes parts, which a retained story would rewrite. */
const storyParts = async (buffer: ArrayBuffer) => {
  const zip = await ZipArchive.read(buffer);
  return Promise.all(
    ["word/document.xml", "word/footnotes.xml"].map(
      async (path) => (await zip.entry(path)?.text()) ?? null,
    ),
  );
};

const openReviewer = async (bytes?: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer((bytes ?? (await fixture())).slice(0));

describe("a list a note operation starts", () => {
  for (const mode of ["direct", "tracked-changes", "suggested"] as const)
    test(`is read in the note's numbering, ${mode}`, async () => {
      const reviewer = await openReviewer();
      const before = reviewer.readNumberingDefinitions();
      const result = reviewer.applyDocumentOperationsToStory({
        story: NOTE,
        batch: {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            {
              id: "op",
              type: "insertAfterBlock",
              blockId: "10000001",
              text: "Listed",
              numbering: { kind: "newList", format: "bullet" },
            },
          ],
        },
        undefinedReferences: "refuse",
      });
      expect(result.applied.map(({ id }) => id)).toEqual(["op"]);
      const listed = reviewer
        .snapshotStory(NOTE)
        ?.blocks.find(
          (block) => block.kind !== "diagnostic" && block.text === "Listed",
        );
      if (!listed || listed.kind === "diagnostic")
        throw new Error("expected the listed note paragraph");
      const reference = listed.listReference;
      if (!reference) throw new Error("expected the paragraph to show a list");
      expect(
        reviewer
          .readNumberingDefinitions(NOTE)
          ?.some(
            ({ numId, level }) =>
              numId === reference.numId && level === reference.level,
          ),
      ).toBe(true);
      expect(reviewer.readNumberingDefinitions()).toEqual(before);
    });

  test("reading a note's numbering leaves the document as it was", async () => {
    const bytes = await fixture();
    const read = await openReviewer(bytes);
    expect(read.readNumberingDefinitions(NOTE)).toEqual(
      read.readNumberingDefinitions(),
    );
    const unread = await openReviewer(bytes);
    expect(await storyParts(await read.toBuffer())).toEqual(
      await storyParts(await unread.toBuffer()),
    );
  });

  test("a story the document lacks has no numbering to read", async () => {
    const reviewer = await openReviewer();
    expect(
      reviewer.readNumberingDefinitions({ type: "endnote", noteId: 9 }),
    ).toBeNull();
  });
});
