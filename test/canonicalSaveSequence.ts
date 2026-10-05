import JSZip from "jszip";
import { panic } from "better-result";
import { extractParagraphXml } from "../packages/core/src/docx/selectiveXmlPatch";
/** Generated adapter save histories share operations, not a PM-derived expected model. */
import { createEmptyDocument } from "../packages/core/src/utils/createDocument";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
} from "../packages/core/src/controller/canonicalOperations";

export const CANONICAL_SAVE_SEEDS = [20261004, 20261005, 20261006] as const;

export const canonicalSaveFixture = (seed: number) => {
  const document = createEmptyDocument({ initialText: "" });
  document.package.document.content = ["Direct", "Tracked", "Formatting", "Untouched"].map(
    (label, index) => ({
      type: "paragraph" as const,
      paraId: (seed + index).toString(16).padStart(8, "0").toUpperCase(),
      content: [
        { type: "run" as const, content: [{ type: "text" as const, text: `${label} ${seed}` }] },
      ],
    }),
  );
  return document;
};

export const canonicalSaveSequence = (seed: number): DocumentOp[][] => {
  const blockId = (index: number) => (seed + index).toString(16).padStart(8, "0").toUpperCase();
  const position = (index: number, offset: number) => ({
    story: "main" as const,
    blockId: blockId(index),
    offset,
  });
  const revision = (index: number) => ({
    id: index,
    author: `Author ${seed}`,
    date: "2026-10-04T12:00:00Z",
  });
  return [
    [
      {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: position(0, 0),
        text: `Edit ${seed % 97} `,
        runProps: "inherit",
      },
    ],
    [
      {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: position(1, 0),
        to: position(1, 1),
        revision: revision(1),
      },
    ],
    [
      {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: position(1, 1),
        text: `Review ${seed % 89} `,
        runProps: "inherit",
        revision: revision(2),
      },
    ],
    [
      {
        type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
        from: position(2, 0),
        to: position(2, 3),
        patch: { bold: true, italic: seed % 2 === 0 },
      },
    ],
    [
      {
        type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
        at: position(2, 3),
        newBlockId: blockId(10),
        newHalf: "second",
      },
    ],
    [
      {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: "main",
        blockId: blockId(2),
        nextBlockId: blockId(10),
        survivor: "first",
      },
    ],
  ];
};

export const canonicalSaveParagraphXml = async (bytes: ArrayBuffer, paraId: string) => {
  const zip = await JSZip.loadAsync(bytes);
  const part = zip.file("word/document.xml") ?? panic("Expected main document part");
  return (
    extractParagraphXml(await part.async("string"), paraId) ??
    panic("Expected preserved paragraph XML")
  );
};

export const canonicalSaveFeatureFlags = (seed: number) => ({
  selectiveSave: true,
  ...(seed % 2 === 0 ? {} : { selectiveSaveMaxBytes: 1 }),
});
