import { expect, test } from "bun:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { FolioDocxReviewer } from "./headless";

// A paragraph inserted before a block content control or custom XML that ends
// the body, holding its last paragraph, goes whole when it is rejected.

const inner: Paragraph = { type: "paragraph", paraId: "10000004", content: [] };
const before: Paragraph = {
  type: "paragraph",
  paraId: "10000001",
  content: [{ type: "run", content: [{ type: "text", text: "before" }] }],
};
const table: BlockContent = {
  type: "table",
  columnWidths: [2400],
  rows: [
    {
      type: "tableRow",
      cells: [
        {
          type: "tableCell",
          content: [{ type: "paragraph", paraId: "10000002", content: [] }],
        },
      ],
    },
  ],
};
const endings = {
  blockSdt: {
    type: "blockSdt",
    properties: { sdtType: "richText", id: 5 },
    content: [inner],
  },
  blockCustomXml: {
    type: "blockCustomXml",
    openingXml: '<w:customXml w:element="standing">',
    closingXml: "</w:customXml>",
    content: [inner],
  },
} satisfies Record<string, BlockContent>;
const anchors = {
  table: { blockId: "10000002", content: [before, table] },
  paragraph: { blockId: "10000001", content: [before] },
} satisfies Record<string, { blockId: string; content: BlockContent[] }>;

const cases = Object.entries(endings).flatMap(([ending, control]) =>
  Object.entries(anchors).map(
    ([anchor, { blockId, content }]) =>
      [anchor, ending, blockId, [...content, control]] as const,
  ),
);

const open = (bytes: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(bytes.slice(0), { session: "canonical" });

const bodyOf = (reviewer: FolioDocxReviewer) =>
  structuredClone(reviewer.toDocument().package.document.content);

for (const [anchor, ending, blockId, content] of cases) {
  const input = async (): Promise<ArrayBuffer> => {
    const document: Document = { package: { document: { content } } };
    return (await prepareCanonicalDocxInput(await createDocx(document)))
      .unwrap()
      .slice().buffer;
  };
  const insert = (
    reviewer: FolioDocxReviewer,
    mode: "tracked-changes" | "suggested",
    text: string,
  ) => {
    const { applied, skipped } = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [{ id: "op", type: "insertAfterBlock", blockId, text }],
      },
      { undefinedReferences: "refuse" },
    );
    expect(skipped).toEqual([]);
    expect(applied.map(({ id }) => id)).toEqual(["op"]);
  };

  test.each(["", "new", "new\nmore"])(
    `a tracked insertion of %j after a ${anchor}, before a ${ending} ending the body, rejects to its source`,
    async (text) => {
      const bytes = await input();
      for (const route of ["individual", "all"] as const) {
        const reviewer = await open(bytes);
        const source = bodyOf(reviewer);
        insert(reviewer, "tracked-changes", text);
        if (route === "all") reviewer.rejectAll();
        else
          for (const change of reviewer.getChanges().toReversed())
            expect(reviewer.rejectChange(change.id)).toBe(true);
        expect(reviewer.getChanges()).toEqual([]);
        expect(bodyOf(reviewer)).toEqual(source);
        expect(bodyOf(await open(await reviewer.toBuffer()))).toEqual(source);
      }
    },
  );

  test(`a suggested insertion after a ${anchor}, before a ${ending} ending the body, saves as rejected and restages`, async () => {
    const bytes = await input();
    const reviewer = await open(bytes);
    const source = bodyOf(reviewer);
    insert(reviewer, "suggested", "new");
    const ids = reviewer.getContent().map(({ id }) => id);
    const records = reviewer.exportPendingSuggestions();
    const saved = await reviewer.toBuffer();
    const restaged = await open(saved);
    expect(bodyOf(restaged)).toEqual(source);
    for (const loaded of restaged.loadPendingSuggestions(records))
      expect(loaded).toMatchObject({ status: "restaged" });
    expect(restaged.getContent().map(({ id }) => id)).toEqual(ids);
    expect(restaged.rejectSuggestion("op")).toBe(true);
    expect(bodyOf(restaged)).toEqual(source);
  });
}
