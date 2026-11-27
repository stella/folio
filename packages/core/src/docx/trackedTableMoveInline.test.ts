import { expect, test } from "bun:test";

import {
  applyDocumentOp,
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";

import type { Document, Table } from "../types/document";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { serializeTable } from "./serializer/tableSerializer";
import { parseTable } from "./tableParser";
import { parseXmlDocument } from "./xmlParser";

const ownerId = "00000001";
const sourceId = "00000002";
const plain = "<w:r><w:t>Owner</w:t></w:r>";
const direction = '<w:dir w:val="rtl">';
const link = '<w:hyperlink w:anchor="Target">';
const atoms =
  '<w:r><w:t>Linked</w:t><w:tab/><w:br/><w:sym w:font="Symbol" w:char="F020"/><w:softHyphen/><w:noBreakHyphen/></w:r>';

const readTable = (xml: string): Table => {
  const element = parseXmlDocument(
    xml.replace(
      "<w:tbl",
      `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"`,
    ),
  );
  if (element === null) throw new TypeError("Expected well-formed XML.");
  const table = parseTable(element, new Map(), null, null, null, null);
  if (table === undefined) throw new TypeError("Expected a table.");
  return table;
};

const tableOf = (document: Document): Table => {
  const table = document.package.document.content.at(0);
  if (table?.type !== "table") throw new TypeError("Expected a table.");
  return table;
};

const inlineContentOf = (table: Table) =>
  table.rows.map(({ cells }) =>
    cells.map(({ content }) =>
      content.map((block) =>
        block.type === "paragraph"
          ? { paraId: block.paraId, content: block.content }
          : block,
      ),
    ),
  );

test.each([
  `${direction}${link}${atoms}</w:hyperlink></w:dir>`,
  `${link}${direction}${atoms}</w:dir></w:hyperlink>`,
  `${atoms}${link}${atoms}</w:hyperlink>${direction}${atoms}</w:dir>`,
  "<w:r><w:tab/></w:r>",
])(
  "tracked table merges preserve moved links, direction, and run atoms: %s",
  (content) => {
    const table = readTable(
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
        `<w:tr><w:tc><w:p w14:paraId="${ownerId}">${plain}</w:p></w:tc>` +
        `<w:tc><w:p w14:paraId="${sourceId}">${content}</w:p></w:tc></w:tr></w:tbl>`,
    );
    const original = {
      package: {
        document: {
          content: [
            table,
            { type: "paragraph", paraId: "00000003", content: [] },
          ],
        },
      },
    } satisfies Document;
    const op = {
      type: DOCUMENT_OP_TYPES.MERGE_CELLS,
      story: OP_STORIES.MAIN,
      blockId: ownerId,
      top: 0,
      bottom: 1,
      left: 0,
      right: 2,
      newBlockIds: [],
    } as const;
    const direct = applyDocumentOp(original, op).unwrap().document;
    const pending = applyDocumentOp(original, {
      ...op,
      newBlockIds: ["00000004"],
      revision: { id: 100, author: "Reviewer", date: "2026-11-27T00:00:00Z" },
      newIds: { revision: [101, 102, 103, 104, 105, 106] },
    }).unwrap();
    expect(
      applyDocumentOps(pending.document, pending.inverse).unwrap().document,
    ).toEqual(original);
    const xml = serializeTable(tableOf(pending.document), serializeParagraph);
    expect(xml).toContain("<w:moveFrom ");
    expect(xml).toContain("<w:moveTo ");
    const reopened = {
      ...original,
      package: {
        document: {
          content: [
            readTable(xml),
            ...original.package.document.content.slice(1),
          ],
        },
      },
    };
    for (const decision of [
      REVISION_DECISIONS.ACCEPT,
      REVISION_DECISIONS.REJECT,
    ]) {
      const expected =
        decision === REVISION_DECISIONS.ACCEPT ? direct : original;
      for (const document of [pending.document, reopened]) {
        const resolved = applyDocumentOp(document, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [104],
          decision,
        }).unwrap().document;
        if (document === pending.document) {
          expect(resolved).toEqual(expected);
        } else {
          expect(inlineContentOf(tableOf(resolved))).toEqual(
            inlineContentOf(tableOf(expected)),
          );
        }
      }
    }
  },
);
