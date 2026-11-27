import { expect, test } from "bun:test";

import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";

import type { Document } from "../types/document";
import { parseParagraph } from "./paragraphParser";
import { FOLIO_REVIEW_HISTORY_NAMESPACE } from "./reviewHistoryNamespace";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument } from "./xmlParser";

const paraId = "00000001";
const parse = (xml: string) => {
  const element = parseXmlDocument(
    xml.replace(
      /^<w:p(?=[\s>])/u,
      `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:folio="${FOLIO_REVIEW_HISTORY_NAMESPACE}"`,
    ),
  );
  if (element === null) throw new TypeError("Expected a paragraph element.");
  return parseParagraph(element, null, null, null, null, null);
};

test.each([
  '<w:hyperlink w:anchor="Target"/>',
  '<w:sdt><w:sdtPr><w:id w:val="12"/><w:richText/></w:sdtPr><w:sdtContent/></w:sdt>',
  '<w:dir w:val="rtl"/>',
  '<w:dir w:val="rtl"><w:hyperlink w:anchor="Target"/></w:dir>',
  '<w:sdt><w:sdtPr><w:id w:val="12"/><w:richText/></w:sdtPr><w:sdtContent><w:hyperlink w:anchor="Target"/></w:sdtContent></w:sdt>',
])(
  "an empty inline container keeps its review after save and reopen: %s",
  (xml) => {
    const original = parse(`<w:p w14:paraId="${paraId}">${xml}</w:p>`);
    const document = {
      package: { document: { content: [original] } },
    } satisfies Document;
    const at = { story: OP_STORIES.MAIN, blockId: paraId, offset: 0 };
    const edit = {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: { ...at, zeroWidthBefore: 0 },
      to: { ...at, zeroWidthBefore: 1 },
    } as const;
    const direct = applyDocumentOp(document, edit).unwrap();
    const tracked = applyDocumentOp(document, {
      ...edit,
      revision: { id: 20, author: "Reviewer", date: "2026-01-02T03:04:05Z" },
    }).unwrap();
    const paragraph = tracked.document.package.document.content.at(0);
    if (paragraph?.type !== "paragraph")
      throw new TypeError("Expected the retained paragraph.");
    expect(tracked.revisions).toEqual([20]);
    const saved = serializeParagraph(paragraph);
    expect(saved).toContain("<w:del ");
    const reopened = {
      package: { document: { content: [parse(saved)] } },
    } satisfies Document;
    for (const decision of Object.values(REVISION_DECISIONS)) {
      const resolved = applyDocumentOp(reopened, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: tracked.revisions,
        decision,
      }).unwrap();
      expect(resolved.document).toEqual(
        decision === REVISION_DECISIONS.ACCEPT ? direct.document : document,
      );
    }
  },
);
