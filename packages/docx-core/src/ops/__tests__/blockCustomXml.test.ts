import { expect, test } from "bun:test";

import type { Document } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { storyParagraphs } from "../blocks";
import { DOCUMENT_OP_TYPES, INHERIT_RUN_PROPS, OP_STORIES } from "../types";

test("document operations address paragraphs inside block custom XML", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: "blockCustomXml",
            openingXml: '<w:customXml w:element="clause">',
            closingXml: "</w:customXml>",
            content: [
              {
                type: "paragraph",
                paraId: "00000001",
                content: [{ type: "run", content: [{ type: "text", text: "clause" }] }],
              },
            ],
          },
        ],
      },
    },
  };

  expect(
    storyParagraphs(document.package.document).map(({ paragraph }) => paragraph.paraId),
  ).toEqual(["00000001"]);
  const edited = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 6 },
    text: " amended",
    runProps: INHERIT_RUN_PROPS,
  });
  expect(edited.isOk()).toBe(true);
  if (edited.isErr()) {
    return;
  }
  const wrapper = edited.value.document.package.document.content.at(0);
  expect(wrapper?.type).toBe("blockCustomXml");
  if (wrapper?.type !== "blockCustomXml") {
    return;
  }
  expect(wrapper.openingXml).toBe('<w:customXml w:element="clause">');
  expect(wrapper.content.at(0)?.type).toBe("paragraph");
  const paragraph = wrapper.content.at(0);
  if (paragraph?.type === "paragraph") {
    expect(paragraph.content).toEqual([
      { type: "run", content: [{ type: "text", text: "clause amended" }] },
    ]);
  }
});
