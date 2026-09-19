/**
 * The two folio-core input boundaries where a character XML cannot hold
 * enters: an agent-supplied operation batch, and a paste.
 *
 * Both drop it there rather than at the writer, so the value folio applies is
 * the value it reports, and the package it saves opens.
 */

import { describe, expect, test } from "bun:test";

import { setContentControlContent } from "../../content-controls/mutateContentControls";
import { parseFolioDocumentOperationBatch } from "../../document-operations";
import type { Document } from "../../types/document";
import { parseClipboardHtml } from "../../utils/clipboard";

const NUL = String.fromCodePoint(0);
const LONE_HIGH_SURROGATE = String.fromCodePoint(0xd8_00);

describe("the document-operations batch boundary", () => {
  test("drops what XML cannot hold from replacement text", () => {
    const batch = parseFolioDocumentOperationBatch({
      version: 1,
      operations: [
        {
          id: "op-1",
          type: "replaceInBlock",
          blockId: "para-1",
          find: "as drafted",
          replace: `as amended\r${NUL} on${LONE_HIGH_SURROGATE}`,
        },
      ],
    });
    const operation = batch.operations[0];
    expect(operation?.type).toBe("replaceInBlock");
    expect(operation && "replace" in operation ? operation.replace : null).toBe(
      "as amended\r on\uFFFD",
    );
  });
});

describe("the content-control fill boundary", () => {
  test("drops what XML cannot hold from the value a control is filled with", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "blockSdt",
              properties: { sdtType: "richText", tag: "party" },
              content: [{ type: "paragraph", content: [] }],
            },
          ],
        },
      },
    };
    const filled = setContentControlContent(document, { tag: "party" }, `Acme${NUL} Ltd`);
    const control = filled.package.document.content.at(0);
    const paragraph = control?.type === "blockSdt" ? control.content.at(0) : undefined;
    const run = paragraph?.type === "paragraph" ? paragraph.content.at(0) : undefined;
    const text = run?.type === "run" ? run.content.at(0) : undefined;
    expect(text?.type === "text" ? text.text : null).toBe("Acme Ltd");
  });
});

describe("the paste boundary", () => {
  test("drops what XML cannot hold from pasted plain text", () => {
    const { runs } = parseClipboardHtml("", `pasted${NUL} text`);
    const content = runs.at(0)?.content.at(0);
    expect(content?.type === "text" ? content.text : null).toBe("pasted text");
  });
});
