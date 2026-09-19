/**
 * The two folio-core input boundaries where a character XML cannot hold
 * enters: an agent-supplied operation batch, and a paste.
 *
 * Both drop it there rather than at the writer, so the value folio applies is
 * the value it reports, and the package it saves opens.
 */

import { describe, expect, test } from "bun:test";

import { parseFolioDocumentOperationBatch } from "../../document-operations";
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

describe("the paste boundary", () => {
  test("drops what XML cannot hold from pasted plain text", () => {
    const { runs } = parseClipboardHtml("", `pasted${NUL} text`);
    const content = runs.at(0)?.content.at(0);
    expect(content?.type === "text" ? content.text : null).toBe("pasted text");
  });
});
