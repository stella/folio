/**
 * A direct replacement that clears formatting the new text must not carry
 * (a highlight or shading), over text that has none, writes plain text: it
 * joins the plain run beside it instead of standing apart with an empty
 * property set.
 */

import { expect, test } from "bun:test";

import type { Document } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { compileEditorIntent } from "../editorIntent";
import { OP_STORIES } from "../types";

const document: Document = {
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000002",
          content: [
            { type: "run", content: [{ type: "text", text: "Cell value" }] },
          ],
        },
      ],
    },
  },
};

const at = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: "00000002",
  offset,
});

for (const runPropsPatch of [
  undefined,
  { highlight: null, shading: null },
] as const)
  test(`replacing a word with ${runPropsPatch === undefined ? "no patch" : "a clearing patch"} leaves one plain run`, () => {
    const compiled = compileEditorIntent(document, {
      intent: {
        type: "replaceText",
        from: at(0),
        to: at(4),
        text: "Updated",
        ...(runPropsPatch === undefined ? {} : { runPropsPatch }),
      },
      mode: { type: "editing" },
    });
    if (compiled.isErr()) throw compiled.error;
    const applied = applyDocumentOps(document, compiled.value.ops);
    if (applied.isErr()) throw applied.error;
    expect(applied.value.document.package.document.content).toEqual([
      {
        type: "paragraph",
        paraId: "00000002",
        content: [
          { type: "run", content: [{ type: "text", text: "Updated value" }] },
        ],
      },
    ]);
  });
