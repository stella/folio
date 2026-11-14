import { expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import {
  DOCUMENT_OP_SCHEMA_VERSION,
  validateDocumentBatch,
} from "@stll/docx-core/ops";

import { createCanonicalSession } from "../canonical/canonicalSession";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";

test("DEBUG note typing wire", async () => {
  const document = createEmptyDocument({ initialText: "Body seed" });
  const body = document.package.document.content.at(0);
  if (body?.type !== "paragraph") throw new Error("no body");
  body.content.push({ type: "run", content: [{ type: "footnoteRef", id: 1 }] });
  document.package.footnotes = [
    {
      type: "footnote",
      id: 1,
      content: [
        {
          type: "paragraph",
          content: [
            { type: "run", content: [{ type: "text", text: "Footnote seed" }] },
          ],
        },
      ],
    },
  ];
  const bytes = (
    await prepareCanonicalDocxInput(await createDocx(document))
  ).unwrap();
  const session = createCanonicalSession(
    await parseDocx(bytes, { preloadFonts: false, detectVariables: false }),
  ).unwrap();
  const story = { kind: "footnote", id: 1 } as const;
  let projection = session.projectStory(story).unwrap();
  let state = EditorState.create({ doc: projection.doc });
  for (const text of [" ", "r", "e"]) {
    const end = state.doc.content.size - 1;
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, end)),
    );
    const commit = session.prepareReplace(state, {
      from: end,
      to: end,
      text,
      story,
      semantic: "typing",
    });
    if (commit.isErr()) {
      console.log("DEBUG refused", commit.error.message);
      break;
    }
    console.log("DEBUG ops", JSON.stringify(commit.value.ops));
    const valid = validateDocumentBatch({
      schema: DOCUMENT_OP_SCHEMA_VERSION,
      opId: `o-${text}`,
      actor: "a",
      baseRev: 0,
      ops: commit.value.ops,
    });
    console.log("DEBUG valid", valid.isOk() ? "ok" : valid.error.message);
    const published = commit.value.publish();
    console.log("DEBUG published", published.isOk());
    projection = session.projectStory(story).unwrap();
    state = EditorState.create({ doc: projection.doc });
  }
  expect(true).toBe(true);
});
