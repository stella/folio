import assert from "node:assert/strict";
import { EditorState } from "prosemirror-state";
import {
  expectedFailure,
  FINDING_SYMPTOMS,
} from "../../../../../test/consumer-scenarios/support/known-issues";
import type { Document } from "../../types/document";
import { expectParagraphAttrs } from "../attrs";
import { schema, singletonManager } from "../schema";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const commands = {
  setTabs: () =>
    singletonManager.requireCommand("setTabs")([
      { position: 1440, alignment: "right", leader: "dot" },
    ]),
  addTabStop: () => singletonManager.requireCommand("addTabStop")(1440, "right", "dot"),
  removeTabStop: () => singletonManager.requireCommand("removeTabStop")(720),
};
for (const [name, factory] of Object.entries(commands)) {
  expectedFailure(
    "LEGACY_PARAGRAPH_TAB_EDITS_LOST",
    `legacy ${name} saves edited imported paragraph tabs`,
    FINDING_SYMPTOMS.LEGACY_PARAGRAPH_TAB_EDITS_LOST,
    () => {
      const original = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                paraId: "12345678",
                content: [{ type: "run", content: [{ type: "text", text: "Paragraph" }] }],
                formatting: { tabs: [{ position: 720, alignment: "left", leader: "none" }] },
              },
            ],
          },
        },
      } satisfies Document;
      let state = EditorState.create({ schema, doc: toProseDoc(original) });
      assert.equal(
        factory()(state, (transaction) => {
          state = state.apply(transaction);
        }),
        true,
      );
      const current = expectParagraphAttrs(state.doc.child(0)).tabs ?? undefined;
      assert.notDeepEqual(current, original.package.document.content[0].formatting.tabs);
      const saved = fromProseDoc(state.doc, original).package.document.content.at(0);
      assert.ok(saved?.type === "paragraph");
      assert.deepEqual(
        saved.formatting?.tabs,
        current,
        `legacy ${name} serialization retains original paragraph tabs`,
      );
    },
  );
}
