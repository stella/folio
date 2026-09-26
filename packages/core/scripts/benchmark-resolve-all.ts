/** Focused resolve-all scaling; use --legacy for the retained range command. */
import { loadavg } from "node:os";
import { history } from "prosemirror-history";
import { EditorState, type Command } from "prosemirror-state";
import { schema } from "@stll/folio-core/prosemirror/schema";
import {
  acceptAllChanges,
  rejectAllChanges,
  acceptChange,
  rejectChange,
} from "@stll/folio-core/prosemirror/commands/comments";
import { ParagraphChangeTrackerExtension } from "@stll/folio-core/prosemirror/extensions/features/ParagraphChangeTrackerExtension";

const sampleCommand = (state: EditorState, command: Command) => {
  let steps = 0;
  const start = performance.now();
  command(state, (transaction) => {
    steps = transaction.steps.length;
    state.apply(transaction);
  });
  return { milliseconds: performance.now() - start, steps };
};

const legacy = process.argv.includes("--legacy");
const loadBefore = loadavg().at(0);
const measurements = [];
for (const blocks of [250, 1000, 2200, 4400]) {
  const doc = schema.node(
    "doc",
    null,
    Array.from({ length: blocks }, (_, index) =>
      schema.node("paragraph", { paraId: index.toString(16).padStart(8, "0") }, [
        schema.text("old ", [
          schema.marks.deletion.create({ revisionId: index * 2 + 1, author: "Reviewer" }),
        ]),
        schema.text("new ", [
          schema.marks.insertion.create({ revisionId: index * 2 + 2, author: "Reviewer" }),
        ]),
      ]),
    ),
  );
  const state = EditorState.create({
    doc,
    plugins: [
      history(),
      ...(ParagraphChangeTrackerExtension().onSchemaReady({ schema }).plugins ?? []),
    ],
  });
  for (const mode of ["accept", "reject"] as const) {
    const accept = legacy ? acceptChange(0, doc.content.size) : acceptAllChanges();
    const reject = legacy ? rejectChange(0, doc.content.size) : rejectAllChanges();
    const command = mode === "accept" ? accept : reject;
    const samples: number[] = [];
    let steps = 0;
    for (let sample = 0; sample < 5; sample++) {
      const result = sampleCommand(state, command);
      steps = result.steps;
      if (sample > 0) samples.push(result.milliseconds);
    }
    samples.sort((left, right) => left - right);
    measurements.push({ blocks, changes: blocks * 2, mode, milliseconds: samples.at(2), steps });
  }
}
console.log(
  JSON.stringify(
    {
      command: legacy ? "range" : "all",
      loadAverage: { before: loadBefore, after: loadavg().at(0) },
      measurements,
    },
    null,
    2,
  ),
);
