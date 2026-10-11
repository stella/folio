import assert from "node:assert/strict";
import { test } from "node:test";

import { deletedNoteSymptom } from "../support/deleted-note-symptom.ts";

const diagnostic =
  "step 0: the reopened package shows something else than the reviewer that saved it";
const state = (notes: string) => ({ blocks: [{ id: "7735D914", text: "Body." }], notes });
const failure = (actual: unknown, expected: unknown, message = diagnostic) => {
  try {
    assert.deepEqual(actual, expected, message);
  } catch (error) {
    return new Error("Replay failed", { cause: new Error("Flow failed", { cause: error }) });
  }
  return undefined;
};

test("deleted-note symptom uses assertion values across note types and retained stories", () => {
  for (const kind of ["footnote", "endnote"]) {
    for (const retained of ["", "[header default] Header.\n"]) {
      const actual = retained.trimEnd();
      const expected = `${retained}[${kind} #2] Note body.`;
      assert.equal(deletedNoteSymptom(failure(state(actual), state(expected))), true);
    }
  }
  assert.equal(
    deletedNoteSymptom(failure(state(""), state("[footnote #2] Foot.\n[endnote #3] End."))),
    true,
  );
});

test("deleted-note symptom rejects other differences and a fixed defect", () => {
  const expected = state("[footnote #2] Note body.");
  for (const error of [
    failure(state("[footnote #2] Changed."), expected),
    failure(state("[footnote #3] Note body."), expected),
    failure({ ...state(""), blocks: [] }, expected),
    failure(state(""), state("[header default] Header.")),
    failure(state(""), expected, "unrelated comparison"),
    failure(expected, expected),
    new Error(`${diagnostic}\n- '[footnote #2] Note body.'`),
  ])
    assert.equal(deletedNoteSymptom(error), false);
  const cycle = new Error("cycle");
  cycle.cause = cycle;
  assert.equal(deletedNoteSymptom(cycle), false);
});
