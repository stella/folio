import assert from "node:assert/strict";
import { test } from "node:test";

import { FINDING_SYMPTOMS, matchesFailureSymptom } from "../support/known-issues.ts";

test("known note rejection matches an extra empty paragraph in either note kind", () => {
  const symptom = FINDING_SYMPTOMS.REJECT_KEEPS_PARAGRAPH_INSERTED_IN_DELETED_NOTE;
  const diagnostic = (story: string, kind: string, text: string) =>
    new Error(
      "end of flow: [rejectAll] the flow's batches replayed tracked and rejected do not give the fixture back (fixture → rejected):\n" +
        `    ${JSON.stringify({ type: story, noteId: 37 })}: .blocks[1]: undefined → ${JSON.stringify({ kind, text })}`,
    );
  for (const story of ["footnote", "endnote"]) {
    for (const kind of ["paragraph", "heading", "listItem"]) {
      assert.equal(matchesFailureSymptom(symptom, diagnostic(story, kind, "")), true);
      assert.equal(matchesFailureSymptom(symptom, diagnostic(story, kind, "Changed.")), false);
    }
  }
  assert.equal(matchesFailureSymptom(symptom, diagnostic("main", "paragraph", "")), false);
  assert.equal(
    matchesFailureSymptom(symptom, new Error("the reopened package shows something else")),
    false,
  );
});
