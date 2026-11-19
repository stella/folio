import assert from "node:assert/strict";
import { test } from "node:test";

import { introducedIssues } from "../support/validity.ts";

const zoom = "word/settings.xml: w:zoom: Missing required attribute percent";
const margin = "word/document.xml: w:pgMar: Missing required attribute footer";

test("a save keeping the source's own schema issues introduces none", () => {
  assert.deepEqual(introducedIssues([zoom], [zoom]), []);
});

test("a save's new issue is reported beside the source's own", () => {
  assert.deepEqual(introducedIssues([zoom, margin], [zoom]), [margin]);
});

test("an issue the save repeats more often than the source has counts once more", () => {
  assert.deepEqual(introducedIssues([zoom, zoom], [zoom]), [zoom]);
});

test("a valid source tolerates no issue", () => {
  assert.deepEqual(introducedIssues([zoom], []), [zoom]);
});
