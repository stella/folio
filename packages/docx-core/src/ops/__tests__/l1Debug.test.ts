import { test } from "bun:test";

import type { Document } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { REVISION_DECISIONS, DOCUMENT_OP_TYPES, OP_STORIES } from "../types";
import { type OpSeed, trackedOpFor } from "./documentArbitraries";

// Scratch debug: L1 seed -1097168474 on the kept-boundary branch.
const document = {"package": {"document": {"content": [{"type": "blockSdt", "properties": {"sdtType": "group", "id": 1}, "content": [{"type": "paragraph", "content": [{"type": "inlineWrapper", "kind": "bidi", "control": "embedding", "content": [{"type": "run", "content": [{"type": "text", "text": "a"}]}]}, {"type": "moveToRangeEnd", "id": 1}], "paraId": "00000001"}]}, {"type": "paragraph", "content": [], "paraId": "00000002"}], "sections": [{"properties": {"pageWidth": 12240, "pageHeight": 15840}, "content": [{"type": "blockSdt", "properties": {"sdtType": "group", "id": 1}, "content": [{"type": "paragraph", "content": [{"type": "inlineWrapper", "kind": "bidi", "control": "embedding", "content": [{"type": "run", "content": [{"type": "text", "text": "a"}]}]}, {"type": "moveToRangeEnd", "id": 1}], "paraId": "00000001"}]}, {"type": "paragraph", "content": [], "paraId": "00000002"}]}], "finalSectionProperties": {"pageWidth": 12240, "pageHeight": 15840}, "comments": [{"id": 7, "author": "A", "content": [{"type": "paragraph", "paraId": "7FFFFFF0", "content": []}]}]}, "footnotes": [{"type": "footnote", "id": 2, "content": [{"type": "paragraph", "paraId": "7FFFFFF1", "content": []}]}], "settings": {"defaultTabStop": 720}, "properties": {"title": "synthetic"}}, "warnings": ["kept"]} as unknown as Document;
const seed = {"kind": 329006644, "block": 0, "first": 106331061, "second": 0, "third": 0, "zeroWidth": {"first": 0, "second": 15320971}, "depth": 1, "text": "x", "formatting": {}, "inherit": true, "runPatch": {}, "paragraphPatch": {}, "content": [], "fresh": 268435456} as unknown as OpSeed;

test("debug L1 seed -1097168474", () => {
  const op = trackedOpFor(document, seed);
  console.error("DEBUGOP", JSON.stringify(op));
  const tracked = applyDocumentOp(structuredClone(document), op);
  const direct = { ...op };
  Reflect.deleteProperty(direct, "revision");
  const directed = applyDocumentOp(structuredClone(document), direct);
  if (tracked.isErr() || directed.isErr()) { console.error("DEBUGERR", JSON.stringify([tracked.isErr() && tracked.error, directed.isErr() && directed.error])); return; }
  console.error("DEBUGTRACKED", JSON.stringify(tracked.value.document.package.document.content), JSON.stringify(tracked.value.revisions));
  console.error("DEBUGDIRECT", JSON.stringify(directed.value.document.package.document.content));
  const accepted = applyDocumentOp(tracked.value.document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [...tracked.value.revisions],
    decision: REVISION_DECISIONS.ACCEPT,
  });
  console.error("DEBUGACCEPTED", accepted.isOk() ? JSON.stringify(accepted.value.document.package.document.content) : JSON.stringify(accepted.error));
});
