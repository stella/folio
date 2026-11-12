import { test } from "bun:test";

import type { Document } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { REVISION_DECISIONS, DOCUMENT_OP_TYPES, OP_STORIES } from "../types";
import { type OpSeed, trackedOpFor } from "./documentArbitraries";

// Scratch debug: L1 seed -1077732988.
const document = {"package": {"document": {"content": [{"type": "paragraph", "content": [{"type": "inlineSdt", "properties": {"sdtType": "richText", "id": 1, "tag": "clause"}, "content": [{"type": "hyperlink", "rId": "rId9", "href": "https://example.org", "children": [{"type": "inlineWrapper", "kind": "bidi", "control": "embedding", "content": [{"type": "commentRangeStart", "id": 24}, {"type": "commentRangeEnd", "id": 1}]}]}, {"type": "run", "content": [{"type": "text", "text": "a"}, {"type": "tab"}]}]}, {"type": "inlineSdt", "properties": {"sdtType": "richText", "id": 2, "tag": "clause"}, "content": [{"type": "hyperlink", "rId": "rId9", "href": "https://example.org", "children": [{"type": "inlineWrapper", "kind": "bidi", "control": "embedding", "content": [{"type": "commentRangeEnd", "id": 24}]}]}, {"type": "hyperlink", "rId": "rId9", "href": "https://example.org", "children": [{"type": "run", "content": [{"type": "text", "text": "a"}]}]}]}], "paraId": "00000001"}], "sections": [{"properties": {"pageWidth": 12240, "pageHeight": 15840}, "content": [{"type": "paragraph", "content": [{"type": "inlineSdt", "properties": {"sdtType": "richText", "id": 1, "tag": "clause"}, "content": [{"type": "hyperlink", "rId": "rId9", "href": "https://example.org", "children": [{"type": "inlineWrapper", "kind": "bidi", "control": "embedding", "content": [{"type": "commentRangeStart", "id": 24}, {"type": "commentRangeEnd", "id": 1}]}]}, {"type": "run", "content": [{"type": "text", "text": "a"}, {"type": "tab"}]}]}, {"type": "inlineSdt", "properties": {"sdtType": "richText", "id": 2, "tag": "clause"}, "content": [{"type": "hyperlink", "rId": "rId9", "href": "https://example.org", "children": [{"type": "inlineWrapper", "kind": "bidi", "control": "embedding", "content": [{"type": "commentRangeEnd", "id": 24}]}]}, {"type": "hyperlink", "rId": "rId9", "href": "https://example.org", "children": [{"type": "run", "content": [{"type": "text", "text": "a"}]}]}]}], "paraId": "00000001"}]}], "finalSectionProperties": {"pageWidth": 12240, "pageHeight": 15840}, "comments": [{"id": 7, "author": "A", "content": [{"type": "paragraph", "paraId": "7FFFFFF0", "content": []}]}]}, "footnotes": [{"type": "footnote", "id": 2, "content": [{"type": "paragraph", "paraId": "7FFFFFF1", "content": []}]}], "settings": {"defaultTabStop": 720}, "properties": {"title": "synthetic"}}, "warnings": ["kept"]} as unknown as Document;
const seed = {"kind": 507463091, "block": 0, "first": 2147483647, "second": 0, "third": 0, "depth": 1, "text": "x", "formatting": {}, "inherit": false, "runPatch": {}, "paragraphPatch": {}, "content": [], "fresh": 268435456} as unknown as OpSeed;

test("debug L1 seed -1077732988", () => {
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
