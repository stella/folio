import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import {
  documentXmlOf,
  inlineTokens,
  listNumberFieldDocx,
  type ParagraphSpec,
  paragraphMarkupOf,
  paragraphNode,
} from "./__tests__/listNumberFieldFixture";
import { parseDocx } from "./parser";
import { getSourceReplayToken, repackDocx } from "./rezip";

// The field-save generator exercises untracked saves. Tracked saves must also
// obey its fold rule when a paragraph changes, while retaining untouched bytes.
test("tracked replay retains folded LISTNUM bytes and normalizes edited captures", async () => {
  const spec = {
    paraId: "20000001",
    marker: "decimal",
    fields: [
      {
        instruction: " LISTNUM ",
        result: "(a)",
        formatting: "bold",
        before: "",
        gap: [],
        tab: true,
      },
    ],
    body: "Body text",
  } satisfies ParagraphSpec;
  const untouchedId = "20000002";
  const original = await listNumberFieldDocx([spec, { ...spec, paraId: untouchedId }]);
  const sourceXml = await documentXmlOf(original);
  const parsed = await parseDocx(original, { preloadFonts: false, sourceReplay: "tracked" });
  const source = parsed.package.document.content.at(0);
  expect(source?.type).toBe("paragraph");
  if (source?.type !== "paragraph") throw new TypeError("Expected numbered paragraph");
  expect(
    source.content.some((item) => item.type === "preservedInline" && item.foldedListNumber),
  ).toBe(true);

  const projected = toProseDoc(parsed);
  const unchanged = fromProseDoc(projected, parsed);
  expect(unchanged.package.document.content.at(0)).toBe(source);
  const unchangedToken = getSourceReplayToken(unchanged);
  expect(unchangedToken).toBeDefined();
  const unchangedSave = await repackDocx(unchanged, {
    updateModifiedDate: false,
    sourceReplay: unchangedToken,
  });
  expect(await documentXmlOf(unchangedSave)).toBe(sourceXml);

  // Typing before the hidden capture keeps it at the leading marker position.
  const { position } = paragraphNode(projected, spec.paraId);
  const state = EditorState.create({ doc: projected });
  const edited = fromProseDoc(state.tr.insertText("Prefix ", position + 1).doc, parsed);
  const changed = edited.package.document.content.at(0);
  expect(changed).not.toBe(source);
  expect(edited.package.document.content.at(1)).toBe(parsed.package.document.content.at(1));
  if (changed?.type !== "paragraph") throw new TypeError("Expected edited paragraph");
  expect(changed.content.some((item) => item.type === "complexField")).toBe(false);
  expect(
    changed.content.some((item) => item.type === "preservedInline" && item.foldedListNumber),
  ).toBe(true);
  const editedToken = getSourceReplayToken(edited);
  expect(editedToken).toBeDefined();
  const saved = await repackDocx(edited, {
    updateModifiedDate: false,
    sourceReplay: editedToken,
  });
  const savedXml = await documentXmlOf(saved);
  expect(paragraphMarkupOf(savedXml, untouchedId)).toBe(paragraphMarkupOf(sourceXml, untouchedId));
  expect(inlineTokens(paragraphMarkupOf(savedXml, spec.paraId))).toEqual([
    "fldChar:begin",
    "code: LISTNUM ",
    "fldChar:separate",
    "text:(a)",
    "fldChar:end",
    "tab",
    "text:Prefix Body text",
  ]);
  const reopened = await parseDocx(saved, { preloadFonts: false });
  const reopenedParagraph = reopened.package.document.content.at(0);
  if (reopenedParagraph?.type !== "paragraph") throw new TypeError("Expected reopened paragraph");
  expect(
    reopenedParagraph.content.some(
      (item) => item.type === "preservedInline" && item.foldedListNumber,
    ),
  ).toBe(true);

  // Removing numbering makes the field visible, including on a tracked save.
  const unnumbered = fromProseDoc(
    state.tr.setNodeMarkup(position, undefined, {
      ...projected.nodeAt(position)?.attrs,
      numPr: null,
    }).doc,
    parsed,
  );
  const unnumberedParagraph = unnumbered.package.document.content.at(0);
  if (unnumberedParagraph?.type !== "paragraph") throw new TypeError("Expected plain paragraph");
  expect(unnumberedParagraph.content.some((item) => item.type === "complexField")).toBe(true);
  expect(
    unnumberedParagraph.content.some(
      (item) => item.type === "preservedInline" && item.foldedListNumber,
    ),
  ).toBe(false);
  expect(unnumbered.package.document.content.at(1)).toBe(parsed.package.document.content.at(1));
  const plainSave = await repackDocx(unnumbered, {
    updateModifiedDate: false,
    sourceReplay: getSourceReplayToken(unnumbered),
  });
  const plainXml = await documentXmlOf(plainSave);
  expect(paragraphMarkupOf(plainXml, untouchedId)).toBe(paragraphMarkupOf(sourceXml, untouchedId));
  const plainReopened = await parseDocx(plainSave, { preloadFonts: false });
  const plainParagraph = plainReopened.package.document.content.at(0);
  if (plainParagraph?.type !== "paragraph")
    throw new TypeError("Expected plain reopened paragraph");
  expect(plainParagraph.content.some((item) => item.type === "complexField")).toBe(true);
});
