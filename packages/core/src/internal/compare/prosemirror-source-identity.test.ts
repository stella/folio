import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";

import { schema } from "../../prosemirror/schema";
import { firstProseMirrorSourceIdentityDifferencePath } from "./prosemirror-source-identity";

const textBox = (groupId: string, anchorId: string, width = 200): PMNode =>
  schema.node(
    "textBox",
    {
      width,
      _docxPlacement: "inlineWithPrevious",
      _docxGroupId: groupId,
      _docxAnchorId: anchorId,
    },
    [schema.node("paragraph")],
  );

const linkedTextBoxes = ({
  anchorIds,
  groupIds,
}: {
  anchorIds: readonly [string, string];
  groupIds: readonly [string, string];
}): PMNode =>
  schema.node("doc", null, [
    schema.node("paragraph", null, [
      schema.node("textBoxAnchor", { anchorId: anchorIds[0] }),
      schema.node("textBoxAnchor", { anchorId: anchorIds[1] }),
    ]),
    textBox(groupIds[0], anchorIds[0]),
    textBox(groupIds[1], anchorIds[1]),
  ]);

describe("ProseMirror source identity", () => {
  test("alpha-compares per-load text-box identities while preserving their links", () => {
    const left = linkedTextBoxes({
      anchorIds: ["left:0:0", "left:0:1"],
      groupIds: ["left:0", "left:0"],
    });
    const right = linkedTextBoxes({
      anchorIds: ["right:0:0", "right:0:1"],
      groupIds: ["right:0", "right:0"],
    });

    expect(firstProseMirrorSourceIdentityDifferencePath(left, right)).toBe("");
  });

  test("rejects a changed text-box grouping equivalence class", () => {
    const left = linkedTextBoxes({
      anchorIds: ["left:0:0", "left:0:1"],
      groupIds: ["left:0", "left:0"],
    });
    const right = linkedTextBoxes({
      anchorIds: ["right:0:0", "right:1:0"],
      groupIds: ["right:0", "right:1"],
    });

    expect(firstProseMirrorSourceIdentityDifferencePath(left, right)).toBe(
      "doc.content[2].attrs._docxGroupId",
    );
  });

  test("rejects a text box linked to the wrong inline anchor", () => {
    const left = linkedTextBoxes({
      anchorIds: ["left:0:0", "left:0:1"],
      groupIds: ["left:0", "left:0"],
    });
    const right = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node("textBoxAnchor", { anchorId: "right:0:0" }),
        schema.node("textBoxAnchor", { anchorId: "right:0:1" }),
      ]),
      textBox("right:0", "right:0:1"),
      textBox("right:0", "right:0:0"),
    ]);

    expect(firstProseMirrorSourceIdentityDifferencePath(left, right)).toBe(
      "doc.content[1].attrs._docxAnchorId",
    );
  });

  test("compares every semantic text-box attribute exactly", () => {
    const left = linkedTextBoxes({
      anchorIds: ["left:0:0", "left:0:1"],
      groupIds: ["left:0", "left:0"],
    });
    const right = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node("textBoxAnchor", { anchorId: "right:0:0" }),
        schema.node("textBoxAnchor", { anchorId: "right:0:1" }),
      ]),
      textBox("right:0", "right:0:0", 201),
      textBox("right:0", "right:0:1"),
    ]);

    expect(firstProseMirrorSourceIdentityDifferencePath(left, right)).toBe(
      "doc.content[1].attrs.width",
    );
  });

  test("rejects malformed nominal identities even when both sides match", () => {
    const malformed = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.node("textBoxAnchor", { anchorId: "" })]),
      textBox("", ""),
    ]);

    expect(firstProseMirrorSourceIdentityDifferencePath(malformed, malformed)).toBe(
      "doc.content[0].content[0].attrs.anchorId",
    );
  });
});
