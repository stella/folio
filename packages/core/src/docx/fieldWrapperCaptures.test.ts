/**
 * A field a resolved change left holding a bookmark or a control reads, for
 * the editor, as the parser reads the markup the serializer writes for it:
 * the field's own children it reads as markup it keeps.
 */
import { expect, test } from "bun:test";

import type { ParagraphContent, Run, SimpleField } from "../types/document";
import { withFieldWrapperCaptures } from "./fieldWrapperCaptures";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});

test("content a resolved change left in a field reads as the capture the parser reads", () => {
  const field: SimpleField = {
    type: "simpleField",
    instruction: "PAGE",
    fieldType: "PAGE",
    content: [
      run("1"),
      { type: "bookmarkStart", id: 3, name: "_Ref3" },
      run("2"),
      {
        type: "inlineSdt",
        properties: { sdtType: "richText" },
        content: [run("9")],
      },
      { type: "bookmarkEnd", id: 3 },
    ],
  };
  const [read] = withFieldWrapperCaptures([field]);
  if (read?.type !== "simpleField")
    throw new TypeError("The field stays a field.");
  expect(read.content.map((item: ParagraphContent) => item.type)).toEqual([
    "run",
    "preservedInline",
    "run",
    "preservedInline",
    "preservedInline",
  ]);
  const [, start, , , end] = read.content;
  expect(start?.type === "preservedInline" && start.xml).toContain(
    'w:bookmarkStart w:id="3" w:name="_Ref3"',
  );
  expect(end?.type === "preservedInline" && end.xml).toContain(
    'w:bookmarkEnd w:id="3"',
  );
});

test("a field holding only what the parser reads is returned as it is", () => {
  const field: SimpleField = {
    type: "simpleField",
    instruction: "PAGE",
    fieldType: "PAGE",
    content: [run("1")],
  };
  const content = [field];
  expect(withFieldWrapperCaptures(content)).toBe(content);
});
