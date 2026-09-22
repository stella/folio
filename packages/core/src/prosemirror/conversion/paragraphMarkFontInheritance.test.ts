import { expect, test } from "bun:test";

import type { Run, TextFormatting } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const sourceWithParagraphMarkFont = (formatting?: TextFormatting) => {
  const source = createEmptyDocument();
  source.package.styles = {
    docDefaults: {
      rPr: { fontFamily: { csTheme: "minorBidi" } },
    },
    styles: [],
  };
  const run: Run = { type: "run", content: [{ type: "text", text: "body" }] };
  if (formatting) {
    run.formatting = formatting;
  }
  source.package.document.content = [
    {
      type: "paragraph",
      formatting: { runProperties: { fontFamily: { csTheme: "minorHAnsi" }, italic: true } },
      content: [run],
    },
  ];
  return source;
};

test("a paragraph-mark font does not become direct body-run formatting", () => {
  for (const direct of [undefined, { fontFamily: { csTheme: "minorHAnsi" } }] as const) {
    const source = sourceWithParagraphMarkFont(direct);
    const saved = fromProseDoc(toProseDoc(source, { styles: source.package.styles }), source);
    const paragraph = saved.package.document.content.at(0);
    const run = paragraph?.type === "paragraph" ? paragraph.content.at(0) : undefined;
    expect(run?.type === "run" ? run.formatting : undefined).toEqual(direct);

    const savedAgain = fromProseDoc(toProseDoc(saved, { styles: saved.package.styles }), saved);
    expect(savedAgain.package.document.content).toEqual(saved.package.document.content);
  }
});

test("an empty paragraph keeps its paragraph-mark font for painting", () => {
  const source = sourceWithParagraphMarkFont();
  const paragraph = source.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    throw new Error("Expected one paragraph");
  }
  paragraph.content = [];
  const prose = toProseDoc(source, { styles: source.package.styles });
  expect(prose.firstChild?.attrs["defaultTextFormatting"]?.fontFamily?.csTheme).toBe("minorHAnsi");
});
