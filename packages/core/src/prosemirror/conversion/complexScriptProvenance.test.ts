import { describe, expect, test } from "bun:test";

import type { Document, TextFormatting } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { schema } from "../schema";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const documentWithRun = (
  formatting: TextFormatting,
  inheritedFormatting?: TextFormatting,
): Document => {
  const document = createEmptyDocument();
  document.package.styles = {
    ...(inheritedFormatting ? { docDefaults: { rPr: inheritedFormatting } } : {}),
    styles: [],
  };
  document.package.document.content = [
    {
      type: "paragraph",
      content: [
        {
          type: "run",
          formatting,
          content: [{ type: "text", text: "direct" }],
        },
      ],
    },
  ];
  return document;
};

const firstRunFormatting = (document: Document): TextFormatting | undefined => {
  const paragraph = document.package.document.content.at(0);
  const run = paragraph?.type === "paragraph" ? paragraph.content.at(0) : undefined;
  return run?.type === "run" ? run.formatting : undefined;
};

const editableRoundTrip = (document: Document): TextFormatting | undefined =>
  firstRunFormatting(
    fromProseDoc(toProseDoc(document, { styles: document.package.styles }), document),
  );

describe("complex-script run-property provenance", () => {
  test("preserves an absent szCs beside a direct size and an inherited complex size", () => {
    expect(editableRoundTrip(documentWithRun({ fontSize: 22 }, { fontSizeCs: 30 }))).toEqual({
      fontSize: 22,
    });
  });

  test("preserves an absent szCs beside a standalone direct size", () => {
    expect(editableRoundTrip(documentWithRun({ fontSize: 22 }))).toEqual({ fontSize: 22 });
  });

  test("does not materialize inherited ordinary or complex formatting on the run", () => {
    expect(
      editableRoundTrip(
        documentWithRun(
          {},
          {
            bold: true,
            boldCs: false,
            italic: true,
            italicCs: false,
            fontSize: 22,
            fontSizeCs: 30,
          },
        ),
      ),
    ).toBeUndefined();
  });

  test("preserves explicit complex properties equal to their ordinary partners", () => {
    expect(
      editableRoundTrip(
        documentWithRun({
          bold: true,
          boldCs: true,
          italic: true,
          italicCs: true,
          fontSize: 22,
          fontSizeCs: 22,
        }),
      ),
    ).toEqual({
      bold: true,
      boldCs: true,
      italic: true,
      italicCs: true,
      fontSize: 22,
      fontSizeCs: 22,
    });
  });

  test("keeps complex mirrors for formatting authored directly in ProseMirror", () => {
    const pmDocument = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("authored", [
          schema.mark("bold"),
          schema.mark("italic"),
          schema.mark("fontSize", { size: 22 }),
        ]),
      ]),
    ]);

    expect(firstRunFormatting(fromProseDoc(pmDocument))).toMatchObject({
      bold: true,
      boldCs: true,
      italic: true,
      italicCs: true,
      fontSize: 22,
      fontSizeCs: 22,
    });
  });
});
