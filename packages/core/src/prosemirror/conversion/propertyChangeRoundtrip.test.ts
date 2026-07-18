import { describe, expect, test } from "bun:test";

import type {
  Paragraph,
  ParagraphPropertyChange,
  Run,
  RunPropertyChange,
} from "../../types/content";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const samplePropertyChange: ParagraphPropertyChange = {
  type: "paragraphPropertyChange",
  info: {
    id: 7,
    author: "Reviewer",
    date: "2026-05-15T12:00:00Z",
  },
  previousFormatting: { alignment: "left" },
  currentFormatting: { alignment: "center" },
};

function paragraphWithPropertyChange(): Paragraph {
  return {
    type: "paragraph",
    formatting: { alignment: "center" },
    propertyChanges: [samplePropertyChange],
    content: [
      {
        type: "run",
        formatting: {},
        content: [{ type: "text", text: "body" }],
      },
    ],
  };
}

describe("paragraph propertyChanges PM round-trip", () => {
  test("toProseDoc copies propertyChanges into the paragraph node attrs", () => {
    const document = {
      package: {
        document: {
          content: [paragraphWithPropertyChange()],
          finalSectionProperties: {},
        },
      },
    };
    const pmDoc = toProseDoc(document as never);
    // The paragraph PM node now carries the entries on a private attr —
    // editor code never reads it; it exists purely so fromProseDoc can
    // restore them after an edit.
    const paragraphNode = pmDoc.firstChild!;
    expect(paragraphNode.attrs["_propertyChanges"]).toEqual([samplePropertyChange]);
  });

  test("fromProseDoc restores propertyChanges back onto the Folio Paragraph", () => {
    // Simulate a no-op edit: convert to PM and immediately back. The
    // entries must survive even though the editor surfaces nothing in
    // UI for them — the previous behaviour silently stripped them on
    // every edit, corrupting the `w:pPrChange` history.
    const original = paragraphWithPropertyChange();
    const document = {
      package: {
        document: {
          content: [original],
          finalSectionProperties: {},
        },
      },
    };
    const pmDoc = toProseDoc(document as never);
    const roundtripped = fromProseDoc(pmDoc).package.document.content[0] as Paragraph | undefined;

    expect(roundtripped?.propertyChanges).toEqual([samplePropertyChange]);
  });

  test("paragraphs without propertyChanges round-trip without inventing them", () => {
    const plain: Paragraph = {
      type: "paragraph",
      formatting: {},
      content: [
        {
          type: "run",
          formatting: {},
          content: [{ type: "text", text: "plain" }],
        },
      ],
    };
    const document = {
      package: {
        document: {
          content: [plain],
          finalSectionProperties: {},
        },
      },
    };
    const pmDoc = toProseDoc(document as never);
    const roundtripped = fromProseDoc(pmDoc).package.document.content[0] as Paragraph | undefined;
    // No phantom propertyChanges should be attached.
    expect(roundtripped?.propertyChanges).toBeUndefined();
  });
});

const sampleRunPropertyChange: RunPropertyChange = {
  type: "runPropertyChange",
  info: {
    id: 8,
    author: "Reviewer",
    date: "2026-07-17T08:00:00Z",
  },
  previousFormatting: { italic: true },
  currentFormatting: { bold: true },
};

describe("run propertyChanges PM round-trip", () => {
  test("preserves run formatting revisions through the editable model", () => {
    const run: Run = {
      type: "run",
      formatting: { bold: true },
      propertyChanges: [sampleRunPropertyChange],
      content: [{ type: "text", text: "reviewed" }],
    };
    const document = {
      package: {
        document: {
          content: [{ type: "paragraph" as const, content: [run] }],
          finalSectionProperties: {},
        },
      },
    };

    const pmDoc = toProseDoc(document as never);
    const text = pmDoc.firstChild?.firstChild;
    expect(text?.marks.find((mark) => mark.type.name === "runPropertyChange")?.attrs).toEqual({
      changes: [sampleRunPropertyChange],
      // Provenance defaults: a parsed run-property change is always a user change.
      provenance: "user",
      suggestionId: null,
    });

    const roundtripped = fromProseDoc(pmDoc).package.document.content.at(0);
    const roundtrippedRun =
      roundtripped?.type === "paragraph"
        ? roundtripped.content.find((content) => content.type === "run")
        : undefined;
    expect(roundtrippedRun?.propertyChanges).toEqual([sampleRunPropertyChange]);
  });

  test("does not invent run formatting revisions", () => {
    const document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph" as const,
              content: [
                {
                  type: "run" as const,
                  formatting: { bold: true },
                  content: [{ type: "text" as const, text: "plain" }],
                },
              ],
            },
          ],
          finalSectionProperties: {},
        },
      },
    };

    const roundtripped = fromProseDoc(toProseDoc(document as never)).package.document.content.at(0);
    const run =
      roundtripped?.type === "paragraph"
        ? roundtripped.content.find((content) => content.type === "run")
        : undefined;
    expect(run?.propertyChanges).toBeUndefined();
  });
});
