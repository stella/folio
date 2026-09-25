/**
 * `fromProseDoc` maps the paragraph `direction` discriminated union to the
 * serialized OOXML `w:bidi` tri-state: a manual RTL/LTR decision becomes
 * `true`/`false` (the latter as `<w:bidi w:val="0"/>`), and an undecided
 * paragraph is omitted.
 *
 * Regression guard: a manual LTR (forced via setLtr) must survive save/reload as
 * `bidi: false`; if it collapsed to "undecided", the seed auto-detector would
 * re-flip an Arabic paragraph back to RTL. These tests lock the conversion end
 * to end (PM direction → model → serialized XML).
 */

import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { serializeParagraph } from "../../docx/serializer/paragraphSerializer";
import type { Document } from "../../types/document";
import type { ParagraphDirection } from "../paragraphDirection";
import { schema } from "../schema";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const RTL: ParagraphDirection = { source: "manual", value: "rtl" };
const LTR: ParagraphDirection = { source: "manual", value: "ltr" };
const AUTO: ParagraphDirection = { source: "auto" };

const paraNode = (text: string, direction: ParagraphDirection | null) =>
  schema.node("doc", null, [schema.node("paragraph", { direction }, [schema.text(text)])]);

const firstParagraphBidi = (doc: Document): unknown => {
  const block = doc.package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("expected a paragraph");
  }
  return block.formatting?.bidi;
};

describe("fromProseDoc direction → bidi (new paragraphs, no original)", () => {
  test("manual LTR becomes bidi=false (not dropped)", () => {
    expect(firstParagraphBidi(fromProseDoc(paraNode("عربي", LTR)))).toBe(false);
  });

  test("manual RTL becomes bidi=true", () => {
    expect(firstParagraphBidi(fromProseDoc(paraNode("عربي", RTL)))).toBe(true);
  });

  test("undecided is omitted", () => {
    expect(firstParagraphBidi(fromProseDoc(paraNode("Agreement", null)))).toBeUndefined();
  });
});

describe("fromProseDoc direction → bidi (changed vs original)", () => {
  // An imported RTL paragraph that the user forces to LTR: the direction now
  // differs from the original, which is where the old truthiness check dropped
  // the explicit `false`.
  test("RTL original forced to LTR keeps bidi=false", () => {
    const original: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              formatting: { bidi: true },
              content: [{ type: "run", content: [{ type: "text", text: "عربي" }] }],
            },
          ],
        },
      },
    };
    const pmDoc = toProseDoc(original);
    const forcedLtr = EditorState.create({ doc: pmDoc }).tr.setNodeMarkup(0, undefined, {
      ...pmDoc.child(0).attrs,
      direction: LTR,
    }).doc;

    expect(firstParagraphBidi(fromProseDoc(forcedLtr, original))).toBe(false);
  });
});

describe("pageBreakBefore tri-state (same class, fallback path)", () => {
  test("explicit pageBreakBefore=false is preserved on a new paragraph", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { pageBreakBefore: false }, [schema.text("Body")]),
    ]);
    const block = fromProseDoc(doc).package.document.content.at(0);
    if (block?.type !== "paragraph") {
      throw new Error("expected a paragraph");
    }
    expect(block.formatting?.pageBreakBefore).toBe(false);
  });
});

describe("manual LTR survives serialization (save invariant)", () => {
  test('direction=manual ltr serializes as <w:bidi w:val="0"/>', () => {
    const block = fromProseDoc(paraNode("عربي", LTR)).package.document.content.at(0);
    if (block?.type !== "paragraph") {
      throw new Error("expected a paragraph");
    }
    expect(serializeParagraph(block)).toContain('<w:bidi w:val="0"/>');
  });
});

/**
 * Regression guard for the load-time auto-bidi seeder
 * (`AutoBidiDetectionExtension.ensureBaseDirectionInState`): it sets
 * `direction: { source: "auto" }` on RTL-led paragraphs that arrived without
 * `w:bidi`, purely so the editor renders them RTL. That must be a rendering
 * decision only — a no-edit save must not turn it into a persisted `w:bidi`
 * the source paragraph never had.
 */
describe("auto-detected direction is view-only and never self-serializes", () => {
  test("auto direction on a new paragraph (no original) omits bidi", () => {
    expect(firstParagraphBidi(fromProseDoc(paraNode("عربي", AUTO)))).toBeUndefined();
  });

  test("no-edit save: a paragraph seeded auto from an original with no w:bidi stays free of w:bidi", () => {
    // Mirrors DOCX import: the source paragraph never had `w:bidi`, and the
    // load-time seeder (not modeled here directly) has already set
    // `direction: { source: "auto" }` on the in-memory PM node, same as it
    // would for an RTL-led paragraph on open. Saving without any user edit
    // must reproduce the original absence of `w:bidi`.
    const original: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [{ type: "run", content: [{ type: "text", text: "عربي" }] }],
            },
          ],
        },
      },
    };
    const pmDoc = toProseDoc(original);
    const seeded = EditorState.create({ doc: pmDoc }).tr.setNodeMarkup(0, undefined, {
      ...pmDoc.child(0).attrs,
      direction: AUTO,
    }).doc;

    expect(firstParagraphBidi(fromProseDoc(seeded, original))).toBeUndefined();
    const block = fromProseDoc(seeded, original).package.document.content.at(0);
    if (block?.type !== "paragraph") {
      throw new Error("expected a paragraph");
    }
    expect(serializeParagraph(block)).not.toContain("w:bidi");
  });

  test("mixed document: auto stays unwritten, manual and undecided round-trip as expected", () => {
    const doc = schema.node("doc", null, [
      schema.node("paragraph", { direction: null }, [schema.text("Agreement")]),
      schema.node("paragraph", { direction: AUTO }, [schema.text("عربي بلا علامة")]),
      schema.node("paragraph", { direction: RTL }, [schema.text("عربي بعلامة")]),
      schema.node("paragraph", { direction: LTR }, [schema.text("عربي بالإكراه")]),
    ]);
    const result = fromProseDoc(doc);
    const bidiOf = (index: number): unknown => {
      const block = result.package.document.content.at(index);
      if (block?.type !== "paragraph") {
        throw new Error("expected a paragraph");
      }
      return block.formatting?.bidi;
    };
    expect(bidiOf(0)).toBeUndefined(); // undecided LTR text
    expect(bidiOf(1)).toBeUndefined(); // auto-detected: view-only, not authored
    expect(bidiOf(2)).toBe(true); // manual RTL
    expect(bidiOf(3)).toBe(false); // manual LTR
  });

  test("explicit direction change from auto to manual RTL is written", () => {
    // The user invokes the direction toggle (ParagraphExtension's `setRtl`),
    // which always produces a `manual` decision — this is what actually makes
    // the paragraph's direction authored content.
    expect(firstParagraphBidi(fromProseDoc(paraNode("عربي", RTL)))).toBe(true);
  });
});
