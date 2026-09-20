/**
 * A tracked-change wrapper keeps what folio does not model, inside itself.
 *
 * `w:ins`, `w:del`, `w:moveFrom` and `w:moveTo` hold run-level content, and
 * the schema lets that content include `w:permStart`, `w:proofErr`,
 * `w:customXml` and the eight custom-XML revision ranges. folio modelled none
 * of them and let them fall off the end of the walk, so a reviewer accepted or
 * rejected a revision whose content was already gone.
 *
 * Position is the whole point here, so every assertion is about the markup
 * being *inside* the wrapper. A capture lifted out beside it survives a save
 * and still breaks the document: accepting the insertion then leaves the
 * markup behind, and rejecting it leaves markup that belonged to a change
 * nobody kept.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { propertyConfig } from "../../../../test/property-testing";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { acceptAllChanges, rejectAllChanges } from "../prosemirror/commands/comments";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph } from "../types/document";

import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** The four `CT_RunTrackChange` wrappers, with the model branch each parses to. */
const WRAPPERS = [
  { tag: "ins", type: "insertion" },
  { tag: "del", type: "deletion" },
  { tag: "moveFrom", type: "moveFrom" },
  { tag: "moveTo", type: "moveTo" },
] as const;

/**
 * Children the wrapper's content model declares and folio models nothing of.
 *
 * One per shape rather than all thirteen: an empty marker, a marker carrying
 * revision attributes, a transparent wrapper with text inside it, and an
 * element from a namespace the content model does not name at all.
 */
const UNMODELLED_CHILDREN = [
  '<w:permStart w:id="7" w:edGrp="everyone"/>',
  '<w:proofErr w:type="spellStart"/>',
  '<w:customXmlInsRangeStart w:id="3" w:author="Reviewer"/>',
  '<w:customXml w:element="party"><w:r><w:t>Acme</w:t></w:r></w:customXml>',
  '<x:note xmlns:x="urn:example:vendor" x:kind="aside">kept</x:note>',
] as const;

const parseParagraphXml = (xml: string): Paragraph => {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("Failed to parse the paragraph fixture");
  }
  return parseParagraph(root, null, null, null, null, null);
};

const paragraphXml = (tag: string, child: string): string =>
  `<w:p xmlns:w="${W}">` +
  `<w:${tag} w:id="1" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">` +
  `<w:r><w:t>text</w:t></w:r>${child}` +
  `</w:${tag}></w:p>`;

/** What the wrapper holds in the saved markup, or `null` when there is none. */
const insideWrapper = (savedXml: string, tag: string): string | null => {
  const open = savedXml.indexOf(`<w:${tag} `);
  const close = savedXml.indexOf(`</w:${tag}>`, open);
  if (open === -1 || close === -1) {
    return null;
  }
  return savedXml.slice(savedXml.indexOf(">", open) + 1, close);
};

const wrapParagraph = (paragraph: Paragraph): Document => ({
  package: { document: { content: [paragraph] } },
});

/** The paragraph as it comes back from the editor, with nothing edited. */
const throughTheEditor = (paragraph: Paragraph): Paragraph => {
  const input = wrapParagraph(paragraph);
  const out = fromProseDoc(toProseDoc(input), input);
  const first = out.package.document.content.at(0);
  if (first?.type !== "paragraph") {
    throw new Error("Expected the round trip to give a paragraph back");
  }
  return first;
};

type Resolution = "accept" | "reject";

/** The document after accepting or rejecting every tracked change in it. */
const resolveAll = (paragraph: Paragraph, resolution: Resolution): Document => {
  const input = wrapParagraph(paragraph);
  let state = EditorState.create({ schema, doc: toProseDoc(input) });
  const command = resolution === "accept" ? acceptAllChanges() : rejectAllChanges();
  command(state, (transaction) => {
    state = state.apply(transaction);
  });
  return fromProseDoc(state.doc, input);
};

describe("a run-level tracked change keeps the children folio does not model", () => {
  for (const { tag, type } of WRAPPERS) {
    test(`w:${tag} keeps a w:permStart inside itself across a save`, () => {
      const paragraph = parseParagraphXml(paragraphXml(tag, '<w:permStart w:id="7"/>'));

      const wrapper = paragraph.content.at(0);
      expect(wrapper?.type).toBe(type);
      if (wrapper?.type !== type) {
        return;
      }
      expect(wrapper.content.some((item) => item.type === "preservedInline")).toBe(true);

      expect(insideWrapper(serializeParagraph(paragraph), tag)).toContain(
        '<w:permStart w:id="7"/>',
      );
    });
  }

  test("a child kept by one wrapper is kept by every wrapper and every kind", () => {
    fc.assert(
      fc.property(
        fc.constantFrom(...WRAPPERS),
        fc.constantFrom(...UNMODELLED_CHILDREN),
        ({ tag }, child) => {
          const saved = serializeParagraph(parseParagraphXml(paragraphXml(tag, child)));
          expect(insideWrapper(saved, tag)).toContain(child);

          // Save, reopen, save: the second save is where a capture that only
          // replays and does not re-parse stops being a fixed point.
          const reparsed = serializeParagraph(
            parseParagraphXml(`<w:p xmlns:w="${W}">${saved.slice(saved.indexOf(">") + 1)}`),
          );
          expect(insideWrapper(reparsed, tag)).toContain(child);
        },
      ),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("the editor gives the child back inside the same wrapper", () => {
    fc.assert(
      fc.property(
        fc.constantFrom(...WRAPPERS),
        fc.constantFrom(...UNMODELLED_CHILDREN),
        ({ tag }, child) => {
          const reopened = throughTheEditor(parseParagraphXml(paragraphXml(tag, child)));
          expect(insideWrapper(serializeParagraph(reopened), tag)).toContain(child);
        },
      ),
      propertyConfig({ numRuns: 100 }),
    );
  });
});

describe("resolving the change resolves the markup inside it", () => {
  // `w:ins` and `w:del` are the pair a reviewer resolves; `w:moveFrom` and
  // `w:moveTo` resolve as they do, and the editor spells a move with the same
  // two marks.
  for (const { tag } of [WRAPPERS[0], WRAPPERS[1]]) {
    const kept = tag === "ins" ? "accept" : "reject";
    const removed = tag === "ins" ? "reject" : "accept";

    test(`${kept}ing a w:${tag} keeps the preserved child`, () => {
      const saved = resolveAll(
        parseParagraphXml(paragraphXml(tag, '<w:permStart w:id="7"/>')),
        kept,
      );
      const paragraph = saved.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") {
        throw new Error("Expected a paragraph");
      }
      // The wrapper is gone, the markup it held is not: resolving a revision
      // settles whether the content stays, not whether folio keeps markup it
      // was already carrying.
      expect(paragraph.content.some((item) => item.type === tagToType(tag))).toBe(false);
      expect(serializeParagraph(paragraph)).toContain('<w:permStart w:id="7"/>');
    });

    test(`${removed}ing a w:${tag} takes the preserved child with it`, () => {
      const saved = resolveAll(
        parseParagraphXml(paragraphXml(tag, '<w:permStart w:id="7"/>')),
        removed,
      );
      const paragraph = saved.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") {
        throw new Error("Expected a paragraph");
      }
      expect(serializeParagraph(paragraph)).not.toContain("<w:permStart");
    });
  }
});

function tagToType(tag: "ins" | "del"): "insertion" | "deletion" {
  return tag === "ins" ? "insertion" : "deletion";
}
