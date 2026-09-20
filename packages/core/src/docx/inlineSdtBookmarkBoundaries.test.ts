/**
 * A bookmark marker inside an inline content control stays inside it.
 *
 * `CT_SdtContentRun` and `CT_RunTrackChange` both reach `w:bookmarkStart` and
 * `w:bookmarkEnd` through `EG_RunLevelElts > EG_RangeMarkupElements`, so
 * `w:sdt > w:sdtContent > w:bookmarkStart` and `w:ins > w:bookmarkStart` are
 * both markup Word writes. folio lifted the marker out of the control to a
 * sibling of it, and that is not a re-spelling: a bookmark whose extent was
 * the control's content came back spanning the control and whatever followed
 * it, so a `REF` field or a link to the bookmark covered the wrong range.
 *
 * Position is the whole assertion, so every check is about which parent the
 * marker comes back under and where it sits among that parent's content. A
 * marker written beside the control reads as surviving to anything that only
 * asks whether the markup is somewhere in the part, which is why the container
 * census cannot see this and these properties can.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState, type Plugin } from "prosemirror-state";

import { propertyConfig } from "../../../../test/property-testing";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { BookmarkBoundaryExtension } from "../prosemirror/extensions/nodes/BookmarkBoundaryExtension";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph, ParagraphContent } from "../types/document";

import { parseParagraph } from "./paragraphParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * The two run-level wrappers that admit a range marker.
 *
 * `innerTag` is the element the saved markup puts the content in: a control
 * holds its content in `w:sdtContent`, a revision directly in `w:ins`.
 */
const WRAPPERS = [
  {
    label: "w:sdt",
    type: "inlineSdt",
    open: '<w:sdt><w:sdtPr><w:tag w:val="bound"/></w:sdtPr><w:sdtContent>',
    close: "</w:sdtContent></w:sdt>",
    innerTag: "sdtContent",
  },
  {
    label: "w:ins",
    type: "insertion",
    open: '<w:ins w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z">',
    close: "</w:ins>",
    innerTag: "ins",
  },
] as const;

type Wrapper = (typeof WRAPPERS)[number];

const MARKERS = [
  { label: "w:bookmarkStart", type: "bookmarkStart" },
  { label: "w:bookmarkEnd", type: "bookmarkEnd" },
] as const;

type Marker = (typeof MARKERS)[number];

const BOOKMARK_ID = 9;
const START_XML = `<w:bookmarkStart w:id="${BOOKMARK_ID}" w:name="anchor"/>`;
const END_XML = `<w:bookmarkEnd w:id="${BOOKMARK_ID}"/>`;

const markerXml = (marker: Marker): string =>
  marker.type === "bookmarkStart" ? START_XML : END_XML;

const runXml = (text: string): string => `<w:r><w:t>${text}</w:t></w:r>`;

/** The text each run in a fixture of `runs` runs carries, in order. */
const runTexts = (runs: number): string[] =>
  Array.from({ length: runs }, (_, index) => `r${index}`);

/**
 * A wrapper holding `runs` runs with the marker at `ordinal` among them.
 *
 * The marker's partner sits outside the wrapper so the range is balanced: an
 * unpaired boundary is an orphan the editor's integrity pass deletes, and a
 * fixture built from one would measure that pass rather than the wrapper.
 */
const paragraphXml = (wrapper: Wrapper, marker: Marker, runs: number, ordinal: number): string => {
  const children = runTexts(runs).map(runXml);
  children.splice(ordinal, 0, markerXml(marker));
  const partner = marker.type === "bookmarkStart" ? END_XML : START_XML;
  return (
    `<w:p xmlns:w="${W}">` +
    (marker.type === "bookmarkStart" ? "" : partner) +
    `${wrapper.open}${children.join("")}${wrapper.close}` +
    (marker.type === "bookmarkStart" ? partner : "") +
    "</w:p>"
  );
};

const parseParagraphXml = (xml: string): Paragraph => {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("Failed to parse the paragraph fixture");
  }
  return parseParagraph(root, null, null, null, null, null);
};

/** What the wrapper holds in the saved markup, or `null` when it is gone. */
const insideWrapper = (savedXml: string, innerTag: string): string | null => {
  const open = savedXml.indexOf(`<w:${innerTag}`);
  const close = savedXml.indexOf(`</w:${innerTag}>`, open);
  if (open === -1 || close === -1) {
    return null;
  }
  return savedXml.slice(savedXml.indexOf(">", open) + 1, close);
};

const textOf = (xml: string): string =>
  [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu)].map(([, text]) => text ?? "").join("");

/**
 * The text on each side of the marker inside the wrapper.
 *
 * Adjacent runs with the same formatting merge into one on the way back from
 * the editor, so the ordinal among a list of items is not stable across the
 * round trip and the text on either side of the marker is. Both say the same
 * thing about position: the marker is between the same two characters.
 */
type Split = { before: string; after: string };

const splitSavedMarkup = (savedXml: string, wrapper: Wrapper, marker: Marker): Split | null => {
  const inside = insideWrapper(savedXml, wrapper.innerTag);
  if (inside === null) {
    return null;
  }
  const at = inside.indexOf(`<w:${marker.type} `);
  if (at === -1) {
    return null;
  }
  return { before: textOf(inside.slice(0, at)), after: textOf(inside.slice(at)) };
};

/** The same split over the model, or `null` when the wrapper no longer holds it. */
const splitModel = (paragraph: Paragraph, wrapper: Wrapper, marker: Marker): Split | null => {
  const held = paragraph.content.find((item) => item.type === wrapper.type);
  if (held?.type !== wrapper.type) {
    return null;
  }
  const at = held.content.findIndex((item) => item.type === marker.type);
  if (at === -1) {
    return null;
  }
  return {
    before: modelText(held.content.slice(0, at)),
    after: modelText(held.content.slice(at)),
  };
};

const modelText = (content: readonly ParagraphContent[]): string => {
  let text = "";
  for (const item of content) {
    if (item.type === "run") {
      for (const child of item.content) {
        if (child.type === "text") {
          text += child.text;
        }
      }
      continue;
    }
    if (item.type === "inlineSdt" || item.type === "insertion" || item.type === "deletion") {
      text += modelText(item.content);
    }
  }
  return text;
};

const wrapParagraph = (paragraph: Paragraph): Document => ({
  package: { document: { content: [paragraph] } },
});

/** The paragraph as it comes back from the editor with nothing edited. */
const throughTheEditor = (paragraph: Paragraph): Paragraph => {
  const input = wrapParagraph(paragraph);
  const first = fromProseDoc(toProseDoc(input), input).package.document.content.at(0);
  if (first?.type !== "paragraph") {
    throw new Error("Expected the round trip to give a paragraph back");
  }
  return first;
};

const expectedSplit = (runs: number, ordinal: number): Split => ({
  before: runTexts(ordinal).join(""),
  after: runTexts(runs).slice(ordinal).join(""),
});

describe("a bookmark marker inside an inline wrapper keeps its parent and its place", () => {
  for (const wrapper of WRAPPERS) {
    for (const marker of MARKERS) {
      test(`${wrapper.label} keeps a ${marker.label} inside itself across a save`, () => {
        fc.assert(
          fc.property(
            fc
              .integer({ min: 0, max: 3 })
              .chain((runs) => fc.tuple(fc.constant(runs), fc.integer({ min: 0, max: runs }))),
            ([runs, ordinal]) => {
              const paragraph = parseParagraphXml(paragraphXml(wrapper, marker, runs, ordinal));
              expect(splitModel(paragraph, wrapper, marker)).toEqual(expectedSplit(runs, ordinal));
              expect(splitSavedMarkup(serializeParagraph(paragraph), wrapper, marker)).toEqual(
                expectedSplit(runs, ordinal),
              );
            },
          ),
          propertyConfig(),
        );
      });

      test(`${wrapper.label} keeps a ${marker.label} inside itself through the editor`, () => {
        fc.assert(
          fc.property(
            fc
              .integer({ min: 0, max: 3 })
              .chain((runs) => fc.tuple(fc.constant(runs), fc.integer({ min: 0, max: runs }))),
            ([runs, ordinal]) => {
              const rebuilt = throughTheEditor(
                parseParagraphXml(paragraphXml(wrapper, marker, runs, ordinal)),
              );
              expect(splitModel(rebuilt, wrapper, marker)).toEqual(expectedSplit(runs, ordinal));
            },
          ),
          propertyConfig(),
        );
      });
    }
  }
});

/**
 * A range that opens inside the control and closes outside it.
 *
 * Both halves are converted at their own structural position, so the pairing
 * pass has to look inside the control to find the start. If it does not, the
 * start is an orphan and the integrity plugin deletes it on the next edit,
 * taking the bookmark with it.
 */
const CROSSING_PARAGRAPH =
  `<w:p xmlns:w="${W}">` +
  '<w:sdt><w:sdtPr><w:tag w:val="bound"/></w:sdtPr><w:sdtContent>' +
  START_XML +
  runXml("inside") +
  "</w:sdtContent></w:sdt>" +
  runXml("after") +
  END_XML +
  "</w:p>";

const boundaryIntegrityPlugin = (): Plugin => {
  const plugin = BookmarkBoundaryExtension().onSchemaReady({ schema }).plugins?.at(0);
  if (!plugin) {
    throw new Error("BookmarkBoundaryExtension must enforce boundary integrity");
  }
  return plugin;
};

const controlContentOf = (paragraph: Paragraph): readonly ParagraphContent[] => {
  const control = paragraph.content.find((item) => item.type === "inlineSdt");
  if (control?.type !== "inlineSdt") {
    throw new Error("The rebuilt paragraph lost its content control");
  }
  return control.content;
};

describe("a bookmark that starts inside the control and ends outside it", () => {
  test("keeps both halves through the editor", () => {
    const rebuilt = throughTheEditor(parseParagraphXml(CROSSING_PARAGRAPH));
    expect(controlContentOf(rebuilt).at(0)?.type).toBe("bookmarkStart");
    expect(rebuilt.content.at(-1)?.type).toBe("bookmarkEnd");
  });

  test("survives an edit rather than being deleted as an orphan", () => {
    const input = wrapParagraph(parseParagraphXml(CROSSING_PARAGRAPH));
    const state = EditorState.create({
      doc: toProseDoc(input),
      plugins: [boundaryIntegrityPlugin()],
    });
    const applied = state.applyTransaction(state.tr.insertText("!", 1));

    const first = fromProseDoc(applied.state.doc, input).package.document.content.at(0);
    if (first?.type !== "paragraph") {
      throw new Error("Expected a paragraph");
    }
    expect(controlContentOf(first).some((item) => item.type === "bookmarkStart")).toBe(true);
    expect(first.content.some((item) => item.type === "bookmarkEnd")).toBe(true);
  });
});

/**
 * W-5's hoist with a marker inside the control.
 *
 * A revision that covers every child of a control is written back around it,
 * and the marker is one of those children: the hoist has to carry it along
 * rather than refuse because of it or leave it behind.
 */
describe("a revision over a whole control that holds a marker", () => {
  test("hoists to w:ins > w:sdt with the marker still inside the control", () => {
    const paragraph = parseParagraphXml(
      `<w:p xmlns:w="${W}">` +
        '<w:sdt><w:sdtPr><w:tag w:val="bound"/></w:sdtPr><w:sdtContent>' +
        '<w:ins w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z">' +
        START_XML +
        runXml("inside") +
        "</w:ins></w:sdtContent></w:sdt>" +
        END_XML +
        "</w:p>",
    );

    const rebuilt = throughTheEditor(paragraph);
    const revision = rebuilt.content.at(0);
    if (revision?.type !== "insertion") {
      throw new Error(`Expected an insertion, got ${revision?.type ?? "nothing"}`);
    }
    const control = revision.content.at(0);
    if (control?.type !== "inlineSdt") {
      throw new Error("The revision should hold the control");
    }
    expect(control.content.map((item) => item.type)).toEqual(["bookmarkStart", "run"]);
  });
});
