/**
 * The anonymization highlight owner: which ranges it finds, that a read
 * rescans only the paragraphs an edit replaced, that the occurrences follow
 * edits in a canonical session (the person's, an undo's, a peer's) without
 * touching the document's history, and that every painted story is covered.
 */

import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";
import type { Mark, Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  type TextPosition,
} from "@stll/docx-core/ops";

import { createCanonicalSession } from "../canonical/canonicalSession";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { readModel } from "../types/documentModel";
import {
  type AnonymizationShownMatch,
  type AnonymizationStorySource,
  createAnonymizationHighlights,
  type FolioAnonymizationTerm,
  NO_ANONYMIZATION_TERMS,
  slugAnonymizationLabel,
} from "./anonymizationHighlights";
import { legacyTextAddressAt } from "./legacyTextAddress";

// Minimal schema with one mark so we can model DOCX-style
// formatting boundaries (bold split mid-paragraph) — the
// situation that originally caused the cross-node-match bug.
const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      group: "block",
      content: "inline*",
      attrs: { paraId: { default: null } },
      toDOM: () => ["p", 0],
    },
    text: { group: "inline" },
  },
  marks: {
    bold: { toDOM: () => ["strong", 0] },
  },
});

type ParagraphSpec = string | (string | { text: string; bold: true })[];

const para = (...spec: ParagraphSpec[]): PMNode[] =>
  spec.map((entry, index) => {
    const attrs = { paraId: `P${String(index).padStart(7, "0")}` };
    if (typeof entry === "string") {
      return schema.node("paragraph", attrs, [schema.text(entry)]);
    }
    const bold = schema.marks.bold!.create();
    return schema.node(
      "paragraph",
      attrs,
      entry.map((part) =>
        typeof part === "string"
          ? schema.text(part)
          : schema.text(part.text, [bold] as Mark[]),
      ),
    );
  });

const bodySource = (doc: PMNode): AnonymizationStorySource => ({
  story: OP_STORIES.MAIN,
  doc,
  addressAt: (position) => legacyTextAddressAt(doc, position),
});

const makeState = (
  doc: PMNode,
  terms: readonly FolioAnonymizationTerm[] = [],
): { doc: PMNode; terms: readonly FolioAnonymizationTerm[] } => ({
  doc,
  terms,
});

const getMatches = (state: {
  doc: PMNode;
  terms: readonly FolioAnonymizationTerm[];
}): readonly AnonymizationShownMatch[] =>
  createAnonymizationHighlights().read(state.terms, () => [
    bodySource(state.doc),
  ]).shown;

const sliceFromTo = (doc: PMNode, from: number, to: number): string =>
  doc.textBetween(from, to, "");

describe("anonymization highlights: literal matching", () => {
  test("matches a contiguous term within a single text node", () => {
    const doc = schema.node("doc", null, para("Hello Pavel Novák there"));
    const state = makeState(doc, [
      { canonical: "Pavel Novák", label: "person" },
    ]);
    const matches = getMatches(state);
    expect(matches).toHaveLength(1);
    expect(sliceFromTo(doc, matches[0]!.from, matches[0]!.to)).toBe(
      "Pavel Novák",
    );
    expect(matches[0]!.label).toBe("person");
    expect(matches[0]!.canonical).toBe("Pavel Novák");
  });

  test("is case-insensitive", () => {
    const doc = schema.node("doc", null, para("hello PAVEL NOVÁK there"));
    const matches = getMatches(
      makeState(doc, [{ canonical: "Pavel Novák", label: "person" }]),
    );
    expect(matches).toHaveLength(1);
    expect(sliceFromTo(doc, matches[0]!.from, matches[0]!.to)).toBe(
      "PAVEL NOVÁK",
    );
  });

  test("respects word boundaries — no partial-word hits", () => {
    const doc = schema.node("doc", null, para("Johnsonville and John ate"));
    const matches = getMatches(
      makeState(doc, [{ canonical: "John", label: "person" }]),
    );
    expect(matches).toHaveLength(1);
    expect(sliceFromTo(doc, matches[0]!.from, matches[0]!.to)).toBe("John");
  });

  test("treats punctuation and whitespace as boundaries", () => {
    const doc = schema.node(
      "doc",
      null,
      para("Names: Pavel Novák, Karel Dvořák."),
    );
    const matches = getMatches(
      makeState(doc, [
        { canonical: "Pavel Novák", label: "person" },
        { canonical: "Karel Dvořák", label: "person" },
      ]),
    );
    expect(matches).toHaveLength(2);
    const surfaces = matches.map((m) => sliceFromTo(doc, m.from, m.to));
    expect(surfaces).toContain("Pavel Novák");
    expect(surfaces).toContain("Karel Dvořák");
  });

  test("normalises whitespace runs (NBSP, narrow no-break, multi-space)", () => {
    // U+00A0 NBSP between first/last name; matching pattern uses
    // a regular space — the matcher should treat any whitespace
    // run as equivalent.
    const doc = schema.node("doc", null, para("Pavel Novák"));
    const matches = getMatches(
      makeState(doc, [{ canonical: "Pavel Novák", label: "person" }]),
    );
    expect(matches).toHaveLength(1);
    expect(sliceFromTo(doc, matches[0]!.from, matches[0]!.to)).toBe(
      "Pavel Novák",
    );
  });

  test("escapes regex metacharacters in canonical terms", () => {
    const doc = schema.node(
      "doc",
      null,
      para("ENCAP s.r.o. and friends s_r_o text"),
    );
    const matches = getMatches(
      makeState(doc, [{ canonical: "ENCAP s.r.o.", label: "organization" }]),
    );
    expect(matches).toHaveLength(1);
    expect(sliceFromTo(doc, matches[0]!.from, matches[0]!.to)).toBe(
      "ENCAP s.r.o.",
    );
  });
});

describe("anonymization highlights: cross-node matching (regression)", () => {
  // DOCX often emits separate runs at every formatting boundary
  // (bold/italic/spacing change). The PM doc mirrors this — a
  // single visible phrase may live in two adjacent text nodes.
  // The matcher must join text within a textblock before running
  // regex so these phrases still match.
  test("matches a term spanning a bold/non-bold boundary", () => {
    const doc = schema.node(
      "doc",
      null,
      para([{ text: "Ing. ", bold: true }, "Pavel Novák, dat. nar."]),
    );
    const matches = getMatches(
      makeState(doc, [{ canonical: "Ing. Pavel Novák", label: "person" }]),
    );
    expect(matches).toHaveLength(1);
    expect(sliceFromTo(doc, matches[0]!.from, matches[0]!.to)).toBe(
      "Ing. Pavel Novák",
    );
  });

  test("does NOT join text across paragraph boundaries", () => {
    // Two separate paragraphs whose concatenation would form
    // the term; the matcher must not bridge across paragraph
    // boundaries even when the next paragraph starts mid-name.
    const doc = schema.node("doc", null, [...para("Pavel"), ...para(" Novák")]);
    const matches = getMatches(
      makeState(doc, [{ canonical: "Pavel Novák", label: "person" }]),
    );
    expect(matches).toHaveLength(0);
  });

  test("yields PM positions that map back to the original text", () => {
    // Two cross-node occurrences in separate paragraphs; both
    // ranges should round-trip through doc.textBetween to the
    // expected surface form.
    const doc = schema.node("doc", null, [
      ...para([{ text: "Ing. ", bold: true }, "Pavel Novák"]),
      ...para([{ text: "Ing. ", bold: true }, "Pavel Novák again"]),
    ]);
    const matches = getMatches(
      makeState(doc, [{ canonical: "Ing. Pavel Novák", label: "person" }]),
    );
    expect(matches).toHaveLength(2);
    for (const match of matches) {
      expect(sliceFromTo(doc, match.from, match.to)).toBe("Ing. Pavel Novák");
    }
  });
});

describe("anonymization highlights: variants and dedup", () => {
  test("matches any provided variant alongside the canonical", () => {
    const doc = schema.node(
      "doc",
      null,
      para("Acme Corp. and ACME CORPORATION are the same"),
    );
    const matches = getMatches(
      makeState(doc, [
        {
          canonical: "Acme Corporation",
          label: "organization",
          variants: ["Acme Corp.", "ACME CORPORATION"],
        },
      ]),
    );
    expect(matches).toHaveLength(2);
    const surfaces = matches.map((m) => sliceFromTo(doc, m.from, m.to));
    expect(surfaces).toContain("Acme Corp.");
    expect(surfaces).toContain("ACME CORPORATION");
    for (const m of matches) {
      expect(m.canonical).toBe("Acme Corporation");
    }
  });

  test("prefers the longer surface when a shorter variant is nested", () => {
    // Both "Pavel Novák" and the longer "Ing. Pavel
    // Novák" can match at the same locus — dedup keeps the
    // longer span so we never paint twice over the same letters.
    const doc = schema.node("doc", null, para("see Ing. Pavel Novák below"));
    const matches = getMatches(
      makeState(doc, [
        { canonical: "Ing. Pavel Novák", label: "person" },
        { canonical: "Pavel Novák", label: "person" },
      ]),
    );
    expect(matches).toHaveLength(1);
    expect(sliceFromTo(doc, matches[0]!.from, matches[0]!.to)).toBe(
      "Ing. Pavel Novák",
    );
  });

  test("skips empty canonicals and empty variants without crashing", () => {
    const doc = schema.node("doc", null, para("the real deal here"));
    const matches = getMatches(
      makeState(doc, [
        { canonical: "", label: "person" },
        { canonical: "real", label: "person", variants: ["", "fake"] },
      ]),
    );
    expect(matches).toHaveLength(1);
    expect(sliceFromTo(doc, matches[0]!.from, matches[0]!.to)).toBe("real");
  });
});

describe("anonymization highlights: reads", () => {
  const terms: readonly FolioAnonymizationTerm[] = [
    { canonical: "Alice", label: "person" },
  ];

  test("reads no story while there are no terms", () => {
    const highlights = createAnonymizationHighlights();
    const view = highlights.read(NO_ANONYMIZATION_TERMS, () =>
      panic("Read a story without terms"),
    );
    expect(view.shown).toEqual([]);
    expect(view.matches).toEqual([]);
  });

  test("reports occurrences as text positions in their paragraph", () => {
    const doc = schema.node("doc", null, para("Dear Alice,", "Alice wrote."));
    const { matches } = createAnonymizationHighlights().read(terms, () => [
      bodySource(doc),
    ]);
    expect(matches).toEqual([
      {
        start: { story: OP_STORIES.MAIN, blockId: "P0000000", offset: 5 },
        end: { story: OP_STORIES.MAIN, blockId: "P0000000", offset: 10 },
        label: "person",
        canonical: "Alice",
      },
      {
        start: { story: OP_STORIES.MAIN, blockId: "P0000001", offset: 0 },
        end: { story: OP_STORIES.MAIN, blockId: "P0000001", offset: 5 },
        label: "person",
        canonical: "Alice",
      },
    ]);
  });

  test("rescans only the paragraphs an edit replaced", () => {
    const highlights = createAnonymizationHighlights();
    let addressed = 0;
    const source = (doc: PMNode): AnonymizationStorySource => ({
      ...bodySource(doc),
      addressAt: (position) => {
        addressed += 1;
        return legacyTextAddressAt(doc, position);
      },
    });
    const before = EditorState.create({
      doc: schema.node("doc", null, para("Alice one.", "Two.", "Alice three.")),
    });
    const first = highlights.read(terms, () => [source(before.doc)]);
    expect(addressed).toBe(4);
    // Reading again rescans nothing and keeps the occurrences.
    expect(highlights.read(terms, () => [source(before.doc)]).matches).toBe(
      first.matches,
    );
    expect(addressed).toBe(4);
    // Typing in the middle paragraph rescans only it; the occurrences of the
    // others did not move within their paragraphs, so the list stays.
    const typed = before.apply(before.tr.insertText("x", 13));
    const second = highlights.read(terms, () => [source(typed.doc)]);
    expect(addressed).toBe(4);
    expect(second.matches).toBe(first.matches);
    expect(second.shown.map(({ from }) => from)).toEqual([1, 20]);
    // Typing before an occurrence moves it within its paragraph.
    const moved = typed.apply(typed.tr.insertText("Hi ", 1));
    const third = highlights.read(terms, () => [source(moved.doc)]);
    expect(addressed).toBe(6);
    expect(third.matches).not.toBe(first.matches);
    expect(third.matches.map(({ start }) => start.offset)).toEqual([3, 0]);
    // New terms rescan everything.
    highlights.read([...terms], () => [source(moved.doc)]);
    expect(addressed).toBe(10);
  });

  test("keeps each story's occurrences with that story", () => {
    const body = schema.node("doc", null, para("Alice in the body."));
    const header = schema.node("doc", null, para("Alice in the header."));
    const story = { kind: "header", rId: "rId7" } as const;
    const { shown, matches } = createAnonymizationHighlights().read(
      terms,
      () => [
        bodySource(body),
        {
          story,
          doc: header,
          addressAt: (position) => {
            const address = legacyTextAddressAt(header, position);
            return address === null ? null : { ...address, story };
          },
        },
      ],
    );
    expect(shown.map((match) => match.story)).toEqual([OP_STORIES.MAIN, story]);
    expect(matches.map(({ start }) => start.story)).toEqual([
      OP_STORIES.MAIN,
      story,
    ]);
  });
});

/** A document from Markdown whose paragraphs carry ids, as a parsed DOCX's do. */
const withParagraphIds = (markdown: string) => {
  const document = fromMarkdown(markdown);
  for (const [index, block] of document.package.document.content.entries())
    if (block.type === "paragraph")
      block.paraId = (index + 1).toString(16).toUpperCase().padStart(8, "0");
  return document;
};

describe("anonymization highlights: canonical session", () => {
  const terms: readonly FolioAnonymizationTerm[] = [
    { canonical: "Alice", label: "person" },
  ];

  const openCanonical = () => {
    const session = createCanonicalSession(
      withParagraphIds("Dear Alice, hello.\n\nAlice wrote back."),
    ).unwrap();
    let state = EditorState.create({ doc: session.projection.doc });
    const highlights = createAnonymizationHighlights();
    const read = () =>
      highlights.read(terms, () => [
        {
          story: OP_STORIES.MAIN,
          doc: session.projection.doc,
          addressAt: (position) =>
            session.projection.addressAt(position).unwrapOr(null),
        },
      ]);
    const shownText = () =>
      read().shown.map(({ from, to }) =>
        session.projection.doc.textBetween(from, to),
      );
    const ids = readModel(session.document).package.document.content.flatMap(
      (block) =>
        block.type === "paragraph" && block.paraId !== undefined
          ? [block.paraId]
          : [],
    );
    const at = (paragraph: number, offset: number): TextPosition => ({
      story: OP_STORIES.MAIN,
      blockId: ids.at(paragraph) ?? panic("No paragraph"),
      offset,
    });
    const insert = (position: TextPosition, text: string): DocumentOp => ({
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: position,
      text,
      runProps: INHERIT_RUN_PROPS,
    });
    const typeLocally = (position: TextPosition, text: string) => {
      const commit = session
        .prepareOperations(state, [insert(position, text)])
        .unwrap();
      state = state.apply(commit.transaction);
      expect(commit.publish().isOk()).toBe(true);
    };
    let revision = 0;
    const typeAsPeer = (position: TextPosition, text: string) => {
      revision += 1;
      const commit = session
        .prepareRemote(state, {
          type: "rebased",
          batch: {
            schema: DOCUMENT_OP_SCHEMA_VERSION,
            opId: `peer-${String(revision)}`,
            actor: "peer",
            baseRev: revision - 1,
            revision,
            ops: [insert(position, text)],
          },
        })
        .unwrap();
      state = state.apply(commit.transaction);
      expect(commit.publish().isOk()).toBe(true);
    };
    return { session, read, shownText, at, typeLocally, typeAsPeer };
  };

  test("occurrences follow edits before, inside and after them", () => {
    const canonical = openCanonical();
    expect(canonical.shownText()).toEqual(["Alice", "Alice"]);
    expect(canonical.read().matches.map(({ start }) => start.offset)).toEqual([
      5, 0,
    ]);
    // Before: the occurrence moves with its text.
    canonical.typeLocally(canonical.at(0, 0), "Oh ");
    expect(canonical.read().matches.map(({ start }) => start.offset)).toEqual([
      8, 0,
    ]);
    expect(canonical.shownText()).toEqual(["Alice", "Alice"]);
    // After: nothing moves.
    canonical.typeLocally(canonical.at(1, 16), "!");
    expect(canonical.read().matches.map(({ start }) => start.offset)).toEqual([
      8, 0,
    ]);
    // Inside: the word is no longer the term.
    canonical.typeLocally(canonical.at(1, 2), "x");
    expect(canonical.shownText()).toEqual(["Alice"]);
  });

  test("a peer's edit moves occurrences in place", () => {
    const canonical = openCanonical();
    canonical.typeAsPeer(canonical.at(1, 0), "Then ");
    expect(canonical.read().matches.map(({ start }) => start.offset)).toEqual([
      5, 5,
    ]);
    expect(canonical.shownText()).toEqual(["Alice", "Alice"]);
  });

  test("reading highlights changes neither the document nor its history", () => {
    const canonical = openCanonical();
    const version = canonical.session.version;
    const document = canonical.session.document;
    canonical.read();
    canonical.read();
    expect(canonical.session.version).toBe(version);
    expect(canonical.session.document).toBe(document);
    expect(canonical.session.canUndo).toBe(false);
  });
});

describe("slugAnonymizationLabel", () => {
  test("lowercases and replaces non-alphanumeric runs with a single dash", () => {
    expect(slugAnonymizationLabel("Person Name")).toBe("person-name");
    expect(slugAnonymizationLabel("ORG/Unit")).toBe("org-unit");
  });
});
