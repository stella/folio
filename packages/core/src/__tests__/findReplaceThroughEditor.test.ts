/**
 * Find and replace writes through the body editor's command owner, the same
 * way on both sessions: the canonical session journals it (tracked while
 * suggesting, one undo step) and the default session applies the command's
 * transaction, which its suggestion plugin tracks.
 *
 * For each engine, document shape and operation, the replacement made while
 * editing reads as the expected text; the one made while suggesting, with
 * every change accepted, reads the same, and with every change rejected reads
 * as the source; one undo restores the source. Matches replacement policy
 * refuses are skipped and leave the document as it was.
 */

import { describe, expect, test } from "bun:test";
import { undo } from "prosemirror-history";
import type { Node as PMNode } from "prosemirror-model";
import { EditorView } from "prosemirror-view";

import { createCanonicalEditorHarness } from "../../../../test/canonicalEditorHarness";
import { acquireHarnessDom } from "../../../../test/harnessDom";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { resolveFindMatchRange } from "../prosemirror/findReplaceSelection";
import {
  replaceFindMatch,
  replaceFindMatches,
} from "../prosemirror/replaceFindMatches";
import type { BlockContent, Document } from "../types/document";
import {
  createDefaultFindOptions,
  findInDocument,
  type FindMatch,
} from "../utils/findReplace";
import { createHarnessState, type EditorMode } from "./editorHarness";

const REPLACEMENT = "folioReplacement";

const SHAPES = [
  {
    name: "plain paragraphs",
    markdown: "alpha editor beta\n\ngamma editor delta",
    query: "editor",
  },
  {
    name: "a match across a run boundary",
    markdown: "alpha **edi**tor beta\n\ngamma editor",
    query: "editor",
  },
  {
    name: "matches in table cells",
    markdown:
      "intro text\n\n| editor one | two |\n|---|---|\n| three | four editor |",
    query: "editor",
  },
] as const;

const OPERATIONS = ["current", "all"] as const;
type Operation = (typeof OPERATIONS)[number];

/** Every paragraph carries an identity, as a parsed package does. */
const withParagraphIds = (document: Document): Document => {
  let next = 1;
  const visit = (blocks: readonly BlockContent[]) => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        block.paraId = (next++).toString(16).padStart(8, "0").toUpperCase();
      } else if (block.type === "table") {
        for (const row of block.rows)
          for (const cell of row.cells) visit(cell.content);
      }
    }
  };
  visit(document.package.document.content);
  return document;
};

/** The body's paragraphs as they read with one kind of revision hidden. */
const readLines = (doc: PMNode, hidden: "deletion" | "insertion"): string[] => {
  const lines: string[] = [];
  doc.descendants((node) => {
    if (!node.isTextblock) return true;
    let text = "";
    node.descendants((child) => {
      if (
        child.isText &&
        !child.marks.some((mark) => mark.type.name === hidden)
      )
        text += child.text ?? "";
      return true;
    });
    lines.push(text);
    return false;
  });
  return lines;
};

const expectedLines = (
  source: readonly string[],
  query: string,
  operation: Operation,
): string[] => {
  if (operation === "all")
    return source.map((line) => line.replaceAll(query, REPLACEMENT));
  const first = source.findIndex((line) => line.includes(query));
  return source.map((line, index) =>
    index === first ? line.replace(query, REPLACEMENT) : line,
  );
};

type Engine = {
  view: EditorView;
  undo: () => boolean;
  dispose: () => void;
};

const ENGINES = {
  canonical: (source: Document, mode: EditorMode): Engine => {
    const driver = createCanonicalEditorHarness(source, mode);
    return {
      view: driver.view,
      undo: () => driver.history.undo(),
      dispose: () => {
        expect(driver.refusals).toEqual([]);
        driver.dispose();
      },
    };
  },
  legacy: (source: Document, mode: EditorMode): Engine => {
    const releaseDom = acquireHarnessDom();
    const host = document.createElement("div");
    document.body.append(host);
    const view = new EditorView(host, {
      state: createHarnessState(source, mode),
    });
    return {
      view,
      undo: () => undo(view.state, (transaction) => view.dispatch(transaction)),
      dispose: () => {
        view.destroy();
        host.remove();
        releaseDom();
      },
    };
  },
} as const;

const replace = (
  engine: Engine,
  matches: readonly FindMatch[],
  operation: Operation,
) => {
  if (operation === "all")
    return replaceFindMatches(engine.view, matches, REPLACEMENT);
  const first = matches.at(0);
  if (first === undefined) throw new TypeError("The shape has no match.");
  return {
    replacedCount: replaceFindMatch(engine.view, first, REPLACEMENT) ? 1 : 0,
    skippedCount: 0,
  };
};

describe("find and replace through the editor", () => {
  test.each(
    (["canonical", "legacy"] as const).flatMap((engine) =>
      SHAPES.flatMap((shape) =>
        OPERATIONS.map(
          (operation) => [engine, shape.name, operation, shape] as const,
        ),
      ),
    ),
  )("%s: %s, replace %s", (engineName, _name, operation, shape) => {
    const source = withParagraphIds(fromMarkdown(shape.markdown));
    const matches = findInDocument(
      source,
      shape.query,
      createDefaultFindOptions(),
    );
    expect(matches.length).toBeGreaterThan(1);
    const expectedCount = operation === "all" ? matches.length : 1;

    const editing = ENGINES[engineName](source, "editing");
    try {
      const original = readLines(editing.view.state.doc, "deletion");
      const expected = expectedLines(original, shape.query, operation);
      expect(replace(editing, matches, operation)).toEqual({
        replacedCount: expectedCount,
        skippedCount: 0,
      });
      expect(readLines(editing.view.state.doc, "deletion")).toEqual(expected);
      expect(readLines(editing.view.state.doc, "insertion")).toEqual(expected);
      expect(editing.undo()).toBe(true);
      expect(readLines(editing.view.state.doc, "deletion")).toEqual(original);

      const suggesting = ENGINES[engineName](source, "suggesting");
      try {
        expect(replace(suggesting, matches, operation)).toEqual({
          replacedCount: expectedCount,
          skippedCount: 0,
        });
        // Accepted, the suggestion reads as the direct edit; rejected, as
        // the source.
        expect(readLines(suggesting.view.state.doc, "deletion")).toEqual(
          expected,
        );
        expect(readLines(suggesting.view.state.doc, "insertion")).toEqual(
          original,
        );
        expect(suggesting.undo()).toBe(true);
        expect(readLines(suggesting.view.state.doc, "deletion")).toEqual(
          original,
        );
        expect(readLines(suggesting.view.state.doc, "insertion")).toEqual(
          original,
        );
      } finally {
        suggesting.dispose();
      }
    } finally {
      editing.dispose();
    }
  });

  test.each(["canonical", "legacy"] as const)(
    "%s: matches policy refuses are skipped and change nothing",
    (engineName) => {
      const source = withParagraphIds(fromMarkdown("alpha editor beta"));
      const [body] = findInDocument(
        source,
        "editor",
        createDefaultFindOptions(),
      );
      if (body === undefined) throw new TypeError("The body match is missing.");
      const refused: FindMatch[] = [
        { ...body, story: { kind: "footnote", id: 1 } },
        { ...body, story: { kind: "header", rId: "rId1" } },
        { ...body, story: "main", commentId: 7 },
        { ...body, replacement: { status: "blocked", reason: "field-result" } },
      ];
      const engine = ENGINES[engineName](source, "editing");
      try {
        const before = engine.view.state.doc;
        for (const match of refused)
          expect(replaceFindMatch(engine.view, match, REPLACEMENT)).toBe(false);
        expect(replaceFindMatches(engine.view, refused, REPLACEMENT)).toEqual({
          replacedCount: 0,
          skippedCount: refused.length,
        });
        expect(engine.view.state.doc).toBe(before);
        expect(
          replaceFindMatches(engine.view, [...refused, body], REPLACEMENT),
        ).toEqual({ replacedCount: 1, skippedCount: refused.length });
        expect(readLines(engine.view.state.doc, "deletion")).toEqual([
          `alpha ${REPLACEMENT} beta`,
        ]);
      } finally {
        engine.dispose();
      }
    },
  );

  test.each(["canonical", "legacy"] as const)(
    "%s: a match that already reads the replacement counts without an edit",
    (engineName) => {
      const source = withParagraphIds(fromMarkdown("alpha editor beta"));
      const matches = findInDocument(
        source,
        "editor",
        createDefaultFindOptions(),
      );
      const engine = ENGINES[engineName](source, "editing");
      try {
        const before = engine.view.state.doc;
        expect(replaceFindMatches(engine.view, matches, "editor")).toEqual({
          replacedCount: 1,
          skippedCount: 0,
        });
        expect(engine.view.state.doc).toBe(before);
      } finally {
        engine.dispose();
      }
    },
  );

  // The check bites: a replacement that skips the command owner, as a raw
  // transaction on the view, writes nothing on the canonical session, so the
  // expectations above fail for it.
  test("canonical: a raw replacement transaction leaves the body unchanged", () => {
    const source = withParagraphIds(fromMarkdown("alpha editor beta"));
    const [match] = findInDocument(
      source,
      "editor",
      createDefaultFindOptions(),
    );
    if (match === undefined) throw new TypeError("The body match is missing.");
    const driver = createCanonicalEditorHarness(source, "editing");
    try {
      const range = resolveFindMatchRange(driver.view.state.doc, match);
      if (range === null) throw new TypeError("The match did not resolve.");
      const before = readLines(driver.view.state.doc, "deletion");
      driver.view.dispatch(
        driver.view.state.tr.insertText(REPLACEMENT, range.from, range.to),
      );
      expect(readLines(driver.view.state.doc, "deletion")).toEqual(before);
      expect(before).not.toEqual([`alpha ${REPLACEMENT} beta`]);
    } finally {
      driver.dispose();
    }
  });
});
