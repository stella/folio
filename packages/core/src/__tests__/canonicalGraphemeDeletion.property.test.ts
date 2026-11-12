/**
 * Backspace and Delete at a grapheme boundary remove exactly one whole
 * grapheme cluster in the canonical session: ZWJ sequences, flags, skin tones,
 * keycaps and letters with combining marks, also beside and across tracked
 * changes. Editing removes the cluster; suggesting marks it deleted, so
 * accepting reads as editing and rejecting restores the source. No saved view
 * ever holds a lone surrogate or a piece of a cluster the source did not have.
 */

import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";

import {
  createCanonicalEditorHarness,
  resolveCanonicalHarnessDocument,
  saveCanonicalHarnessDocument,
} from "../../../../test/canonicalEditorHarness";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type { ParagraphContent } from "../types/content";
import type { Document } from "../types/document";
import { type EditorMode, HARNESS_AUTHOR } from "./editorHarness";

const segmenter = new Intl.Segmenter(undefined, { granularity: "grapheme" });
const clustersOf = (text: string): string[] =>
  Array.from(segmenter.segment(text), ({ segment }) => segment);
/** Offsets of `text` that stand between two clusters, ends included. */
const boundariesOf = (text: string): Set<number> => {
  const boundaries = new Set([text.length]);
  for (const { index } of segmenter.segment(text)) boundaries.add(index);
  return boundaries;
};

const MULTI_CODE_POINT_CLUSTERS = [
  // ZWJ family.
  "\u{1F468}\u{200D}\u{1F469}\u{200D}\u{1F467}\u{200D}\u{1F466}",
  // ZWJ profession with a skin tone.
  "\u{1F9D1}\u{1F3FD}\u{200D}\u{1F4BB}",
  // ZWJ person with a skin tone and a hair component.
  "\u{1F9D1}\u{1F3FB}\u{200D}\u{1F9B0}",
  // Regional indicator flags.
  "\u{1F1E8}\u{1F1FF}",
  "\u{1F1F8}\u{1F1F0}",
  // A skin tone modifier.
  "\u{1F44D}\u{1F3FC}",
  // Keycaps: digit or sign, U+FE0F, U+20E3.
  "1\u{FE0F}\u{20E3}",
  "#\u{FE0F}\u{20E3}",
  // Letters with one or more combining marks.
  "e\u{301}",
  "a\u{308}\u{304}",
  "n\u{303}\u{323}\u{307}",
] as const;
const CLUSTERS = [
  ...MULTI_CODE_POINT_CLUSTERS,
  "a",
  "Z",
  " ",
  // A single code point outside the Basic Multilingual Plane.
  "\u{1F600}",
] as const;

const REVISIONS = [
  "plain",
  "insertOwn",
  "insertOther",
  "deleteOwn",
  "deleteOther",
] as const;
type Revision = (typeof REVISIONS)[number];
type Segment = { revision: Revision; text: string };

/** Ordered pairs of distinct visible revisions a split cluster straddles. */
const SPLIT_SIDES = [
  ["plain", "insertOwn"],
  ["plain", "insertOther"],
  ["insertOwn", "plain"],
  ["insertOther", "plain"],
  ["insertOwn", "insertOther"],
  ["insertOther", "insertOwn"],
] as const satisfies readonly (readonly [Revision, Revision])[];

const OTHER_AUTHOR = "Other reviewer";

const documentFor = (segments: readonly Segment[]): Document => {
  let id = 40;
  const run = (text: string): ParagraphContent => ({
    type: "run",
    content: [{ type: "text", text }],
  });
  const content = segments.map(({ revision, text }): ParagraphContent => {
    if (revision === "plain") return run(text);
    id += 1;
    const info = {
      id,
      author: revision.endsWith("Own") ? HARNESS_AUTHOR : OTHER_AUTHOR,
      date: "2026-10-07T00:00:00Z",
    };
    return revision.startsWith("insert")
      ? { type: "insertion", info, content: [run(text)] }
      : { type: "deletion", info, content: [run(text)] };
  });
  return {
    package: {
      document: {
        content: [{ type: "paragraph", paraId: "12345678", content }],
      },
    },
  } satisfies Document;
};

type Views = {
  /** The text a reader sees now: everything but tracked deletions. */
  current: string;
  /** Tracked deleted text, in document order. */
  deleted: string;
  /** Every text node as saved. */
  nodes: string[];
};

const viewsOf = (model: Document): Views => {
  const views: Views = { current: "", deleted: "", nodes: [] };
  const visit = (
    items: readonly ParagraphContent[],
    deleted: boolean,
  ): void => {
    for (const item of items) {
      if (item.type === "run") {
        for (const part of item.content) {
          if (part.type !== "text")
            panic(`Unexpected run content ${part.type}.`);
          views.nodes.push(part.text);
          if (deleted) views.deleted += part.text;
          else views.current += part.text;
        }
      } else if (item.type === "insertion") visit(item.content, deleted);
      else if (item.type === "deletion") visit(item.content, true);
      else if (item.type !== "bookmarkStart" && item.type !== "bookmarkEnd")
        panic(`Unexpected paragraph content ${item.type}.`);
    }
  };
  const blocks = model.package.document.content;
  const paragraph = blocks.at(0);
  if (blocks.length !== 1 || paragraph?.type !== "paragraph")
    return panic("The grapheme fixture must stay one paragraph.");
  visit(paragraph.content, false);
  return views;
};

type Caret = { position: number; offset: number };

/**
 * Every caret position of the projected paragraph that stands on a cluster
 * boundary of the visible text, with its visible offset. Inside struck text
 * the caret also stands on a boundary of that text.
 */
const caretsOf = (doc: PMNode): { visible: string; carets: Caret[] } => {
  const paragraphs: { node: PMNode; position: number }[] = [];
  doc.descendants((node, position) => {
    if (node.type.name === "paragraph") paragraphs.push({ node, position });
    return paragraphs.length === 0;
  });
  const paragraph =
    paragraphs.at(0) ?? panic("The projection lost its paragraph.");
  let visible = "";
  const candidates = new Map<number, number>();
  const start = paragraph.position + 1;
  paragraph.node.forEach((child, childOffset) => {
    if (!child.isText || child.text === undefined)
      panic(`Unexpected inline ${child.type.name}.`);
    const text = child.text;
    const deleted = child.marks.some(({ type }) => type.name === "deletion");
    const own = boundariesOf(text);
    for (let unit = 0; unit <= text.length; unit += 1) {
      if (deleted && !own.has(unit)) continue;
      const position = start + childOffset + unit;
      if (!candidates.has(position))
        candidates.set(position, visible.length + (deleted ? 0 : unit));
    }
    if (!deleted) visible += text;
  });
  if (candidates.size === 0) candidates.set(start, 0);
  const boundaries = boundariesOf(visible);
  const carets = [...candidates]
    .filter(([, offset]) => boundaries.has(offset))
    .map(([position, offset]) => ({ position, offset }));
  return { visible, carets };
};

/** The cluster a key at `offset` must remove from `visible`. */
const clusterAt = (
  visible: string,
  offset: number,
  key: "Backspace" | "Delete",
): { from: number; to: number } => {
  for (const { index, segment } of segmenter.segment(visible)) {
    const to = index + segment.length;
    if (key === "Backspace" ? to === offset : index === offset)
      return { from: index, to };
  }
  return panic("No cluster beside the caret.", { offset });
};

/** A lone surrogate reads as a surrogate code point under the `u` flag. */
const LONE_SURROGATE = /\p{Cs}/u;

const expectWholeClusters = (
  label: string,
  text: string,
  alphabet: ReadonlySet<string>,
) => {
  expect(LONE_SURROGATE.test(text), `${label} holds a lone surrogate`).toBe(
    false,
  );
  for (const cluster of clustersOf(text))
    expect(
      alphabet.has(cluster),
      `${label} holds ${JSON.stringify(cluster)}, a piece of no source cluster`,
    ).toBe(true);
};

const KEYS = ["Backspace", "Delete"] as const;
const MODES = [
  "editing",
  "suggesting",
] as const satisfies readonly EditorMode[];

const checkParagraph = async (segments: readonly Segment[]) => {
  const source = documentFor(segments);
  const sourceViews = viewsOf(source);
  const rejectedSource = viewsOf(
    resolveCanonicalHarnessDocument(source, "reject").model,
  ).current;
  const alphabet = new Set(clustersOf(sourceViews.current));
  const probe = createCanonicalEditorHarness(source, "editing");
  const { visible, carets } = (() => {
    try {
      return caretsOf(probe.state.doc);
    } finally {
      probe.dispose();
    }
  })();
  expect(visible).toBe(sourceViews.current);
  let checked = 0;
  for (const mode of MODES)
    for (const key of KEYS)
      for (const { position, offset } of carets) {
        // At a visible paragraph edge the key joins paragraphs instead.
        if (key === "Backspace" ? offset === 0 : offset === visible.length)
          continue;
        const removed = clusterAt(visible, offset, key);
        const expected =
          visible.slice(0, removed.from) + visible.slice(removed.to);
        const label = `${mode} ${key} at ${String(position)} (visible offset ${String(offset)})`;
        const driver = createCanonicalEditorHarness(source, mode);
        const snapshot = (() => {
          try {
            driver.history.setSelection(position, position);
            expect(driver.state.selection.from, label).toBe(position);
            expect(driver.pressKey(key), label).toBe(true);
            expect(driver.refusals, label).toEqual([]);
            return driver.snapshot();
          } finally {
            driver.dispose();
          }
        })();
        const saved = viewsOf(
          (await saveCanonicalHarnessDocument(snapshot)).model,
        );
        for (const node of saved.nodes)
          expect(LONE_SURROGATE.test(node), `${label}: saved text node`).toBe(
            false,
          );
        expect(saved.current, `${label}: current text`).toBe(expected);
        expectWholeClusters(`${label}: current text`, saved.current, alphabet);
        const accepted = viewsOf(
          resolveCanonicalHarnessDocument(snapshot, "accept").model,
        ).current;
        expect(accepted, `${label}: accepted text`).toBe(expected);
        expectWholeClusters(`${label}: accepted text`, accepted, alphabet);
        if (mode === "editing")
          expect(saved.deleted, `${label}: struck text`).toBe(
            sourceViews.deleted,
          );
        else
          expect(
            viewsOf(resolveCanonicalHarnessDocument(snapshot, "reject").model)
              .current,
            `${label}: rejected text`,
          ).toBe(rejectedSource);
        checked += 1;
      }
  expect(checked).toBeGreaterThan(0);
};

const clusterArb = fc.constantFrom(...CLUSTERS);
const segmentArb = fc.record({
  revision: fc.constantFrom(...REVISIONS),
  text: fc
    .array(clusterArb, { minLength: 1, maxLength: 2 })
    .map((clusters) => clusters.join("")),
});
/** A multi-code-point cluster cut between two visible revisions. */
const splitArb = fc
  .record({
    cluster: fc.constantFrom(...MULTI_CODE_POINT_CLUSTERS),
    cut: fc.nat(),
    sides: fc.constantFrom(...SPLIT_SIDES),
  })
  .map(({ cluster, cut, sides: [left, right] }): Segment[] => {
    const codePoints = [...cluster];
    const at = 1 + (cut % (codePoints.length - 1));
    return [
      { revision: left, text: codePoints.slice(0, at).join("") },
      { revision: right, text: codePoints.slice(at).join("") },
    ];
  });
const paragraphArb = fc
  .tuple(fc.array(segmentArb, { maxLength: 3 }), splitArb, fc.nat())
  .map(([segments, split, at]) => {
    const paragraph = segments.slice();
    paragraph.splice(at % (segments.length + 1), 0, ...split);
    return paragraph;
  });

const PERSON_WITH_SKIN_TONE = MULTI_CODE_POINT_CLUSTERS[2];
const EXAMPLES: [Segment[]][] = [
  // A ZWJ person with a skin tone in plain text, with deletions on both sides.
  [
    [
      { revision: "plain", text: `a${PERSON_WITH_SKIN_TONE}` },
      { revision: "deleteOther", text: "Z" },
      { revision: "plain", text: PERSON_WITH_SKIN_TONE },
      { revision: "deleteOwn", text: "\u{1F1E8}\u{1F1FF}" },
    ],
  ],
  // A ZWJ sequence cut at its joiner between plain text and an own insertion.
  [
    [
      { revision: "plain", text: "a\u{1F468}\u{200D}" },
      {
        revision: "insertOwn",
        text: "\u{1F469}\u{200D}\u{1F467}\u{200D}\u{1F466}",
      },
      { revision: "deleteOther", text: "Z" },
    ],
  ],
  // A keycap cut before its enclosing mark, another author's insertion after.
  [
    [
      { revision: "insertOther", text: "1\u{FE0F}" },
      { revision: "plain", text: "\u{20E3}Z" },
    ],
  ],
  // A flag cut between its regional indicators across two authors.
  [
    [
      { revision: "deleteOther", text: "a" },
      { revision: "insertOther", text: "\u{1F1F8}" },
      { revision: "insertOwn", text: "\u{1F1F0}" },
      { revision: "deleteOwn", text: "e\u{301}" },
    ],
  ],
  // Combining marks after a deletion and on an own insertion.
  [
    [
      { revision: "plain", text: "Z" },
      { revision: "deleteOther", text: "\u{1F44D}\u{1F3FC}" },
      { revision: "plain", text: "a\u{308}" },
      { revision: "insertOwn", text: "\u{304} " },
    ],
  ],
];

test(
  "canonical Backspace and Delete remove exactly one whole grapheme cluster in editing and suggesting",
  async () => {
    await assertProperty(fc.asyncProperty(paragraphArb, checkParagraph), {
      numRuns: 12,
      examples: EXAMPLES,
      id: "canonical Backspace and Delete remove exactly one whole grapheme cluster in editing and suggesting",
    });
  },
  propertyTestTimeout(120_000),
);
