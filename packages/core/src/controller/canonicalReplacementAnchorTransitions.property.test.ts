import { afterEach, expect, setSystemTime, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { getTrackedChangeStatsFromDoc } from "../ai-edits/read";
import {
  type CanonicalCommit,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import type {
  Comment,
  Document,
  Paragraph,
  ParagraphContent,
} from "../types/document";
import type { ReadonlyDocument } from "../types/documentModel";

afterEach(() => setSystemTime());

const TEXTS = ["abcdefghij", "klmnopqrst"] as const;
const BLOCK_IDS = ["10000001", "10000002"] as const;
const REPLACEMENT = "XY";

/** A gap between characters: the paragraph and the characters before it. */
type Gap = { paragraph: 0 | 1; offset: number };
type Anchor =
  { kind: "range"; start: Gap; end: Gap } | { kind: "point"; at: Gap };

const order = (left: Gap, right: Gap): number =>
  left.paragraph - right.paragraph || left.offset - right.offset;

/** What a paragraph says, with each comment marker as a token in place. */
const token = (node: { type: string; id?: unknown }): string | undefined => {
  switch (node.type) {
    case "commentRangeStart":
      return `[S${String(node.id)}]`;
    case "commentRangeEnd":
      return `[E${String(node.id)}]`;
    case "commentReference":
      return `[R${String(node.id)}]`;
    default:
      return undefined;
  }
};

const inlineView = (items: readonly unknown[]): string => {
  let out = "";
  for (const item of items) {
    if (typeof item !== "object" || item === null || !("type" in item))
      continue;
    const node = item as { type: string; id?: unknown };
    const marker = token(node);
    if (marker !== undefined) {
      out += marker;
      continue;
    }
    if (node.type === "text" && "text" in item && typeof item.text === "string")
      out += item.text;
    for (const key of ["content", "children"] as const)
      if (key in item) {
        const nested: unknown = (item as Record<string, unknown>)[key];
        if (Array.isArray(nested)) out += inlineView(nested);
      }
  }
  return out;
};

const view = (document: ReadonlyDocument) => ({
  paragraphs: document.package.document.content.flatMap((block) =>
    block.type === "paragraph" ? [inlineView(block.content)] : [],
  ),
  comments: (document.package.document.comments ?? [])
    .map(({ id }) => id)
    .sort((left, right) => left - right),
});

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

/** Markers each gap holds, in the order the anchors were drawn. */
const markersAt = (anchors: readonly Anchor[]) => {
  const at = new Map<string, ParagraphContent[]>();
  const put = (gap: Gap, node: ParagraphContent) => {
    const key = `${gap.paragraph}:${gap.offset}`;
    at.set(key, [...(at.get(key) ?? []), node]);
  };
  for (const [id, anchor] of anchors.entries()) {
    if (anchor.kind === "point") {
      put(anchor.at, { type: "commentReference", id });
      continue;
    }
    put(anchor.start, { type: "commentRangeStart", id });
    put(anchor.end, { type: "commentRangeEnd", id });
    put(anchor.end, { type: "commentReference", id });
  }
  return at;
};

const sourceDocument = (anchors: readonly Anchor[]): Document => {
  const at = markersAt(anchors);
  const paragraphs = TEXTS.map((text, paragraph): Paragraph => {
    const content: ParagraphContent[] = [];
    let pending = "";
    for (let offset = 0; offset <= text.length; offset++) {
      const markers = at.get(`${paragraph}:${offset}`) ?? [];
      if (markers.length > 0) {
        if (pending !== "") content.push(run(pending));
        pending = "";
        content.push(...markers);
      }
      pending += text.charAt(offset);
    }
    if (pending !== "") content.push(run(pending));
    const paraId = BLOCK_IDS[paragraph];
    if (paraId === undefined)
      throw new TypeError("Every fixture paragraph has a block id");
    return { type: "paragraph", paraId, content };
  });
  return {
    package: {
      document: {
        content: paragraphs,
        comments: anchors.map((anchor, id): Comment => ({
          id,
          anchorKind: anchor.kind,
          author: "Reviewer",
          content: [
            {
              type: "paragraph",
              paraId: (0x20000000 + id).toString(16),
              content: [run(`Comment ${id}`)],
            },
          ],
        })),
      },
    },
  };
};

/**
 * The policy, on the token view: the words in `[from, to)` go, every marker
 * they held moves to the start of the new words, in its order, and a ranged
 * comment whose every word went goes with them.
 */
const expectedView = (
  anchors: readonly Anchor[],
  from: Gap,
  to: Gap,
): ReturnType<typeof view> => {
  const at = markersAt(anchors);
  // The words and markers of `[start, end]`; no marker stands at an edge of
  // the replaced words, so the edges need no side.
  const stretch = (
    paragraph: 0 | 1,
    start: number,
    end: number,
    words: boolean,
  ) => {
    let out = "";
    for (let offset = start; offset <= end; offset++) {
      for (const node of at.get(`${paragraph}:${offset}`) ?? [])
        out += token(node) ?? "";
      if (words && offset < end) out += TEXTS[paragraph].charAt(offset);
    }
    return out;
  };
  const whole = (paragraph: 0 | 1) =>
    stretch(paragraph, 0, TEXTS[paragraph].length, true);
  const moved =
    from.paragraph === to.paragraph
      ? stretch(from.paragraph, from.offset, to.offset, false)
      : stretch(from.paragraph, from.offset, TEXTS[0].length, false) +
        stretch(to.paragraph, 0, to.offset, false);
  const joined =
    stretch(from.paragraph, 0, from.offset, true) +
    moved +
    REPLACEMENT +
    stretch(to.paragraph, to.offset, TEXTS[to.paragraph].length, true);
  let paragraphs =
    from.paragraph === to.paragraph
      ? ([0, 1] as const).map((paragraph) =>
          paragraph === from.paragraph ? joined : whole(paragraph),
        )
      : [joined];
  // A ranged comment that covered words and covers none now goes with them.
  const covers = (texts: readonly string[], id: number) => {
    const all = texts.join("\n");
    return (
      all
        .slice(all.indexOf(`[S${id}]`), all.indexOf(`[E${id}]`))
        .replaceAll(/\[[SER]\d+\]/gu, "")
        .replaceAll("\n", "") !== ""
    );
  };
  const source = [whole(0), whole(1)];
  const retired = anchors.flatMap((anchor, id) =>
    anchor.kind === "range" && covers(source, id) && !covers(paragraphs, id)
      ? [id]
      : [],
  );
  for (const id of retired)
    paragraphs = paragraphs.map((paragraph) =>
      paragraph
        .replace(`[S${id}]`, "")
        .replace(`[E${id}]`, "")
        .replace(`[R${id}]`, ""),
    );
  return {
    paragraphs,
    comments: anchors.map((_, id) => id).filter((id) => !retired.includes(id)),
  };
};

const gapArbitrary = fc.record({
  paragraph: fc.constantFrom(0 as const, 1 as const),
  offset: fc.integer({ min: 0, max: TEXTS[0].length }),
});

const caseArbitrary = fc
  .record({
    from: gapArbitrary,
    to: gapArbitrary,
    anchors: fc.array(
      fc.oneof(
        fc.tuple(gapArbitrary, gapArbitrary).map(([left, right]): Anchor => {
          const [start, end] =
            order(left, right) <= 0 ? [left, right] : [right, left];
          return { kind: "range", start, end };
        }),
        gapArbitrary.map((at): Anchor => ({ kind: "point", at })),
      ),
      { minLength: 1, maxLength: 3 },
    ),
    mode: fc.constantFrom("editing" as const, "suggesting" as const),
  })
  .map(({ from, to, anchors, mode }) =>
    order(from, to) <= 0
      ? { from, to, anchors, mode }
      : { from: to, to: from, anchors, mode },
  )
  .filter(({ from, to, anchors }) => {
    if (order(from, to) === 0) return false;
    // A marker at an edge of the replaced words sits on either side of
    // them at the editor's choice; the property draws markers off the edges.
    const offEdges = (gap: Gap) =>
      order(gap, from) !== 0 && order(gap, to) !== 0;
    return anchors.every((anchor) =>
      anchor.kind === "point"
        ? offEdges(anchor.at)
        : offEdges(anchor.start) &&
          offEdges(anchor.end) &&
          // An empty range is a point the source already collapsed.
          order(anchor.start, anchor.end) < 0,
    );
  });

test(
  "replacing words over comment markers moves them to the start of the new words, directly, accepted, rejected, undone and saved",
  async () => {
    setSystemTime(new Date("2026-01-02T03:04:05Z"));
    await assertProperty(
      fc.asyncProperty(caseArbitrary, async ({ from, to, anchors, mode }) => {
        const source = sourceDocument(anchors);
        const session = createCanonicalSession(source).unwrap();
        session.setMode(
          mode === "editing"
            ? { type: mode }
            : { type: mode, author: "Reviewer" },
        );
        const original = session.document;
        let state = EditorState.create({ doc: session.projection.doc });
        const publish = (commit: CanonicalCommit) => {
          state = publishCanonicalProjection({
            session,
            state,
            commit,
          }).unwrap().state;
        };
        // The editor position right after a gap's characters, counting
        // words only; no marker stands at a drawn edge, so its side needs
        // no choice.
        const position = (gap: Gap) => {
          const paragraph = session.projection.paragraph(
            BLOCK_IDS[gap.paragraph],
          );
          if (paragraph === undefined) throw new Error("Missing paragraph");
          if (gap.offset === 0) return paragraph.start;
          let seen = 0;
          let at: number | undefined;
          state.doc.nodesBetween(
            paragraph.start,
            paragraph.start + paragraph.node.content.size,
            (node, offset) => {
              if (at !== undefined || !node.isText) return at === undefined;
              const length = node.text?.length ?? 0;
              if (seen + length >= gap.offset) at = offset + gap.offset - seen;
              seen += length;
              return false;
            },
          );
          if (at === undefined) throw new Error("Missing gap");
          return at;
        };
        state = state.apply(
          state.tr.setSelection(
            TextSelection.create(state.doc, position(from), position(to)),
          ),
        );
        const expected = expectedView(anchors, from, to);
        publish(
          session
            .prepareReplace(state, {
              from: state.selection.from,
              to: state.selection.to,
              text: REPLACEMENT,
              semantic: "typing",
            })
            .unwrap(),
        );
        if (mode === "suggesting") {
          const revisionIds = [...getTrackedChangeStatsFromDoc(state.doc).ids];
          // Pending, every marker stays among the deleted words.
          expect(view(session.document).comments).toEqual(
            view(original).comments,
          );
          publish(
            session
              .prepareResolve(state, { revisionIds, resolution: "reject" })
              .unwrap(),
          );
          expect(view(session.document)).toEqual(view(original));
          expect(session.document).toEqual(original);
          publish(session.prepareUndo(state).unwrap());
          publish(
            session
              .prepareResolve(state, { revisionIds, resolution: "accept" })
              .unwrap(),
          );
        }
        expect(view(session.document)).toEqual(expected);
        const saved = await serializeCanonicalSave({
          snapshot: session.captureSaveSnapshot(),
        });
        const reopened = await parseDocx(saved.buffer, {
          preloadFonts: false,
          detectVariables: false,
        });
        expect(view(reopened)).toEqual(expected);
        while (session.canUndo) publish(session.prepareUndo(state).unwrap());
        expect(session.document).toEqual(original);
      }),
      {
        id: "replacement anchor transitions",
        seed: 3_640_001,
        numRuns: 120,
      },
    );
  },
  propertyTestTimeout(60_000),
);
