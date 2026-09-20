/**
 * A comment or tracked move anchored on a point survives the editor.
 *
 * Word writes a comment left at an insertion point as `w:commentRangeStart`
 * immediately followed by `w:commentRangeEnd`; a tracked move whose range
 * covers nothing is spelled the same way. The editor projects a comment as a
 * mark over the content its range covers and a move range as a marker placed
 * around the move's wrappers, so a range over nothing had no carrier and both
 * markers were dropped: the comment kept listing, because its reference is an
 * atom of its own, but the anchor position was gone and the save wrote a
 * reference with no range.
 *
 * The properties below are the whole contract, and each failed before the
 * `rangeAnchor` node existed:
 *
 * - the pair comes back at the position it was authored at, inside the wrapper
 *   it was authored inside;
 * - a second round trip changes nothing;
 * - an edit elsewhere in the paragraph leaves the anchor where it was.
 *
 * The containers are the inline ones the container census measures. The
 * revision wrappers and the inline content control are not among them by
 * construction rather than by omission: `TRACKED_CHANGE_WRAPPER_CONTENT` and
 * `INLINE_SDT_CONTENT` admit neither half of a range, so the parser lifts a
 * marker out to the paragraph and the model has no state where one sits
 * inside. The census, which builds those packages for real, is what checks
 * that lifting.
 */

import { describe, expect, test } from "bun:test";

import type {
  Document,
  InlineWrapper,
  Paragraph,
  ParagraphContent,
  Run,
} from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const run = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });

const MOVE_MARKER = { name: "mv", author: "Reviewer", date: "2026-01-01T00:00:00Z" } as const;

/** The three kinds of range a point can carry, as the pair of markers. */
const EMPTY_RANGES = {
  comment: [
    { type: "commentRangeStart", id: 1 },
    { type: "commentRangeEnd", id: 1 },
  ],
  moveFrom: [
    { type: "moveFromRangeStart", id: 2, ...MOVE_MARKER },
    { type: "moveFromRangeEnd", id: 2 },
  ],
  moveTo: [
    { type: "moveToRangeStart", id: 3, ...MOVE_MARKER },
    { type: "moveToRangeEnd", id: 3 },
  ],
} as const satisfies Record<string, readonly [ParagraphContent, ParagraphContent]>;

type RangeKind = keyof typeof EMPTY_RANGES;

/**
 * Where the pair sits: directly in the paragraph, or inside one of the four
 * transparent inline wrappers the projection lifts.
 */
const CONTAINERS = {
  "w:p": null,
  "w:bdo": { type: "inlineWrapper", kind: "bidi", control: "override", direction: "rtl" },
  "w:dir": { type: "inlineWrapper", kind: "bidi", control: "embedding" },
  "w:smartTag": { type: "inlineWrapper", kind: "smartTag", element: "place" },
  "w:customXml": { type: "inlineWrapper", kind: "customXml", element: "tagged" },
} as const satisfies Record<string, Omit<InlineWrapper, "content"> | null>;

type ContainerName = keyof typeof CONTAINERS;

/**
 * `before` text, the pair in its container, `after` text.
 *
 * A comment carries its `w:commentReference` right after the range, which is
 * how every point comment in the public corpus is spelled: the only thing
 * between the end marker and the reference is the run that holds it.
 */
const paragraphContentFor = (container: ContainerName, kind: RangeKind): ParagraphContent[] => {
  const [start, end] = EMPTY_RANGES[kind];
  const wrapper = CONTAINERS[container];
  const middle: ParagraphContent[] =
    wrapper === null ? [start, end] : [{ ...wrapper, content: [start, end] }];
  const reference: ParagraphContent[] =
    kind === "comment" ? [{ type: "commentReference", id: start.id }] : [];
  return [run("before"), ...middle, ...reference, run("after")];
};

const documentWith = (content: Paragraph["content"]): Document => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [{ type: "paragraph", paraId: "C0000001", content }],
        comments: [{ id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z", content: [] }],
      },
    },
  };
};

const paragraphContentOf = (document: Document): ParagraphContent[] => {
  const [block] = document.package.document.content;
  if (block?.type !== "paragraph") {
    throw new Error("The story lost its paragraph");
  }
  return block.content;
};

const roundTrip = (source: Document): Document => fromProseDoc(toProseDoc(source), source);

const containerNames = Object.keys(CONTAINERS) as ContainerName[];
const rangeKinds = Object.keys(EMPTY_RANGES) as RangeKind[];

describe("a range that spans no content", () => {
  for (const container of containerNames) {
    for (const kind of rangeKinds) {
      test(`${kind} in ${container}: comes back where it was authored`, () => {
        const source = documentWith(paragraphContentFor(container, kind));
        expect(paragraphContentOf(roundTrip(source))).toEqual(paragraphContentFor(container, kind));
      });

      test(`${kind} in ${container}: a second round trip is a fixed point`, () => {
        const source = documentWith(paragraphContentFor(container, kind));
        const once = roundTrip(source);
        expect(paragraphContentOf(roundTrip(once))).toEqual(paragraphContentOf(once));
      });

      test(`${kind} in ${container}: an edit elsewhere leaves the anchor alone`, () => {
        const edited = paragraphContentFor(container, kind).map((item) =>
          item.type === "run" &&
          item.content[0]?.type === "text" &&
          item.content[0].text === "after"
            ? run("after, edited")
            : item,
        );
        expect(paragraphContentOf(roundTrip(documentWith(edited)))).toEqual(edited);
      });
    }
  }

  test("a range over content still travels as a mark, not an anchor", () => {
    const [start, end] = EMPTY_RANGES.comment;
    const content: ParagraphContent[] = [
      start,
      run("covered"),
      end,
      { type: "commentReference", id: start.id },
    ];
    expect(paragraphContentOf(roundTrip(documentWith(content)))).toEqual(content);
    let anchors = 0;
    toProseDoc(documentWith(content)).descendants((node) => {
      if (node.type.name === "rangeAnchor") {
        anchors += 1;
      }
      return true;
    });
    expect(anchors).toBe(0);
  });

  test("an unpaired marker keeps the path it had", () => {
    const [start] = EMPTY_RANGES.comment;
    const content: ParagraphContent[] = [start, run("covered")];
    let anchors = 0;
    toProseDoc(documentWith(content)).descendants((node) => {
      if (node.type.name === "rangeAnchor") {
        anchors += 1;
      }
      return true;
    });
    expect(anchors).toBe(0);
  });
});
