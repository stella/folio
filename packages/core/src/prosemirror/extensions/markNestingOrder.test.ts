/**
 * The editor nests a leaf's marks in the order the save leg nests its elements.
 *
 * Mark rank is registration order (`MARK_NESTING_ORDER`); element nesting is
 * decided in `extractParagraphContent`. Nothing but this held the two together,
 * and they drifted: the editor DOM read `<a><span class="docx-insertion">` while
 * the save wrote `w:ins > w:hyperlink > w:r`, so a rule written against one
 * nesting was written against a document the other leg does not produce.
 *
 * Only the marks the save leg gives an element of its own are ranked by it. A
 * formatting mark is a run property, one `w:rPr` inside the single `w:r`, so its
 * DOM rank is a presentation decision; the test states that by carrying bold
 * through and finding it on the run rather than in the chain.
 *
 * A comment is a range, not a tree: it is two markers the save leg writes in
 * document order, and a transparent wrapper it happens to open inside says the
 * same thing at either level (`fromProseDoc.ts:2388-2394`). So the comment's
 * place in the chain is read from whether its range opens before the run and
 * closes after it, which is what "encloses" means for a range.
 */

import { describe, expect, test } from "bun:test";
import { Window } from "happy-dom";
import { DOMSerializer, type Mark } from "prosemirror-model";

import type { ParagraphContent, Run } from "../../types/document";
import { proseDocToBlocks } from "../conversion/fromProseDoc";
import { inlineWrapperLayer } from "../inlineWrapperStack";
import { schema } from "../schema";
import { MARK_NESTING_ORDER, type SchemaMarkName } from "./markRegistry";

const paragraphWith = (marks: readonly Mark[]) =>
  schema.node("paragraph", undefined, [schema.text("x", [...marks])]);

// ---------------------------------------------------------------------------
// The DOM leg
// ---------------------------------------------------------------------------

const window = new Window();
const document = window.document as unknown as globalThis.Document;
const serializer = DOMSerializer.fromSchema(schema);

/** The elements a leaf's marks wrap it in, outermost first. */
const domElements = (marks: readonly Mark[]): Element[] => {
  const host = document.createElement("div");
  host.append(serializer.serializeFragment(paragraphWith(marks).content, { document }));
  const chain: Element[] = [];
  for (let element = host.firstElementChild; element; element = element.firstElementChild) {
    chain.push(element);
  }
  return chain;
};

/** What a mark's own element looks like, read off a leaf that carries only it. */
const elementSignature = (element: Element): string =>
  `${element.tagName.toLowerCase()}.${element.getAttribute("class") ?? ""}`;

const domChain = (marks: readonly Mark[]): string[] => {
  const bySignature = new Map(
    marks.map((mark) => {
      const [own] = domElements([mark]);
      expect(own).toBeDefined();
      // SAFETY: asserted defined above; every mark under test spells an element.
      return [elementSignature(own!), mark.type.name];
    }),
  );
  return domElements(marks).map((element) => {
    const name = bySignature.get(elementSignature(element));
    expect(name).toBeDefined();
    return name ?? elementSignature(element);
  });
};

// ---------------------------------------------------------------------------
// The save leg
// ---------------------------------------------------------------------------

/**
 * The mark that writes each element the save leg can emit, `null` where no
 * mark does.
 *
 * Total over `ParagraphContent` rather than a `switch` with a `default`: a
 * member added to the model, or an element that gains a mark of its own, is a
 * compile error here rather than an element silently left out of the chain
 * this test compares.
 */
const CONTAINER_MARK = {
  run: "runIdentity",
  hyperlink: "hyperlink",
  insertion: "insertion",
  deletion: "deletion",
  inlineWrapper: "inlineWrapper",
  // Ranges and boundaries, not containers: two markers in document order.
  commentRangeStart: null,
  commentRangeEnd: null,
  commentReference: null,
  moveFromRangeStart: null,
  moveFromRangeEnd: null,
  moveToRangeStart: null,
  moveToRangeEnd: null,
  bookmarkStart: null,
  bookmarkEnd: null,
  // Elements the save leg writes from a node rather than from a mark.
  moveFrom: null,
  moveTo: null,
  simpleField: null,
  complexField: null,
  inlineSdt: null,
  mathEquation: null,
  preservedInline: null,
} as const satisfies Record<ParagraphContent["type"], SchemaMarkName | null>;

/** The mark that writes this element, or `undefined` when no mark does. */
const containerMark = (item: ParagraphContent): SchemaMarkName | undefined =>
  CONTAINER_MARK[item.type] ?? undefined;

/** What this element holds, or `undefined` when it holds no inline list. */
const childrenOf = (item: ParagraphContent): readonly ParagraphContent[] | undefined => {
  switch (item.type) {
    case "hyperlink":
      return item.children;
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo":
    case "inlineSdt":
    case "inlineWrapper":
      return item.content;
    default:
      return undefined;
  }
};

type FoundRun = {
  /** The elements enclosing the run, outermost first. */
  containers: SchemaMarkName[];
  run: Run;
};

const findRun = (
  items: readonly ParagraphContent[],
  trail: SchemaMarkName[] = [],
): FoundRun | undefined => {
  for (const item of items) {
    if (item.type === "run") {
      return { containers: trail, run: item };
    }
    const children = childrenOf(item);
    if (!children) {
      continue;
    }
    const name = containerMark(item);
    const found = findRun(children, name ? [...trail, name] : trail);
    if (found) {
      return found;
    }
  }
  return undefined;
};

/** True when some comment range opens before the run and closes after it. */
const rangeEnclosesRun = (items: readonly ParagraphContent[]): boolean => {
  const openedBefore = new Set<number>();
  const closedAfter = new Set<number>();
  let seenRun = false;

  const visit = (list: readonly ParagraphContent[]): void => {
    for (const item of list) {
      if (item.type === "run") {
        seenRun = true;
        continue;
      }
      if (item.type === "commentRangeStart") {
        if (!seenRun) {
          openedBefore.add(item.id);
        }
        continue;
      }
      if (item.type === "commentRangeEnd") {
        if (seenRun) {
          closedAfter.add(item.id);
        }
        continue;
      }
      const children = childrenOf(item);
      if (children) {
        visit(children);
      }
    }
  };
  visit(items);

  return [...openedBefore].some((id) => closedAfter.has(id));
};

/** The mark elements the save leg puts around the leaf's run, outermost first. */
const saveChain = (marks: readonly Mark[]): FoundRun => {
  const [block] = proseDocToBlocks(schema.node("doc", undefined, [paragraphWith(marks)]));
  if (block?.type !== "paragraph") {
    throw new Error("the leaf did not save as one paragraph");
  }
  const found = findRun(block.content);
  if (!found) {
    throw new Error("the leaf did not save as a run");
  }
  return rangeEnclosesRun(block.content)
    ? { ...found, containers: ["comment", ...found.containers] }
    : found;
};

// ---------------------------------------------------------------------------

const COMMENT = schema.mark("comment", { commentId: 7 });
const WRAPPER = schema.mark("inlineWrapper", {
  stack: [inlineWrapperLayer({ kind: "bidi", control: "override", direction: "rtl" })],
});
const HYPERLINK = schema.mark("hyperlink", { href: "https://example.com/" });
const BOLD = schema.mark("bold");

const REVISIONS = [
  schema.mark("insertion", { revisionId: 3, author: "Reviewer" }),
  schema.mark("deletion", { revisionId: 4, author: "Reviewer" }),
] as const;

describe("a leaf under a comment, a revision, a wrapper, a link and bold", () => {
  for (const revision of REVISIONS) {
    const marks = [COMMENT, revision, WRAPPER, HYPERLINK, BOLD];

    test(`nests the same elements either leg saves it (${revision.type.name})`, () => {
      const { containers } = saveChain(marks);
      const ranked = new Set<string>(containers);

      expect(containers).toEqual([
        "comment",
        "inlineWrapper",
        revision.type.name,
        "hyperlink",
      ] as SchemaMarkName[]);
      expect(domChain(marks).filter((name) => ranked.has(name))).toEqual(containers);
    });

    test(`spells the change inside the wrapper's <bdo> (${revision.type.name})`, () => {
      const wrapper = domElements(marks).at(domChain(marks).indexOf("inlineWrapper"));

      expect(wrapper?.tagName.toLowerCase()).toBe("bdo");
      expect(wrapper?.firstElementChild?.getAttribute("class")).toBe(`docx-${revision.type.name}`);
      expect(wrapper?.firstElementChild?.firstElementChild?.tagName.toLowerCase()).toBe("a");
    });

    test(`ranks the run's own mark inside every container (${revision.type.name})`, () => {
      const { containers } = saveChain(marks);
      const rank = (name: SchemaMarkName): number => MARK_NESTING_ORDER.indexOf(name);

      for (const container of containers) {
        expect(rank(container)).toBeLessThan(rank(CONTAINER_MARK.run));
      }
    });

    test(`carries bold on the run, which the save leg does not rank (${revision.type.name})`, () => {
      const { run } = saveChain(marks);

      expect(run.formatting?.bold).toBe(true);
      expect(domChain(marks)).toContain("bold");
    });
  }
});
