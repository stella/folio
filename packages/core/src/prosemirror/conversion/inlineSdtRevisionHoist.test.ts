/**
 * A revision around a whole inline content control survives the editor.
 *
 * `w:ins > w:sdt` used to reach the editor as a control whose leaves carried
 * nothing: the control is an `inline*` node rather than an atom, so it is not
 * a run carrier and the revision had nowhere to go. The save leg then wrote
 * the control back beside the revision that had held it, and the text the
 * reviewer inserted was no longer inserted.
 *
 * The revision now rides the leaves the control holds, and a revision that
 * covers all of them is written back around the control. What that costs is
 * the same thing the transparent wrapper costs: `w:ins > w:sdt` and
 * `w:sdt > w:ins` are the same marks on the same leaves, so both come back
 * revision-outermost.
 */

import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { pluginsForHeadlessRevisionResolution } from "../../internal/headlessRevisionResolutionGuard";
import type { Document, InlineSdt, Paragraph, ParagraphContent, Run } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import {
  acceptAllChanges,
  rejectAllChanges,
  resolveAllChangesInHeadlessState,
} from "../commands/comments";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const run = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });

const REVISION_INFO = { id: 4, author: "Reviewer", date: "2026-01-01T00:00:00Z" };

const control = (content: InlineSdt["content"]): InlineSdt => ({
  type: "inlineSdt",
  properties: { sdtType: "richText", tag: "bound" },
  content,
});

/** `w:ins > w:sdt > w:r`: the reviewer inserted the control. */
const REVISION_OUTSIDE_CONTROL: ParagraphContent = {
  type: "insertion",
  info: REVISION_INFO,
  content: [control([run("inside")])],
};

/** `w:sdt > w:ins > w:r`: the same span with the revision under the control. */
const REVISION_INSIDE_CONTROL: ParagraphContent = control([
  { type: "insertion", info: REVISION_INFO, content: [run("inside")] },
]);

/** A revision over part of the content, which has no outermost form. */
const REVISION_OVER_PART: ParagraphContent = control([
  run("kept"),
  { type: "insertion", info: REVISION_INFO, content: [run("added")] },
]);

const documentWith = (content: Paragraph["content"]): Document => {
  const template = createEmptyDocument();
  return {
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [{ type: "paragraph", paraId: "C0000001", content }],
      },
    },
  };
};

const paragraphContentOf = (document: Document): ParagraphContent[] => {
  const block = document.package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("The rebuilt document lost its paragraph");
  }
  return block.content;
};

/** The paragraph as a save rebuilds it from the editor. */
const saved = (content: Paragraph["content"]): ParagraphContent[] => {
  const source = documentWith(content);
  return paragraphContentOf(fromProseDoc(toProseDoc(source), source));
};

const insertedControl = (item: ParagraphContent | undefined): InlineSdt => {
  if (item?.type !== "insertion") {
    throw new Error(`Expected an insertion, got ${item?.type ?? "nothing"}`);
  }
  const inner = item.content.at(0);
  if (inner?.type !== "inlineSdt") {
    throw new Error(`Expected the insertion to hold a control, got ${inner?.type ?? "nothing"}`);
  }
  return inner;
};

const textIn = (items: readonly ParagraphContent[]): string => {
  let text = "";
  const visit = (content: readonly ParagraphContent[]): void => {
    for (const item of content) {
      switch (item.type) {
        case "run":
          for (const part of item.content) {
            if (part.type === "text") {
              text += part.text;
            }
          }
          break;
        case "inlineSdt":
        case "inlineWrapper":
        case "insertion":
        case "deletion":
        case "moveFrom":
        case "moveTo":
          visit(item.content);
          break;
        default:
          break;
      }
    }
  };
  visit(items);
  return text;
};

describe("a revision that covers a whole inline content control", () => {
  for (const authored of [REVISION_OUTSIDE_CONTROL, REVISION_INSIDE_CONTROL]) {
    const order = authored.type === "insertion" ? "revision outside" : "revision inside";

    test(`comes back around the control (${order})`, () => {
      const rebuilt = saved([authored]);
      expect(rebuilt).toHaveLength(1);
      expect(insertedControl(rebuilt.at(0)).properties.tag).toBe("bound");
      expect(textIn(rebuilt)).toBe("inside");
    });

    test(`saving the result again does not change it (${order})`, () => {
      const once = saved([authored]);
      expect(saved(once)).toEqual(once);
    });
  }

  test("a revision over part of the content stays inside the control", () => {
    const rebuilt = saved([REVISION_OVER_PART]);
    const only = rebuilt.at(0);
    if (only?.type !== "inlineSdt") {
      throw new Error(`Expected a control, got ${only?.type ?? "nothing"}`);
    }
    expect(only.content.map((item) => item.type)).toEqual(["run", "insertion"]);
    expect(textIn(rebuilt)).toBe("keptadded");
    expect(saved(rebuilt)).toEqual(rebuilt);
  });
});

describe("resolving a revision that covers a whole control", () => {
  const resolved = (
    content: Paragraph["content"],
    command: typeof acceptAllChanges,
  ): ParagraphContent[] => {
    const source = documentWith(content);
    const state = EditorState.create({ doc: toProseDoc(source) });
    let next = state;
    command()(state, (transaction) => {
      next = state.apply(transaction);
    });
    return paragraphContentOf(fromProseDoc(next.doc, source));
  };

  const INSERTED = [REVISION_OUTSIDE_CONTROL, run(" after")];
  const DELETED = [
    { type: "deletion", info: REVISION_INFO, content: [control([run("inside")])] } as const,
    run(" after"),
  ];

  test("accepting an inserted control keeps the control", () => {
    const rebuilt = resolved(INSERTED, acceptAllChanges);
    expect(rebuilt.at(0)?.type).toBe("inlineSdt");
    expect(textIn(rebuilt)).toBe("inside after");
  });

  test("rejecting an inserted control removes the control", () => {
    const rebuilt = resolved(INSERTED, rejectAllChanges);
    expect(rebuilt.every((item) => item.type !== "inlineSdt")).toBe(true);
    expect(textIn(rebuilt)).toBe(" after");
  });

  test("accepting a deleted control removes the control", () => {
    const rebuilt = resolved(DELETED, acceptAllChanges);
    expect(rebuilt.every((item) => item.type !== "inlineSdt")).toBe(true);
    expect(textIn(rebuilt)).toBe(" after");
  });

  test("rejecting a deleted control keeps the control", () => {
    const rebuilt = resolved(DELETED, rejectAllChanges);
    expect(rebuilt.at(0)?.type).toBe("inlineSdt");
    expect(textIn(rebuilt)).toBe("inside after");
  });

  // The headless resolver rewrites the inline content in one pass instead of
  // deleting ranges, so it decides the control's fate in its own code. The two
  // have to agree, or which one ran would change the document.
  test("the headless resolver removes it too", () => {
    const source = documentWith(INSERTED);
    const state = EditorState.create({
      doc: toProseDoc(source),
      plugins: [...pluginsForHeadlessRevisionResolution([])],
    });
    const rebuilt = paragraphContentOf(
      fromProseDoc(resolveAllChangesInHeadlessState(state, "reject").doc, source),
    );
    expect(rebuilt.every((item) => item.type !== "inlineSdt")).toBe(true);
    expect(textIn(rebuilt)).toBe(" after");
  });
});
