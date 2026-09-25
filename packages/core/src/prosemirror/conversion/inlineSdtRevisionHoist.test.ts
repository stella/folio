/**
 * A revision around a whole inline content control survives the editor, and
 * so does one inside it.
 *
 * `w:ins > w:sdt` used to reach the editor as a control whose leaves carried
 * nothing: the control is an `inline*` node rather than an atom, so it is not
 * a run carrier and the revision had nowhere to go. The save leg then wrote
 * the control back beside the revision that had held it, and the text the
 * reviewer inserted was no longer inserted.
 *
 * The revision now rides the leaves the control holds, and the control records
 * that the revision encloses it. `w:sdt > w:ins` puts the same marks on the
 * same leaves without the record, so each comes back in the order it was
 * authored, and resolving the first removes the control while resolving the
 * second only empties it.
 */

import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { pluginsForHeadlessRevisionResolution } from "../../internal/headlessRevisionResolutionGuard";
import type { Document, InlineSdt, Paragraph, ParagraphContent, Run } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import {
  acceptAIEditRevision,
  acceptChange,
  acceptAllChanges,
  findNextChange,
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
  test("comes back around the control when it was authored around it", () => {
    const rebuilt = saved([REVISION_OUTSIDE_CONTROL]);
    expect(rebuilt).toHaveLength(1);
    expect(insertedControl(rebuilt.at(0)).properties.tag).toBe("bound");
    expect(textIn(rebuilt)).toBe("inside");
  });

  test("stays inside the control when it was authored inside it", () => {
    expect(saved([REVISION_INSIDE_CONTROL])).toEqual([REVISION_INSIDE_CONTROL]);
  });

  for (const authored of [REVISION_OUTSIDE_CONTROL, REVISION_INSIDE_CONTROL]) {
    const order = authored.type === "insertion" ? "revision outside" : "revision inside";
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
  const resolvedDoc = (content: Paragraph["content"], command: typeof acceptAllChanges): PMNode => {
    const state = EditorState.create({ doc: toProseDoc(documentWith(content)) });
    let next = state;
    command()(state, (transaction) => {
      next = state.apply(transaction);
    });
    return next.doc;
  };
  const resolved = (
    content: Paragraph["content"],
    command: typeof acceptAllChanges,
  ): ParagraphContent[] =>
    paragraphContentOf(fromProseDoc(resolvedDoc(content, command), documentWith(content)));

  const INSERTED = [REVISION_OUTSIDE_CONTROL, run(" after")];
  const DELETED = [
    { type: "deletion", info: REVISION_INFO, content: [control([run("inside")])] } as const,
    run(" after"),
  ];
  const CONTENT_INSERTED = [REVISION_INSIDE_CONTROL, run(" after")];
  const CONTENT_DELETED = [
    control([{ type: "deletion", info: REVISION_INFO, content: [run("inside")] }]),
    run(" after"),
  ];
  const emptiedControl = (rebuilt: readonly ParagraphContent[]): InlineSdt => {
    const first = rebuilt.at(0);
    if (first?.type !== "inlineSdt") {
      throw new Error(`Expected the control to stay, got ${first?.type ?? "nothing"}`);
    }
    expect(first.content).toEqual([]);
    return first;
  };
  const enclosingRevisionsIn = (doc: PMNode): unknown[] => {
    const found: unknown[] = [];
    doc.descendants((node) => {
      if (node.type.name === "sdt") {
        found.push(node.attrs["_docxEnclosingRevisionIds"]);
      }
      return true;
    });
    return found;
  };

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

  test("accepting the deletion of a control's content leaves the control, emptied", () => {
    const rebuilt = resolved(CONTENT_DELETED, acceptAllChanges);
    expect(emptiedControl(rebuilt).properties.tag).toBe("bound");
    expect(textIn(rebuilt)).toBe(" after");
  });

  test("rejecting the insertion of a control's content leaves the control, emptied", () => {
    const rebuilt = resolved(CONTENT_INSERTED, rejectAllChanges);
    expect(emptiedControl(rebuilt).properties.tag).toBe("bound");
    expect(textIn(rebuilt)).toBe(" after");
  });

  // The leaves carry the nested insertion's id, with the deletion around the
  // control only among their ancestors.
  test("accepting an enclosing deletion by its id removes the control as accept-all does", () => {
    const nested = [
      {
        type: "deletion",
        info: { ...REVISION_INFO, id: 5 },
        content: [
          control([{ type: "insertion", info: { ...REVISION_INFO, id: 6 }, content: [run("in")] }]),
        ],
      } as const,
      run(" after"),
    ];
    const byId = resolved(nested, () => acceptAIEditRevision(5));
    expect(byId.every((item) => item.type !== "inlineSdt")).toBe(true);
    expect(textIn(byId)).toBe(" after");
    expect(resolved(nested, acceptAllChanges)).toEqual(byId);
  });

  test("accepting the next change resolves the control a revision encloses", () => {
    for (const [content, controls] of [
      [DELETED, 0],
      [CONTENT_DELETED, 1],
    ] as const) {
      const rebuilt = resolved(content, () => (state, dispatch) => {
        const range = findNextChange(state, 0);
        return range !== null && acceptChange(range.from, range.to)(state, dispatch);
      });
      expect(rebuilt.filter((item) => item.type === "inlineSdt")).toHaveLength(controls);
      expect(textIn(rebuilt)).toBe(" after");
    }
  });

  test("rejecting the deletion of a control's content restores it", () => {
    expect(resolved(CONTENT_DELETED, rejectAllChanges)).toEqual([
      control([run("inside")]),
      run(" after"),
    ]);
  });

  // The headless resolver rewrites the inline content in one pass instead of
  // deleting ranges, so it decides the control's fate in its own code. The two
  // have to agree, or which one ran would change the document.
  const headlessResolvedDoc = (content: Paragraph["content"], mode: "accept" | "reject") => {
    const state = EditorState.create({
      doc: toProseDoc(documentWith(content)),
      plugins: [...pluginsForHeadlessRevisionResolution([])],
    });
    return resolveAllChangesInHeadlessState(state, mode).doc;
  };
  const headlessResolved = (
    content: Paragraph["content"],
    mode: "accept" | "reject",
  ): ParagraphContent[] =>
    paragraphContentOf(fromProseDoc(headlessResolvedDoc(content, mode), documentWith(content)));

  test("the headless resolver removes an enclosed control too", () => {
    const rebuilt = headlessResolved(INSERTED, "reject");
    expect(rebuilt.every((item) => item.type !== "inlineSdt")).toBe(true);
    expect(textIn(rebuilt)).toBe(" after");
  });

  test("the headless resolver empties a control whose content was revised", () => {
    for (const [content, mode] of [
      [CONTENT_DELETED, "accept"],
      [CONTENT_INSERTED, "reject"],
    ] as const) {
      const rebuilt = headlessResolved(content, mode);
      expect(emptiedControl(rebuilt).properties.tag).toBe("bound");
      expect(textIn(rebuilt)).toBe(" after");
    }
  });

  // A resolved revision no longer exists; a control that kept naming it would
  // claim whichever later revision is given the same id.
  test("a control forgets the revision that enclosed it once it is resolved", () => {
    expect(enclosingRevisionsIn(toProseDoc(documentWith(INSERTED)))).toEqual([[REVISION_INFO.id]]);
    expect(enclosingRevisionsIn(resolvedDoc(INSERTED, acceptAllChanges))).toEqual([null]);
    expect(enclosingRevisionsIn(headlessResolvedDoc(INSERTED, "accept"))).toEqual([null]);
  });
});
