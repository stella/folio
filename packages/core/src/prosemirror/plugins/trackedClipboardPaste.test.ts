import { afterAll, beforeAll, expect, test } from "bun:test";
import { DOMParser, type Mark, type Node as PmNode } from "prosemirror-model";
import {
  EditorState,
  TextSelection,
  type Transaction,
} from "prosemirror-state";

import { acquireHarnessDom } from "../../../../../test/harnessDom";
import { fromMarkdown } from "../../markdown/fromMarkdown";
import { fromProseDoc } from "../conversion/fromProseDoc";
import { toProseDoc } from "../conversion/toProseDoc";
import { transformPastedHTML } from "../extensions/features/PasteStyleInlinerExtension";
import { schema } from "../schema";
import { mintRevisionId } from "./revisionIds";
import {
  createSuggestionModePlugin,
  handleSuggestionPaste,
  suggestionModeKey,
} from "./suggestionMode";

let releaseDom: ReturnType<typeof acquireHarnessDom>;

beforeAll(() => {
  releaseDom = acquireHarnessDom();
});

afterAll(() => releaseDom());

const revisionMarkForText = (
  doc: PmNode | null,
  text: string,
  type: "insertion" | "deletion",
): Mark | undefined => {
  if (doc === null) return;
  let found: Mark | undefined;
  doc.descendants((node) => {
    if (!node.isText || node.text !== text) return;
    found = node.marks.find((mark) => mark.type.name === type);
  });
  return found;
};

const parseClipboardHtml = (html: string) => {
  const container = document.createElement("div");
  container.innerHTML = transformPastedHTML(html);
  return DOMParser.fromSchema(schema).parseSlice(container);
};

const copiedWordHtml =
  '<p>keep <ins cite="mailto:Ann%20Author" datetime="2026-10-22T14:18">added</ins> ' +
  '<del cite="mailto:Bob%20Reviewer" datetime="2026-10-22T14:19">removed</del></p>';

test("tracked HTML paste keeps source deletions while proposing the rich slice", () => {
  const slice = parseClipboardHtml(copiedWordHtml);
  const original = schema.node("doc", null, [
    schema.node("paragraph", null, [schema.text("replace")]),
  ]);
  let sourceDeletion = false;
  slice.content.descendants((node) => {
    if (
      node.isText &&
      node.text === "removed" &&
      node.marks.some(
        (mark) =>
          mark.type.name === "deletion" &&
          mark.attrs["author"] === "Bob Reviewer" &&
          mark.attrs["date"] === "2026-10-22T14:19",
      )
    ) {
      sourceDeletion = true;
    }
  });
  expect(sourceDeletion).toBe(true);
  let state = EditorState.create({
    doc: original,
    selection: TextSelection.create(original, 1, 8),
    plugins: [createSuggestionModePlugin(true, "Reviewer")],
  });
  const fittedPaste = state.tr.replaceRange(8, 8, slice).doc;
  let fittedDeletion = false;
  fittedPaste.descendants((node) => {
    if (
      node.isText &&
      node.text === "removed" &&
      node.marks.some(
        (mark) =>
          mark.type.name === "deletion" &&
          mark.attrs["author"] === "Bob Reviewer" &&
          mark.attrs["date"] === "2026-10-22T14:19",
      )
    ) {
      fittedDeletion = true;
    }
  });
  expect(fittedDeletion).toBe(true);
  const view = {
    get state() {
      return state;
    },
    dispatch(transaction: Transaction) {
      state = state.apply(transaction);
    },
  };
  const pluginState = suggestionModeKey.getState(state);
  if (!pluginState) throw new Error("suggestion-mode state is missing");

  expect(handleSuggestionPaste(view, slice, pluginState)).toBe(true);
  const revisions: {
    text: string;
    kind: string;
    author: string;
    date: string | null;
    revisionId: number;
    revisionAncestors: { type: string; author: string; date: string | null }[];
  }[] = [];
  state.doc.descendants((node) => {
    if (!node.isText) return;
    for (const mark of node.marks) {
      if (mark.type.name !== "insertion" && mark.type.name !== "deletion") {
        continue;
      }
      const author = mark.attrs["author"];
      const date = mark.attrs["date"];
      const revisionId = mark.attrs["revisionId"];
      if (
        typeof author !== "string" ||
        (typeof date !== "string" && date !== null) ||
        typeof revisionId !== "number"
      ) {
        continue;
      }
      revisions.push({
        text: node.text ?? "",
        kind: mark.type.name,
        author,
        date,
        revisionId,
        revisionAncestors: mark.attrs["_docxRevisionAncestors"] ?? [],
      });
    }
  });

  const removed = revisions.find(
    ({ text, kind }) => text === "removed" && kind === "deletion",
  );
  expect(removed).toMatchObject({
    author: "Bob Reviewer",
    date: "2026-10-22T14:19",
  });
  expect(removed?.revisionId).toBeGreaterThan(0);
  expect(
    removed?.revisionAncestors.some(
      ({ type, author }) => type === "insertion" && author === "Reviewer",
    ),
  ).toBe(true);
  expect(
    revisions.some(
      ({ text, kind }) => text === "added" && kind === "insertion",
    ),
  ).toBe(true);
});

test("an untracked HTML paste remains parseable as an ordinary rich slice", () => {
  const slice = parseClipboardHtml("<p>plain <strong>rich</strong> text</p>");
  expect(slice.content.textBetween(0, slice.content.size)).toBe(
    "plain rich text",
  );
});

test("native revision paste allocates a unique claimed id per element", () => {
  const slice = parseClipboardHtml(
    '<p><ins cite="mailto:Ann">first</ins><del cite="mailto:Bob">second</del>' +
      '<ins cite="mailto:Ann">third</ins></p>',
  );
  const ids: number[] = [];
  slice.content.descendants((node) => {
    if (!node.isText) return;
    for (const mark of node.marks) {
      if (mark.type.name === "insertion" || mark.type.name === "deletion") {
        ids.push(mark.attrs["revisionId"]);
      }
    }
  });

  expect(ids).toHaveLength(3);
  expect(new Set(ids).size).toBe(3);
  expect(ids).not.toContain(mintRevisionId());
});

test("repeated clipboard transforms mint fresh ids and reread authors", () => {
  const sameHtml = '<p><ins cite="mailto:Ann">same</ins></p>';
  const first = parseClipboardHtml(sameHtml);
  const repeated = parseClipboardHtml(sameHtml);
  const changed = parseClipboardHtml(
    '<p><ins cite="mailto:Bea">same</ins></p>',
  );
  const firstMark = revisionMarkForText(
    first.content.firstChild,
    "same",
    "insertion",
  );
  const repeatedMark = revisionMarkForText(
    repeated.content.firstChild,
    "same",
    "insertion",
  );
  const changedMark = revisionMarkForText(
    changed.content.firstChild,
    "same",
    "insertion",
  );

  expect(firstMark?.attrs["revisionId"]).not.toBe(
    repeatedMark?.attrs["revisionId"],
  );
  expect(firstMark?.attrs["revisionId"]).not.toBe(
    changedMark?.attrs["revisionId"],
  );
  expect(firstMark?.attrs["author"]).toBe("Ann");
  expect(repeatedMark?.attrs["author"]).toBe("Ann");
  expect(changedMark?.attrs["author"]).toBe("Bea");
});

test("native clipboard normalization ignores supplied revision identities", () => {
  const suppliedId = mintRevisionId();
  const slice = parseClipboardHtml(
    `<p><ins cite="mailto:Ann" data-revision-id="${suppliedId}" ` +
      'data-author="Forged" data-revision-ancestors="[]">text</ins></p>',
  );
  const mark = revisionMarkForText(
    slice.content.firstChild,
    "text",
    "insertion",
  );

  expect(mark?.attrs["revisionId"]).not.toBe(suppliedId);
  expect(mark?.attrs["author"]).toBe("Ann");
  expect(mark?.attrs["_docxRevisionAncestors"]).toBeNull();
});

test("schema-only native parsing remints reused DOM without inferring ancestry", () => {
  const container = document.createElement("div");
  container.innerHTML =
    '<p><ins cite="mailto:Outer"><ins cite="mailto:Inner" ' +
    'datetime="2026-10-20T10:00:00Z">nested</ins></ins></p>';
  const nestedElement = container.querySelector("ins ins");
  if (nestedElement === null) {
    throw new Error("The nested native revision element is missing");
  }
  const parseDirectly = () => {
    const slice = DOMParser.fromSchema(schema).parseSlice(container);
    return revisionMarkForText(
      slice.content.firstChild,
      "nested",
      "insertion",
    );
  };

  const first = parseDirectly();
  nestedElement.setAttribute("cite", "mailto:Changed");
  nestedElement.setAttribute("datetime", "2026-10-22T12:00:00Z");
  const second = parseDirectly();
  expect(first?.attrs["revisionId"]).not.toBe(second?.attrs["revisionId"]);
  expect(first?.attrs["author"]).toBe("Inner");
  expect(second?.attrs["author"]).toBe("Changed");
  expect(second?.attrs["date"]).toBe("2026-10-22T12:00:00Z");
  expect(first?.attrs["_docxRevisionAncestors"]).toBeNull();
});

test.each([
  {
    outerTag: "ins",
    innerTag: "ins",
    outerType: "insertion",
    innerType: "insertion",
    outerAuthor: "Outer insert",
    innerAuthor: "Inner insert",
  },
  {
    outerTag: "del",
    innerTag: "del",
    outerType: "deletion",
    innerType: "deletion",
    outerAuthor: "Outer delete",
    innerAuthor: "Inner delete",
  },
] as const)(
  "native nested <$outerTag> preserves same-type revision ancestry",
  ({ outerTag, innerTag, outerType, innerType, outerAuthor, innerAuthor }) => {
    const slice = parseClipboardHtml(
      `<p><${outerTag} cite="mailto:${encodeURIComponent(outerAuthor)}" ` +
        'datetime="2026-10-20T10:00:00Z">outer ' +
        `<${innerTag} cite="mailto:${encodeURIComponent(innerAuthor)}" ` +
        'datetime="2026-10-21T11:00:00Z">nested</' +
        `${innerTag}> after</${outerTag}></p>`,
    );

    const paragraph = slice.content.firstChild;
    if (paragraph === null) {
      throw new Error("The pasted paragraph is missing");
    }
    const outer = revisionMarkForText(paragraph, "outer ", outerType);
    const nested = revisionMarkForText(paragraph, "nested", innerType);
    if (outer === undefined || nested === undefined) {
      throw new Error("The pasted revision marks are missing");
    }

    expect(outer.attrs["revisionId"]).not.toBe(nested.attrs["revisionId"]);
    expect(nested.attrs["_docxRevisionAncestors"]).toEqual([
      {
        type: outerType,
        revisionId: outer.attrs["revisionId"],
        author: outerAuthor,
        date: "2026-10-20T10:00:00Z",
        outerWrapperCount: 0,
      },
    ]);

    const source = fromMarkdown("source");
    const saved = fromProseDoc(
      schema.node("doc", null, [paragraph]),
      source,
      { stylesheetSource: { type: "package" } },
    );
    const reopened = toProseDoc(saved);
    expect(
      revisionMarkForText(reopened, "outer ", outerType)?.attrs,
    ).toMatchObject({
      revisionId: outer.attrs["revisionId"],
      author: outerAuthor,
      date: "2026-10-20T10:00:00Z",
    });
    expect(
      revisionMarkForText(reopened, "nested", innerType)?.attrs,
    ).toMatchObject({
      revisionId: nested.attrs["revisionId"],
      author: innerAuthor,
      date: "2026-10-21T11:00:00Z",
      _docxRevisionAncestors: nested.attrs["_docxRevisionAncestors"],
    });
  },
);

test.each([
  {
    outerTag: "ins",
    innerTag: "del",
    outerType: "insertion",
    innerType: "deletion",
  },
  {
    outerTag: "del",
    innerTag: "ins",
    outerType: "deletion",
    innerType: "insertion",
  },
] as const)(
  "native mixed nested <$outerTag> <$innerTag> keeps ancestry order",
  ({ outerTag, innerTag, outerType, innerType }) => {
    const outerAuthor = "Outer reviewer";
    const innerAuthor = "Inner reviewer";
    const slice = parseClipboardHtml(
      `<p><${outerTag} cite="mailto:${outerAuthor}" ` +
        'datetime="2026-10-20T10:00:00Z">outer ' +
        `<${innerTag} cite="mailto:${innerAuthor}" ` +
        'datetime="2026-10-21T11:00:00Z">nested</' +
        `${innerTag}> after</${outerTag}></p>`,
    );

    const paragraph = slice.content.firstChild;
    if (paragraph === null) {
      throw new Error("The pasted paragraph is missing");
    }
    const outer = revisionMarkForText(paragraph, "outer ", outerType);
    const nested = revisionMarkForText(paragraph, "nested", innerType);
    if (outer === undefined || nested === undefined) {
      throw new Error("The pasted revision marks are missing");
    }

    expect(outer.attrs["revisionId"]).not.toBe(nested.attrs["revisionId"]);
    expect(nested.attrs["_docxRevisionAncestors"]).toEqual([
      {
        type: outerType,
        revisionId: outer.attrs["revisionId"],
        author: outerAuthor,
        date: "2026-10-20T10:00:00Z",
        outerWrapperCount: 0,
      },
    ]);

    const source = fromMarkdown("source");
    const saved = fromProseDoc(
      schema.node("doc", null, [paragraph]),
      source,
      { stylesheetSource: { type: "package" } },
    );
    const reopened = toProseDoc(saved);
    expect(
      revisionMarkForText(reopened, "outer ", outerType)?.attrs,
    ).toMatchObject({
      revisionId: outer.attrs["revisionId"],
      author: outerAuthor,
      date: "2026-10-20T10:00:00Z",
    });
    expect(
      revisionMarkForText(reopened, "nested", innerType)?.attrs,
    ).toMatchObject({
      revisionId: nested.attrs["revisionId"],
      author: innerAuthor,
      date: "2026-10-21T11:00:00Z",
      _docxRevisionAncestors: nested.attrs["_docxRevisionAncestors"],
    });
  },
);
