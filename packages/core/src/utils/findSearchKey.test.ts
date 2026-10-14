import { expect, test } from "bun:test";
import { FIND_SEARCH_KEY_OPTION_CHANGES } from "../../typecheck/findSearchKey.typecheck";
import { createEmptyDocument } from "./createDocument";
import { createDefaultFindOptions } from "./findReplace";
import { createFindQueryKey, createFindSearchKey, sameFindSearchKey } from "./findSearchKey";
import type { Document } from "../types/document";

const document = createEmptyDocument({ initialText: "hit" });
const base = { query: "hit", options: createDefaultFindOptions(), document };

test.each(Object.entries(FIND_SEARCH_KEY_OPTION_CHANGES))(
  "%s changes the immutable completed search key",
  (_kind, options) => {
    const mutable = createDefaultFindOptions();
    const previous = createFindSearchKey({ ...base, options: mutable });
    Object.assign(mutable, options);
    expect(sameFindSearchKey(previous, createFindSearchKey({ ...base, options: mutable }))).toBe(
      false,
    );
    expect(sameFindSearchKey(previous, createFindSearchKey(base))).toBe(true);
  },
);

test("equivalent options, fresh models and reordered scope have the same search key", () => {
  const bodyVersion = {};
  const previous = createFindSearchKey({
    ...base,
    bodyVersion,
    context: {
      renderedStories: [
        { kind: "header", rId: "rId1" },
        { kind: "footer", rId: "rId2" },
      ],
    },
  });
  const current = createFindSearchKey({
    ...base,
    document: structuredClone(document),
    bodyVersion,
    options: { ...createDefaultFindOptions(), displayMode: "all-markup" },
    context: {
      renderedStories: [
        { kind: "footer", rId: "rId2" },
        { kind: "header", rId: "rId1" },
      ],
    },
  });
  expect(sameFindSearchKey(previous, current)).toBe(true);
  expect(createFindQueryKey("hit", base.options)).toBe(previous.query);
  expect(sameFindSearchKey(null, current)).toBe(false);
});

test("query, body version and rendered story scope each invalidate the key", () => {
  const previous = createFindSearchKey({ ...base, bodyVersion: {} });
  expect(
    sameFindSearchKey(
      previous,
      createFindSearchKey({ ...base, query: "other", bodyVersion: previous.bodyVersion }),
    ),
  ).toBe(false);
  expect(sameFindSearchKey(previous, createFindSearchKey({ ...base, bodyVersion: {} }))).toBe(
    false,
  );
  expect(
    sameFindSearchKey(
      previous,
      createFindSearchKey({
        ...base,
        bodyVersion: previous.bodyVersion,
        context: { renderedStories: [] },
      }),
    ),
  ).toBe(false);
});

test("a model source version detects body edits without an editor version token", () => {
  const source = structuredClone(document);
  const previous = createFindSearchKey({ ...base, document: source });
  source.package.document.content = createEmptyDocument({
    initialText: "changed",
  }).package.document.content;
  expect(sameFindSearchKey(previous, createFindSearchKey({ ...base, document: source }))).toBe(
    false,
  );
});

test.each(["header", "footer", "footnote", "endnote", "comment"] as const)(
  "%s edits invalidate the key even with an unchanged editor body",
  (kind) => {
    const content = createEmptyDocument({ initialText: "side hit" }).package.document.content;
    const source: Document = {
      ...document,
      package: {
        ...document.package,
        headers: new Map([["rId1", { type: "header", content }]]),
        footers: new Map([["rId2", { type: "footer", content }]]),
        footnotes: [{ type: "footnote", id: 37, content }],
        endnotes: [{ type: "endnote", id: 41, content }],
        document: {
          ...document.package.document,
          comments: [{ id: 91, author: "Reviewer", content }],
        },
      },
    };
    const bodyVersion = {};
    const previous = createFindSearchKey({ ...base, document: source, bodyVersion });
    const changed = structuredClone(source);
    const replacement = createEmptyDocument({ initialText: "changed" }).package.document.content;
    switch (kind) {
      case "header":
        changed.package.headers?.set("rId1", { type: "header", content: replacement });
        break;
      case "footer":
        changed.package.footers?.set("rId2", { type: "footer", content: replacement });
        break;
      case "footnote":
        changed.package.footnotes = [{ type: "footnote", id: 37, content: replacement }];
        break;
      case "endnote":
        changed.package.endnotes = [{ type: "endnote", id: 41, content: replacement }];
        break;
      case "comment":
        changed.package.document.comments = [{ id: 91, author: "Reviewer", content: replacement }];
        break;
    }
    expect(
      sameFindSearchKey(previous, createFindSearchKey({ ...base, document: changed, bodyVersion })),
    ).toBe(false);
  },
);
