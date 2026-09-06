import { describe, expect, test } from "bun:test";

import { isKeydownInShortcutScope } from "./editorShortcuts";

/**
 * Bun's test runtime has no DOM, so roots and targets are the narrow surface
 * the predicate actually reads: a `contains` check against nodes identified by
 * `nodeType`. `contains` is `Node.contains` semantics — a node contains itself.
 */
type StubNode = { nodeType: number; contains(other: unknown): boolean };

const node = (...descendants: readonly unknown[]): StubNode => {
  const stub: StubNode = {
    nodeType: 1,
    contains: (other) => other === stub || descendants.includes(other),
  };
  return stub;
};

describe("isKeydownInShortcutScope", () => {
  test("'document' answers every press, wherever it lands", () => {
    const root = node();
    expect(isKeydownInShortcutScope({ target: node() }, { scope: "document", roots: [root] })).toBe(
      true,
    );
    expect(isKeydownInShortcutScope({ target: null }, { scope: "document", roots: [] })).toBe(true);
  });

  test("'none' answers nothing, even inside the editor", () => {
    const root = node();
    expect(isKeydownInShortcutScope({ target: root }, { scope: "none", roots: [root] })).toBe(
      false,
    );
  });

  test("'editor' answers a press inside a root", () => {
    const target = node();
    const root = node(target);
    expect(isKeydownInShortcutScope({ target }, { scope: "editor", roots: [root] })).toBe(true);
    expect(isKeydownInShortcutScope({ target: root }, { scope: "editor", roots: [root] })).toBe(
      true,
    );
  });

  test("'editor' ignores a press outside every root", () => {
    const root = node();
    expect(isKeydownInShortcutScope({ target: node() }, { scope: "editor", roots: [root] })).toBe(
      false,
    );
  });

  test("'editor' answers a press inside a second root, such as a portaled dialog", () => {
    const target = node();
    const dialog = node(target);
    expect(isKeydownInShortcutScope({ target }, { scope: "editor", roots: [node(), dialog] })).toBe(
      true,
    );
  });

  test("'editor' ignores an unmounted root", () => {
    expect(
      isKeydownInShortcutScope({ target: node() }, { scope: "editor", roots: [null, undefined] }),
    ).toBe(false);
  });

  test("'editor' ignores a target that is not a node", () => {
    const target = { addEventListener: () => undefined };
    const root = node(target);
    expect(isKeydownInShortcutScope({ target }, { scope: "editor", roots: [root] })).toBe(false);
    expect(isKeydownInShortcutScope({ target: null }, { scope: "editor", roots: [root] })).toBe(
      false,
    );
  });
});
