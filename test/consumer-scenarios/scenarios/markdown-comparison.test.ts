import assert from "node:assert/strict";
import { test } from "node:test";

import { comparableMarkdown } from "../support/markdown-comparison.ts";

const fragments = [
  "Notices go to ",
  "[the notice address](https://example.test/notices)",
  " dated 1 May.",
];

test("revision span segmentation is a fixed point for every boundary and revision kind", () => {
  for (const kind of ["ins", "del"]) {
    for (const author of ["Second Reviewer", "Právník", "法務"]) {
      const wrap = (text: string, id: number) =>
        `<${kind} author="${author}" id="${String(id)}">${text}</${kind}>`;
      const expected = comparableMarkdown(wrap(fragments.join(""), 123));
      for (let mask = 0; mask < 1 << (fragments.length - 1); mask += 1) {
        const groups: string[] = [];
        let group = "";
        for (const [index, fragment] of fragments.entries()) {
          group += fragment;
          if (mask & (1 << index) || index === fragments.length - 1) {
            groups.push(wrap(group, index));
            group = "";
          }
        }
        const projected = comparableMarkdown(groups.join(""));
        assert.equal(projected, expected);
        assert.equal(comparableMarkdown(projected), projected);
      }
    }
  }
});

test("normalization keeps author, kind, text, hyperlink and untracked-gap changes detectable", () => {
  const original = '<del author="First" id="1">Before [link](https://example.test/a) after</del>';
  for (const mutation of [
    original.replace('author="First"', 'author="Second"'),
    original.replaceAll("del", "ins"),
    original.replace("Before", "Changed"),
    original.replace("example.test/a", "example.test/b"),
    '<del author="First" id="1">Before </del>untracked<del author="First" id="2">[link](https://example.test/a) after</del>',
    '<del author="First" id="1">Before </del><del author="Second" id="2">[link](https://example.test/a) after</del>',
  ]) {
    assert.notEqual(comparableMarkdown(mutation), comparableMarkdown(original));
  }
});
