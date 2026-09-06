import { describe, expect, test } from "bun:test";

import { htmlCommentEnd } from "./htmlComments";

describe("htmlCommentEnd", () => {
  test("closes an abrupt comment at its own '>'", () => {
    expect(htmlCommentEnd("<!--><p>keep me</p>", 0)).toBe(5);
    expect(htmlCommentEnd("<!---><p>keep me</p>", 0)).toBe(6);
  });

  test("closes an ordinary comment past its terminator", () => {
    expect(htmlCommentEnd("<!-- note --><p>keep me</p>", 0)).toBe(13);
  });

  test("reports an unterminated comment", () => {
    expect(htmlCommentEnd("before<!--dangling", 6)).toBeNull();
  });

  test("a comment whose body starts with a dash is not abrupt", () => {
    // `<!---x-->`: the third dash opens the body, it does not close the comment.
    expect(htmlCommentEnd("<!---x-->", 0)).toBe(9);
  });
});
