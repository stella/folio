import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";

describe("header and footer pointer ownership", () => {
  test("covers every descendant in inactive and editing modes", () => {
    const css = readFileSync(new URL("prosemirror-layer.css", import.meta.url), "utf-8");

    expect(css).toMatch(
      /\.layout-page-header \*,\n\.layout-page-footer \* \{\n  pointer-events: none;\n\}/u,
    );
    expect(css).toMatch(
      /\.paged-editor--editing-header \.layout-page-header \*,\n\.paged-editor--editing-footer \.layout-page-footer \* \{\n  pointer-events: auto;\n\}/u,
    );
  });
});
