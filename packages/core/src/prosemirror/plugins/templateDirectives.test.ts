import { describe, expect, test } from "bun:test";

import { schema } from "../schema";
import type { DirectiveRange } from "./templateDirectives";
import { computeBlockDepths, scanDirectives } from "./templateDirectives";

const docOf = (...paragraphs: string[]) =>
  schema.node(
    "doc",
    null,
    paragraphs.map((text) => schema.node("paragraph", null, text ? [schema.text(text)] : null)),
  );

describe("scanDirectives", () => {
  test("recognizes @num and @ref numbering markers as their own kinds", () => {
    const doc = docOf(
      "Clause {{@num:scope}}. Scope of authority.",
      "As set out in Clause {{@ref:scope}}, signed on {{signing_date}}.",
    );
    const tokens = scanDirectives(doc).map((r) => `${r.kind}:${r.expr}`);

    expect(tokens).toContain("num:scope");
    expect(tokens).toContain("ref:scope");
    expect(tokens).toContain("placeholder:signing_date");
    // The numbering markers must not also be claimed as plain placeholders.
    expect(tokens.filter((token) => token === "placeholder:@num:scope")).toEqual([]);
  });

  test("still recognizes clause slots and plain fields alongside them", () => {
    const doc = docOf("Party {{tenant.name}} acts under {{@clause:Indemnity}}.");
    const tokens = scanDirectives(doc)
      .map((r) => `${r.kind}:${r.expr}`)
      .sort();

    expect(tokens).toEqual(["clause:Indemnity", "placeholder:tenant.name"]);
  });

  test("emits mid-line conditional markers as inline (block:false) ranges", () => {
    const doc = docOf(
      "the Buyer{{#if hasSpouse}} and their spouse{{#else}} alone{{/if}} hereby agrees.",
    );
    const tokens = scanDirectives(doc).map((r) => `${r.kind}:${r.expr}:${String(r.block)}`);

    expect(tokens).toEqual(["if:hasSpouse:false", "else::false", "endif::false"]);
  });

  test("inline range positions cover the markers in document order", () => {
    const doc = docOf("A{{#if x}}B{{/if}}C");
    const ranges = scanDirectives(doc);

    expect(ranges).toHaveLength(2);
    const [opener, closer] = ranges;
    expect(opener?.kind).toBe("if");
    expect(closer?.kind).toBe("endif");
    expect(doc.textBetween(opener?.from ?? 0, opener?.to ?? 0)).toBe("{{#if x}}");
    expect(doc.textBetween(closer?.from ?? 0, closer?.to ?? 0)).toBe("{{/if}}");
  });

  test("whole-paragraph directives keep block:true", () => {
    const doc = docOf("{{#if hasSpouse}}", "Spouse paragraph.", "{{/if}}");
    const blockKinds = scanDirectives(doc)
      .filter((r) => r.block)
      .map((r) => r.kind);

    expect(blockKinds).toEqual(["if", "endif"]);
  });

  test("emits mid-line each markers as inline (block:false) ranges", () => {
    const doc = docOf("Items: {{#each items}}{{items.name}}{{/each}} end.");
    const tokens = scanDirectives(doc).map((r) => `${r.kind}:${r.expr}:${String(r.block)}`);

    expect(tokens).toContain("each:items:false");
    expect(tokens).toContain("endeach::false");
    // The field inside the inline loop still gets its chip.
    expect(tokens).toContain("placeholder:items.name:false");
  });
});

describe("computeBlockDepths", () => {
  // Builds a block opener/closer range at a given position. Only `from`, `kind`,
  // and `block` drive the depth math, so the rest is filler.
  const blockRange = (from: number, kind: DirectiveRange["kind"]): DirectiveRange => ({
    from,
    to: from + 1,
    kind,
    expr: "",
    block: true,
  });

  test("assigns 0-based depth by containment", () => {
    // {{#each}} > {{#if}} > {{#if}}  (outer loop, two nested conditions)
    const ranges: DirectiveRange[] = [
      blockRange(0, "each"),
      blockRange(10, "if"),
      blockRange(20, "endif"),
      blockRange(30, "if"),
      blockRange(40, "endif"),
      blockRange(50, "endeach"),
    ];

    const depths = computeBlockDepths(ranges);

    expect(depths.get(0)).toBe(0); // each
    expect(depths.get(10)).toBe(1); // first nested if
    expect(depths.get(30)).toBe(1); // sibling nested if (back to depth 1)
  });

  test("deeply nested openers keep climbing (visual cap is the overlay's job)", () => {
    const ranges: DirectiveRange[] = [
      blockRange(0, "if"),
      blockRange(1, "each"),
      blockRange(2, "if"),
      blockRange(3, "each"),
      blockRange(4, "if"),
      blockRange(5, "each"),
    ];

    const depths = computeBlockDepths(ranges);

    expect([...depths.values()]).toEqual([0, 1, 2, 3, 4, 5]);
  });

  test("ignores order of input and inline (block:false) markers", () => {
    const ranges: DirectiveRange[] = [
      blockRange(30, "endif"),
      blockRange(0, "each"),
      { from: 5, to: 6, kind: "if", expr: "x", block: false }, // inline: no rail
      blockRange(10, "if"),
      blockRange(40, "endeach"),
    ];

    const depths = computeBlockDepths(ranges);

    expect(depths.get(0)).toBe(0); // each
    expect(depths.get(10)).toBe(1); // if nested inside each
    expect(depths.has(5)).toBe(false); // inline if excluded
  });

  test("tolerates unbalanced closers without going negative", () => {
    const ranges: DirectiveRange[] = [
      blockRange(0, "endif"), // stray closer, no opener
      blockRange(10, "if"),
      blockRange(20, "endif"),
    ];

    const depths = computeBlockDepths(ranges);

    expect(depths.get(10)).toBe(0);
  });

  test("kind-aware: a stray {{/each}} does not shrink a foreign block's depth", () => {
    // {{#if}} {{/each}}(stray, no open each) {{#each}} {{/if}}
    // A blind open/close counter decrements on the stray {{/each}} and pulls the
    // nested {{#each}} back to depth 0; kind-aware matching leaves it at depth 1.
    const ranges: DirectiveRange[] = [
      blockRange(0, "if"),
      blockRange(10, "endeach"), // stray: no open each to close
      blockRange(20, "each"),
      blockRange(30, "endif"),
    ];

    const depths = computeBlockDepths(ranges);

    expect(depths.get(0)).toBe(0); // outer if
    expect(depths.get(20)).toBe(1); // each is still nested inside the open if
  });

  test("kind-aware: interleaved if/each closers keep opener depths intact", () => {
    // {{#if}} {{#each}} {{/if}} {{/each}} (crossed nesting): the {{/if}} closes the
    // if and discards the improperly-nested each, but recorded depths do not shift.
    const ranges: DirectiveRange[] = [
      blockRange(0, "if"),
      blockRange(10, "each"),
      blockRange(20, "endif"),
      blockRange(30, "endeach"),
    ];

    const depths = computeBlockDepths(ranges);

    expect(depths.get(0)).toBe(0);
    expect(depths.get(10)).toBe(1);
  });
});
