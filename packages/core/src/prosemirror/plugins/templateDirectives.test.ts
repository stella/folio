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
  test("recognizes num() and ref() numbering markers as their own kinds", () => {
    const doc = docOf(
      'Clause {{ num("scope") }}. Scope of authority.',
      'As set out in Clause {{ ref("scope") }}, signed on {{ signing_date }}.',
    );
    const tokens = scanDirectives(doc).map((r) => `${r.kind}:${r.expr}`);

    expect(tokens).toContain("num:scope");
    expect(tokens).toContain("ref:scope");
    // The numbering calls must not also be claimed as plain placeholders.
    expect(tokens.filter((token) => token.startsWith("placeholder:"))).toEqual([
      "placeholder:signing_date",
    ]);
  });

  test("still recognizes clause slots and plain fields alongside them", () => {
    const doc = docOf('Party {{ tenant.name }} acts under {{ clause("Indemnity") }}.');
    const tokens = scanDirectives(doc)
      .map((r) => `${r.kind}:${r.expr}`)
      .sort();

    expect(tokens).toEqual(["clause:Indemnity", "placeholder:tenant.name"]);
  });

  test("emits mid-line conditional markers as inline (block:false) ranges", () => {
    const doc = docOf(
      "the Buyer{% if hasSpouse %} and their spouse{% else %} alone{% endif %} hereby agrees.",
    );
    const tokens = scanDirectives(doc).map((r) => `${r.kind}:${r.expr}:${String(r.block)}`);

    expect(tokens).toEqual(["if:hasSpouse:false", "else::false", "endif::false"]);
  });

  test("inline range positions cover the markers in document order", () => {
    const doc = docOf("A{% if x %}B{% endif %}C");
    const ranges = scanDirectives(doc);

    expect(ranges).toHaveLength(2);
    const [opener, closer] = ranges;
    expect(opener?.kind).toBe("if");
    expect(closer?.kind).toBe("endif");
    expect(doc.textBetween(opener?.from ?? 0, opener?.to ?? 0)).toBe("{% if x %}");
    expect(doc.textBetween(closer?.from ?? 0, closer?.to ?? 0)).toBe("{% endif %}");
  });

  test("whole-paragraph directives keep block:true", () => {
    const doc = docOf("{% if hasSpouse %}", "Spouse paragraph.", "{% endif %}");
    const blockKinds = scanDirectives(doc)
      .filter((r) => r.block)
      .map((r) => r.kind);

    expect(blockKinds).toEqual(["if", "endif"]);
  });

  test("whole-paragraph directive fields use the atomic node endpoint", () => {
    const field = schema.node("field", {
      fieldType: "UNKNOWN",
      instruction: " QUOTE ",
      displayText: "{% if hasSpouse %}",
      fieldKind: "simple",
    });
    const doc = schema.node("doc", null, [schema.node("paragraph", null, [field])]);

    const [range] = scanDirectives(doc);

    expect(range?.block).toBe(true);
    expect(range?.from).toBe(1);
    expect(range?.to).toBe(1 + field.nodeSize);
    expect(doc.textBetween(range?.from ?? 0, range?.to ?? 0)).toBe("{% if hasSpouse %}");
  });

  test("emits mid-line for markers as inline (block:false) ranges", () => {
    const doc = docOf("Items: {% for item in items %}{{ item.name }}{% endfor %} end.");
    const tokens = scanDirectives(doc).map((r) => `${r.kind}:${r.expr}:${String(r.block)}`);

    expect(tokens).toContain("for:items:false");
    expect(tokens).toContain("endfor::false");
    // The field inside the inline loop still gets its chip.
    expect(tokens).toContain("placeholder:item.name:false");
  });

  test("a for range carries the array path as expr and the loop alias", () => {
    // Hosts read `expr` as the array the loop walks and `alias` as the name its
    // body binds each element to; splitting them is what lets a host resolve
    // `{{ item.name }}` back to `items`.
    const doc = docOf("{% for item in contracts.risks %}", "Risk.", "{% endfor %}");

    const opener = scanDirectives(doc).find((r) => r.kind === "for");

    expect(opener).toMatchObject({ expr: "contracts.risks", alias: "item", block: true });
  });

  test("only a for range carries an alias", () => {
    const doc = docOf("{% if premium %}", "Premium.", "{% endif %}");

    expect(scanDirectives(doc).every((r) => r.alias === undefined)).toBe(true);
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
    // {% for %} > {% if %} > {% if %}  (outer loop, two nested conditions)
    const ranges: DirectiveRange[] = [
      blockRange(0, "for"),
      blockRange(10, "if"),
      blockRange(20, "endif"),
      blockRange(30, "if"),
      blockRange(40, "endif"),
      blockRange(50, "endfor"),
    ];

    const depths = computeBlockDepths(ranges);

    expect(depths.get(0)).toBe(0); // for
    expect(depths.get(10)).toBe(1); // first nested if
    expect(depths.get(30)).toBe(1); // sibling nested if (back to depth 1)
  });

  test("deeply nested openers keep climbing (visual cap is the overlay's job)", () => {
    const ranges: DirectiveRange[] = [
      blockRange(0, "if"),
      blockRange(1, "for"),
      blockRange(2, "if"),
      blockRange(3, "for"),
      blockRange(4, "if"),
      blockRange(5, "for"),
    ];

    const depths = computeBlockDepths(ranges);

    expect([...depths.values()]).toEqual([0, 1, 2, 3, 4, 5]);
  });

  test("ignores order of input and inline (block:false) markers", () => {
    const ranges: DirectiveRange[] = [
      blockRange(30, "endif"),
      blockRange(0, "for"),
      { from: 5, to: 6, kind: "if", expr: "x", block: false }, // inline: no rail
      blockRange(10, "if"),
      blockRange(40, "endfor"),
    ];

    const depths = computeBlockDepths(ranges);

    expect(depths.get(0)).toBe(0); // for
    expect(depths.get(10)).toBe(1); // if nested inside for
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

  test("kind-aware: a stray {% endfor %} does not shrink a foreign block's depth", () => {
    // {% if %} {% endfor %}(stray, no open for) {% for %} {% endif %}
    // A blind open/close counter decrements on the stray {% endfor %} and pulls
    // the nested {% for %} back to depth 0; kind-aware matching leaves it at 1.
    const ranges: DirectiveRange[] = [
      blockRange(0, "if"),
      blockRange(10, "endfor"), // stray: no open for to close
      blockRange(20, "for"),
      blockRange(30, "endif"),
    ];

    const depths = computeBlockDepths(ranges);

    expect(depths.get(0)).toBe(0); // outer if
    expect(depths.get(20)).toBe(1); // for is still nested inside the open if
  });

  test("kind-aware: interleaved if/for closers keep opener depths intact", () => {
    // {% if %} {% for %} {% endif %} {% endfor %} (crossed nesting): the
    // {% endif %} closes the if and discards the improperly-nested for, but the
    // recorded depths do not shift.
    const ranges: DirectiveRange[] = [
      blockRange(0, "if"),
      blockRange(10, "for"),
      blockRange(20, "endif"),
      blockRange(30, "endfor"),
    ];

    const depths = computeBlockDepths(ranges);

    expect(depths.get(0)).toBe(0);
    expect(depths.get(10)).toBe(1);
  });
});
