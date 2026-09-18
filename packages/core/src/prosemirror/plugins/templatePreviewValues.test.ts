/**
 * Unit tests for the template fill preview plugin's entry tracking.
 * The plugin keeps a list of matched marker → value entries in sync
 * with the doc; the paged editor projects its overlay from these
 * entries (the inline decorations are derived from the same list), so
 * getting the entries right is the whole correctness story.
 */

import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { schema } from "../schema";
import {
  createTemplatePreviewValuesPlugin,
  templatePreviewValuesKey,
  templatePreviewValueText,
} from "./templatePreviewValues";
import type {
  TemplatePreviewEntry,
  TemplatePreviewHiddenRange,
  TemplatePreviewValues,
} from "./templatePreviewValues";

const docOf = (...paragraphs: string[]): PMNode =>
  schema.node(
    "doc",
    null,
    paragraphs.map((text) => schema.node("paragraph", null, text ? [schema.text(text)] : null)),
  );

const pushPreview = (state: EditorState, preview: TemplatePreviewValues | null): EditorState =>
  state.apply(state.tr.setMeta(templatePreviewValuesKey, { preview }));

const makeState = (doc: PMNode, preview: TemplatePreviewValues | null): EditorState => {
  const plugin = createTemplatePreviewValuesPlugin();
  return pushPreview(EditorState.create({ doc, plugins: [plugin] }), preview);
};

const getEntries = (state: EditorState): readonly TemplatePreviewEntry[] =>
  templatePreviewValuesKey.getState(state)?.entries ?? [];

const getHidden = (state: EditorState): readonly TemplatePreviewHiddenRange[] =>
  templatePreviewValuesKey.getState(state)?.hidden ?? [];

/** Every decoration as a bare position pair, in document order. */
const decorationRanges = (state: EditorState): { from: number; to: number }[] =>
  (templatePreviewValuesKey.getState(state)?.decorationSet.find() ?? []).map(({ from, to }) => ({
    from,
    to,
  }));

/** Node-decoration positions of the doc's nth top-level block. */
const blockRange = (doc: PMNode, index: number): { from: number; to: number } => {
  let from = 0;
  for (let before = 0; before < index; before += 1) {
    from += doc.child(before).nodeSize;
  }
  return { from, to: from + doc.child(index).nodeSize };
};

const sliceFromTo = (doc: PMNode, from: number, to: number): string =>
  doc.textBetween(from, to, "");

describe("templatePreviewValues: entry tracking", () => {
  test("exposes one entry per matched placeholder, spanning the marker", () => {
    const doc = docOf(
      "Tenant {{tenant.name}} signs on {{signing_date}}.",
      "Landlord {{landlord.name}} agrees.",
    );
    const state = makeState(doc, {
      values: { "tenant.name": "Pavel Novák", signing_date: "2026-06-10" },
      mode: "highlighted",
    });

    const entries = getEntries(state);
    expect(entries.map((e) => `${e.expr}=${templatePreviewValueText(e.value)}`)).toEqual([
      "tenant.name=Pavel Novák",
      "signing_date=2026-06-10",
    ]);
    for (const entry of entries) {
      expect(sliceFromTo(state.doc, entry.from, entry.to)).toBe(`{{${entry.expr}}}`);
    }
  });

  test("rich values surface on entries verbatim; rich values with no text are skipped", () => {
    const doc = docOf("Company {{company}} and {{empty}} sign.");
    const richValue = {
      runs: [{ text: "Acme", bold: true }, { text: ", Poznań" }],
    };
    const state = makeState(doc, {
      values: { company: richValue, empty: { runs: [{ text: "" }] } },
      mode: "plain",
    });

    const entries = getEntries(state);
    expect(entries).toHaveLength(1);
    expect(entries[0]!.expr).toBe("company");
    expect(entries[0]!.value).toEqual(richValue);
  });

  test("skips empty values, unmatched fields, and structural directives", () => {
    const doc = docOf(
      'Field {{ tenant.name }} and clause {{ clause("Indemnity") }}.',
      "{% if premium %}",
      "Premium terms for {{ tenant.name }}.",
      "{% endif %}",
    );
    const state = makeState(doc, {
      values: {
        "tenant.name": "",
        "landlord.name": "Unused",
        // A clause slot keys by SLOT NAME, not the `@clause:` patch key,
        // so this value never matches the clause("Indemnity") marker.
        "@clause:Indemnity": "Not a field",
        premium: "true",
      },
      mode: "plain",
    });

    expect(getEntries(state)).toEqual([]);
  });

  test("previews a linked clause slot, keyed by slot name", () => {
    const doc = docOf('Field {{ tenant.name }} and clause {{ clause("Indemnity") }}.');
    const state = makeState(doc, {
      values: {
        "tenant.name": "Pavel Novák",
        Indemnity: "The Supplier shall indemnify the Customer.",
      },
      mode: "highlighted",
    });

    const entries = getEntries(state);
    expect(entries.map((e) => `${e.expr}=${templatePreviewValueText(e.value)}`)).toEqual([
      "tenant.name=Pavel Novák",
      "Indemnity=The Supplier shall indemnify the Customer.",
    ]);
    const clause = entries.find((e) => e.expr === "Indemnity")!;
    expect(sliceFromTo(state.doc, clause.from, clause.to)).toBe('{{ clause("Indemnity") }}');
  });

  test("renders multi-paragraph clause text as one wrapping run", () => {
    const doc = docOf('Clause {{ clause("Terms") }}.');
    const state = makeState(doc, {
      values: { Terms: "First paragraph.\nSecond paragraph." },
      mode: "highlighted",
    });

    const entries = getEntries(state);
    expect(entries).toHaveLength(1);
    expect(templatePreviewValueText(entries[0]!.value)).toBe("First paragraph.\nSecond paragraph.");
  });

  test("skips an unlinked clause slot (no value supplied)", () => {
    const doc = docOf('Clause {{ clause("Missing") }}.');
    const state = makeState(doc, {
      values: { tenant: "Pavel Novák" },
      mode: "plain",
    });

    expect(getEntries(state)).toEqual([]);
  });

  test("builds an inline (hide) + widget (value) decoration pair per entry", () => {
    const doc = docOf("Tenant {{tenant.name}} signs.");
    const state = makeState(doc, {
      values: { "tenant.name": "Pavel Novák" },
      mode: "plain",
    });

    const decorationSet = templatePreviewValuesKey.getState(state)?.decorationSet;
    const decorations = decorationSet?.find() ?? [];
    expect(decorations).toHaveLength(2);
  });

  test("recomputes entry positions when the doc is edited before a marker", () => {
    const doc = docOf("Tenant {{tenant.name}} signs.");
    let state = makeState(doc, {
      values: { "tenant.name": "Pavel Novák" },
      mode: "plain",
    });

    state = state.apply(state.tr.insertText("Dear ", 1));

    const entries = getEntries(state);
    expect(entries).toHaveLength(1);
    expect(sliceFromTo(state.doc, entries[0]!.from, entries[0]!.to)).toBe("{{tenant.name}}");
  });

  test("drops entries whose marker is broken by the edit", () => {
    const doc = docOf("Tenant {{tenant.name}} signs.");
    let state = makeState(doc, {
      values: { "tenant.name": "Pavel Novák" },
      mode: "plain",
    });
    expect(getEntries(state)).toHaveLength(1);

    // Break the closing braces; the marker no longer parses.
    const entry = getEntries(state)[0]!;
    state = state.apply(state.tr.delete(entry.to - 1, entry.to));

    expect(getEntries(state)).toEqual([]);
  });

  test("clearing the preview empties entries and decorations", () => {
    const doc = docOf("Tenant {{tenant.name}} signs.");
    let state = makeState(doc, {
      values: { "tenant.name": "Pavel Novák" },
      mode: "highlighted",
    });
    expect(getEntries(state)).toHaveLength(1);

    state = state.apply(state.tr.setMeta(templatePreviewValuesKey, { preview: null }));

    expect(getEntries(state)).toEqual([]);
    expect(templatePreviewValuesKey.getState(state)?.decorationSet.find() ?? []).toHaveLength(0);
  });
});

/**
 * Conditional hiding: the host reports which `{% if %}` blocks apply and the
 * preview drops the ones that do not, opener through closer. folio pairs the
 * markers and evaluates nothing, so these tests are about pairing, span, and
 * nesting rather than about truthiness.
 */
describe("templatePreviewValues: conditional hiding", () => {
  const conditionalDoc = () =>
    docOf("Intro.", "{% if premium %}", "Premium terms.", "{% endif %}", "Tail.");

  test("hides a false block from its opener through its closer", () => {
    const doc = conditionalDoc();
    const state = makeState(doc, {
      values: {},
      mode: "plain",
      conditions: { premium: false },
    });

    const hidden = getHidden(state);
    expect(hidden).toHaveLength(1);
    expect(hidden[0]!.expr).toBe("premium");
    expect(sliceFromTo(doc, hidden[0]!.from, hidden[0]!.to)).toBe(
      "{% if premium %}Premium terms.{% endif %}",
    );
    // Whole blocks, so the hidden paragraphs leave no empty lines behind.
    expect(decorationRanges(state)).toEqual([
      blockRange(doc, 1),
      blockRange(doc, 2),
      blockRange(doc, 3),
    ]);
  });

  test("leaves a true block, and one no condition mentions, as authored", () => {
    const doc = conditionalDoc();
    for (const conditions of [{ premium: true }, { other: false }]) {
      const state = makeState(doc, { values: {}, mode: "highlighted", conditions });
      expect(getHidden(state)).toEqual([]);
      expect(decorationRanges(state)).toEqual([]);
    }
  });

  test("hides only its own span when the conditional is inline", () => {
    const doc = docOf("Fee {% if waived %}is waived{% endif %} on signature.");
    const state = makeState(doc, {
      values: {},
      mode: "plain",
      conditions: { waived: false },
    });

    const hidden = getHidden(state);
    expect(hidden).toHaveLength(1);
    expect(sliceFromTo(doc, hidden[0]!.from, hidden[0]!.to)).toBe(
      "{% if waived %}is waived{% endif %}",
    );
    // The paragraph holds other text, so only the marked span is decorated.
    expect(decorationRanges(state)).toEqual([{ from: hidden[0]!.from, to: hidden[0]!.to }]);
    expect(decorationRanges(state)).not.toEqual([blockRange(doc, 0)]);
  });

  const nestedDoc = () =>
    docOf(
      "{% if outer %}",
      "Outer body.",
      "{% if inner %}",
      "Inner body.",
      "{% endif %}",
      "{% endif %}",
      "Tail.",
    );

  test("a hidden outer block subsumes a hidden inner one", () => {
    const doc = nestedDoc();
    const state = makeState(doc, {
      values: {},
      mode: "plain",
      conditions: { outer: false, inner: false },
    });

    expect(getHidden(state).map((range) => range.expr)).toEqual(["outer"]);
    expect(decorationRanges(state)).toEqual(
      [0, 1, 2, 3, 4, 5].map((index) => blockRange(doc, index)),
    );
  });

  test("hides an inner false block inside a visible outer one", () => {
    const doc = nestedDoc();
    // `highlighted` keeps every tag, so a true outer block is untouched here
    // and the assertion is about nesting alone.
    const state = makeState(doc, {
      values: {},
      mode: "highlighted",
      conditions: { outer: true, inner: false },
    });

    expect(getHidden(state).map((range) => range.expr)).toEqual(["inner"]);
    expect(decorationRanges(state)).toEqual([2, 3, 4].map((index) => blockRange(doc, index)));
  });

  test("ignores an opener with no closer, and a closer of another kind", () => {
    const unpaired = makeState(docOf("{% if premium %}", "Premium terms."), {
      values: {},
      mode: "plain",
      conditions: { premium: false },
    });
    expect(getHidden(unpaired)).toEqual([]);
    expect(decorationRanges(unpaired)).toEqual([]);

    const foreignCloser = makeState(docOf("{% if premium %}", "Premium terms.", "{% endfor %}"), {
      values: {},
      mode: "plain",
      conditions: { premium: false },
    });
    expect(getHidden(foreignCloser)).toEqual([]);
    expect(decorationRanges(foreignCloser)).toEqual([]);
  });

  test("does not substitute a value inside a hidden block", () => {
    const doc = docOf(
      "{% if premium %}",
      "Premium support for {{tenant.name}}.",
      "{% endif %}",
      "Signed by {{tenant.name}}.",
    );
    const state = makeState(doc, {
      values: { "tenant.name": "Pavel Novák" },
      mode: "plain",
      conditions: { premium: false },
    });

    const entries = getEntries(state);
    expect(entries).toHaveLength(1);
    expect(entries[0]!.from).toBeGreaterThan(getHidden(state)[0]!.to);
  });

  test("switching conditions between pushes updates the decorations", () => {
    const doc = conditionalDoc();
    let state = makeState(doc, {
      values: {},
      mode: "highlighted",
      conditions: { premium: false },
    });
    expect(decorationRanges(state)).toHaveLength(3);

    state = pushPreview(state, { values: {}, mode: "highlighted", conditions: { premium: true } });
    expect(getHidden(state)).toEqual([]);
    expect(decorationRanges(state)).toEqual([]);

    state = pushPreview(state, { values: {}, mode: "highlighted", conditions: { premium: false } });
    expect(getHidden(state)).toHaveLength(1);
    expect(decorationRanges(state)).toHaveLength(3);

    // Same verdict, other mode: `plain` now hides the two tag paragraphs of a
    // block it keeps, so the decorations follow a mode switch too.
    state = pushPreview(state, { values: {}, mode: "plain", conditions: { premium: true } });
    expect(decorationRanges(state)).toEqual([blockRange(doc, 1), blockRange(doc, 3)]);
  });

  test("clearing the preview drops the hidden ranges", () => {
    const state = makeState(conditionalDoc(), {
      values: {},
      mode: "highlighted",
      conditions: { premium: false },
    });
    expect(getHidden(state)).toHaveLength(1);

    const cleared = pushPreview(state, null);
    expect(getHidden(cleared)).toEqual([]);
    expect(decorationRanges(cleared)).toEqual([]);
  });
});

/**
 * A condition tag may carry a filter chain, the way a `{% for x in xs | chain %}`
 * opener already does, while the host keys its verdicts by the bare field path
 * in front of the chain. Resolution is: the expression as written, then that
 * path. A real expression has no such path and keeps exact matching.
 */
describe("templatePreviewValues: condition keys", () => {
  const chainDoc = (expr: string) => docOf(`{% if ${expr} %}`, "Body.", "{% endif %}");

  const hiddenExprs = (doc: PMNode, conditions: Record<string, boolean>): string[] =>
    getHidden(makeState(doc, { values: {}, mode: "highlighted", conditions })).map(
      (range) => range.expr,
    );

  test("matches a filter-chain tag by the bare path in front of the chain", () => {
    const expr = 'buyer_is_a_consumer | checkbox | label("Buyer is a consumer")';
    const doc = chainDoc(expr);

    expect(hiddenExprs(doc, { buyer_is_a_consumer: false })).toEqual([expr]);
    // The verdict travels under the path, but the range reports the expression
    // as authored, which is what the document holds.
    expect(hiddenExprs(doc, { buyer_is_a_consumer: true })).toEqual([]);
  });

  test("a quoted pipe inside a filter argument does not become a key boundary", () => {
    const expr = 'consent_given | label("Yes | No") | ai("Did the buyer consent, yes|no?")';
    const doc = chainDoc(expr);

    expect(hiddenExprs(doc, { consent_given: false })).toEqual([expr]);
    // Nothing keyed by a fragment of the chain resolves.
    expect(hiddenExprs(doc, { 'consent_given | label("Yes ': false })).toEqual([]);
    expect(hiddenExprs(doc, { consent_given_label: false })).toEqual([]);
  });

  test("keeps a real expression on exact matching", () => {
    for (const expr of ["a and b", "items|length > 0", 'ai("a|b")']) {
      const doc = chainDoc(expr);
      // The expression as written is still a key.
      expect(hiddenExprs(doc, { [expr]: false })).toEqual([expr]);
      // Its leading token is not: `items|length > 0` is a comparison, not a
      // field path carrying a chain of known filters.
      expect(hiddenExprs(doc, { a: false, items: false, ai: false })).toEqual([]);
    }
  });

  test("leaves a block no verdict mentions exactly as authored", () => {
    const doc = chainDoc("buyer_is_a_consumer | checkbox");

    expect(hiddenExprs(doc, {})).toEqual([]);
    expect(hiddenExprs(doc, { other_field: false })).toEqual([]);
    // A path that resolves to something other than a boolean is silence, not a
    // verdict, so an inherited property cannot hide a block.
    expect(hiddenExprs(doc, { constructor: false } as Record<string, boolean>)).toEqual([]);
  });
});

/**
 * Tag paragraphs: `plain` mode approximates the generated document, so a block
 * the host has ruled on loses its scaffolding — the whole block when it does
 * not apply, only the tag lines when it does. `highlighted` mode exists to show
 * that scaffolding and keeps it.
 */
describe("templatePreviewValues: directive tag lines", () => {
  const branchDoc = () =>
    docOf(
      "Intro.",
      "{% if premium %}",
      "Premium terms.",
      "{% else %}",
      "Standard terms.",
      "{% endif %}",
      "Tail.",
    );

  test("plain mode hides the tag lines of a block that applies", () => {
    const doc = branchDoc();
    const state = makeState(doc, {
      values: {},
      mode: "plain",
      conditions: { premium: true },
    });

    // Opener, else and closer go; both branch bodies stay, because folio does
    // not evaluate which branch the host meant.
    expect(getHidden(state).map((range) => range.expr)).toEqual(["premium", "premium", "premium"]);
    expect(decorationRanges(state)).toEqual([1, 3, 5].map((index) => blockRange(doc, index)));
  });

  test("highlighted mode keeps the tag lines of a block that applies", () => {
    const state = makeState(branchDoc(), {
      values: {},
      mode: "highlighted",
      conditions: { premium: true },
    });

    expect(getHidden(state)).toEqual([]);
    expect(decorationRanges(state)).toEqual([]);
  });

  test("a block that does not apply loses body and tags alike, in both modes", () => {
    const doc = branchDoc();
    for (const mode of ["plain", "highlighted"] as const) {
      const state = makeState(doc, { values: {}, mode, conditions: { premium: false } });
      expect(getHidden(state)).toHaveLength(1);
      expect(decorationRanges(state)).toEqual(
        [1, 2, 3, 4, 5].map((index) => blockRange(doc, index)),
      );
    }
  });

  test("keeps a paragraph an inline tag shares with body text", () => {
    // The opener sits mid-sentence, so hiding its paragraph would take the
    // authored text with it; only the tag span goes.
    const doc = docOf("Fee {% if waived %}is waived{% endif %} on signature.");
    const state = makeState(doc, {
      values: {},
      mode: "plain",
      conditions: { waived: true },
    });

    const hidden = getHidden(state);
    expect(hidden.map((range) => sliceFromTo(doc, range.from, range.to))).toEqual([
      "{% if waived %}",
      "{% endif %}",
    ]);
    expect(decorationRanges(state)).not.toContainEqual(blockRange(doc, 0));
  });

  test("hides the tags of a true block nested in a visible outer one", () => {
    const doc = docOf(
      "{% if outer %}",
      "Outer body.",
      "{% if inner %}",
      "Inner body.",
      "{% endif %}",
      "{% endif %}",
    );
    const state = makeState(doc, {
      values: {},
      mode: "plain",
      conditions: { inner: true },
    });

    // The outer block carries no verdict, so its tags stay; the inner block's go.
    expect(decorationRanges(state)).toEqual([2, 4].map((index) => blockRange(doc, index)));
  });
});
