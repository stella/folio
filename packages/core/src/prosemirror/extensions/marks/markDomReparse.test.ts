/**
 * A mark's DOM parses back to that mark and nothing else.
 *
 * ProseMirror reads the editor's own DOM back whenever a native edit touches
 * it, and a style parse rule matches any element carrying the declaration. A
 * tracked insertion painted with `text-decoration: underline` therefore came
 * back underlined, and a tracked deletion struck through, on the untouched text
 * next to a composition. Rejecting that revision then left real formatting
 * behind. These hold every mark in the schema to the rule, not just the two
 * that broke it.
 */

import { describe, expect, test } from "bun:test";
import type { Attrs, MarkType, StyleParseRule } from "prosemirror-model";

import { schema } from "../../schema";

/** Attributes for the marks whose DOM differs by value, or that require one. */
const ATTR_SAMPLES: Readonly<Record<string, readonly Attrs[]>> = {
  insertion: [
    { revisionId: 1, author: "Ada", date: null, provenance: "user" },
    { revisionId: 2, author: "Ada", date: null, provenance: "suggested" },
  ],
  deletion: [
    { revisionId: 1, author: "Ada", date: null, provenance: "user" },
    { revisionId: 2, author: "Ada", date: null, provenance: "suggested" },
  ],
};

const styleRules = (markType: MarkType): StyleParseRule[] =>
  (markType.spec.parseDOM ?? []).filter(
    (rule): rule is StyleParseRule => "style" in rule,
  );

const OTHER_MARK_RULES = Object.values(schema.marks).flatMap((markType) =>
  styleRules(markType).map((rule) => ({ owner: markType.name, rule })),
);

/** A browser hands a parse rule `rgb(…)` for a hex colour, but keeps `var(…)` verbatim. */
const asBrowserReadsIt = (value: string): string => {
  const hex = /^#(?<rgb>[0-9a-f]{6})$/iu.exec(value)?.groups?.["rgb"];
  if (hex === undefined) return value;
  const channel = (start: number) =>
    Number.parseInt(hex.slice(start, start + 2), 16);
  return `rgb(${channel(0)}, ${channel(2)}, ${channel(4)})`;
};

const declarationsOf = (style: string): [string, string][] =>
  style
    .split(";")
    .map((declaration) => declaration.trim())
    .filter(Boolean)
    .map((declaration) => {
      const colon = declaration.indexOf(":");
      return [
        declaration.slice(0, colon).trim(),
        asBrowserReadsIt(declaration.slice(colon + 1).trim()),
      ];
    });

const ruleAccepts = (rule: StyleParseRule, property: string, value: string) => {
  const [ruleProperty, wanted] = (rule.style ?? "").split("=");
  if (ruleProperty !== property) return false;
  if (wanted !== undefined && wanted !== value) return false;
  return rule.getAttrs === undefined || rule.getAttrs(value) !== false;
};

/** Every inline style a mark's `toDOM` writes, at any depth of its DOM spec. */
const stylesOf = (spec: unknown): string[] => {
  if (!Array.isArray(spec)) return [];
  return spec.flatMap((part: unknown) => {
    if (Array.isArray(part)) return stylesOf(part);
    if (typeof part !== "object" || part === null) return [];
    const style = (part as Record<string, unknown>)["style"];
    return typeof style === "string" ? [style] : [];
  });
};

const samplesOf = (markType: MarkType): readonly Attrs[] => {
  const samples = ATTR_SAMPLES[markType.name];
  if (samples !== undefined) return samples;
  return [{}];
};

describe("mark DOM reparse", () => {
  for (const markType of Object.values(schema.marks)) {
    test(`${markType.name} is not read back as another mark`, () => {
      const toDOM = markType.spec.toDOM;
      if (toDOM === undefined) return;
      for (const attrs of samplesOf(markType)) {
        let spec: unknown;
        try {
          spec = toDOM(markType.create(attrs), true);
        } catch {
          // Its attrs are required and carry no sample: its DOM has no default.
          continue;
        }
        for (const style of stylesOf(spec)) {
          for (const [property, value] of declarationsOf(style)) {
            const claimedBy = OTHER_MARK_RULES.filter(
              ({ owner, rule }) =>
                owner !== markType.name && ruleAccepts(rule, property, value),
            ).map(({ owner }) => owner);
            expect(
              claimedBy,
              `${markType.name} writes "${property}: ${value}"`,
            ).toEqual([]);
          }
        }
      }
    });
  }
});
