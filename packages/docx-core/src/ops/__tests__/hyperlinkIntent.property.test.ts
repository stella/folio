import { panic } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { Document, Hyperlink, ParagraphContent, Run } from "../../model/document";
import { allocateEditorIntentIds, compileEditorIntent } from "../editorIntent";
import type { HyperlinkEditorIntent } from "../hyperlinkIntent";
import { applyDocumentOps } from "../apply";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { OP_STORIES } from "../types";
import { paragraphVisibleText } from "../editorIntent";

setDefaultTimeout(propertyTestTimeout(30_000));

const original = (styleId = "Hyperlink"): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "12345678",
          content: [
            {
              type: "run",
              formatting: { bold: true, color: { rgb: "123456" } },
              content: [{ type: "text", text: "ab" }],
            },
            {
              type: "hyperlink",
              href: "https://old.example/",
              target: "_blank",
              children: [
                {
                  type: "run",
                  formatting: { italic: true, styleId },
                  content: [{ type: "text", text: "😀cd" }],
                },
              ],
            },
          ],
        },
      ],
    },
  },
});
const positions = [0, 1, 2, 4, 5, 6] as const;
const at = (offset: number) => ({ story: OP_STORIES.MAIN, blockId: "12345678", offset });
type RangeOffsets = { from: number; to: number };
const factories = {
  setHyperlink: ({ from, to }: RangeOffsets): HyperlinkEditorIntent => ({
    type: "setHyperlink",
    from: at(from),
    to: at(to),
    href: "#anchor",
    tooltip: "Tip",
  }),
  removeHyperlink: ({ from, to }: RangeOffsets): HyperlinkEditorIntent => ({
    type: "removeHyperlink",
    from: at(from),
    to: at(to),
  }),
  insertHyperlink: ({ from, to }: RangeOffsets): HyperlinkEditorIntent => ({
    type: "insertHyperlink",
    from: at(from),
    to: at(to),
    href: "https://new.example/",
    text: "New😀",
    tooltip: "Tip",
  }),
} satisfies Record<HyperlinkEditorIntent["type"], (range: RangeOffsets) => HyperlinkEditorIntent>;

type Token = {
  text: string;
  formatting?: Run["formatting"];
  link?: Omit<Hyperlink, "type" | "children">;
};
const tokens = (items: readonly ParagraphContent[], link?: Token["link"]): Token[] => {
  const result: Token[] = [];
  for (const item of items) {
    if (item.type === "hyperlink") {
      const { type: _type, children, ...target } = item;
      result.push(...tokens(children, target));
    } else if (item.type === "run") {
      for (const child of item.content)
        if (child.type === "text") {
          for (const text of child.text.split(""))
            result.push({ text, formatting: item.formatting, link });
        }
    }
  }
  return result;
};
const expectedTokens = (intent: HyperlinkEditorIntent, before: Token[]): Token[] => {
  const start = intent.from.offset;
  const end = intent.to.offset;
  if (intent.type === "insertHyperlink")
    return before.slice(0, start).concat(
      intent.text.split("").map((text) => ({
        text,
        formatting: undefined,
        link: { href: intent.href, tooltip: intent.tooltip },
      })),
      before.slice(end),
    );
  return before.map((token, index) => {
    if (index < start || index >= end) return token;
    const formatting = { ...token.formatting };
    if (intent.type === "setHyperlink") delete formatting.color;
    else if (
      token.link !== undefined &&
      (formatting.styleId === "Hyperlink" || formatting.styleId === intent.hyperlinkStyleId)
    )
      delete formatting.styleId;
    return {
      text: token.text,
      formatting: Object.keys(formatting).length === 0 ? undefined : formatting,
      link:
        intent.type === "setHyperlink" ? { anchor: "anchor", tooltip: intent.tooltip } : undefined,
    };
  });
};

test("hyperlink intents have exact inverses and refuse unrepresentable suggestions", () => {
  assertProperty(
    fc.property(
      fc.constantFrom(...positions),
      fc.constantFrom(...positions),
      fc.constantFrom("Hyperlink", "AliasLink"),
      (left, right, styleId) => {
        const from = Math.min(left, right);
        const to = Math.max(left, right);
        for (const factory of Object.values(factories)) {
          const document = original(styleId);
          const candidate = factory({ from, to });
          const intent =
            candidate.type === "removeHyperlink"
              ? { ...candidate, hyperlinkStyleId: styleId }
              : candidate;
          const ids = allocateEditorIntentIds(document, intent);
          const direct = compileEditorIntent(document, {
            intent,
            mode: { type: "editing", newIds: ids.newIds },
          }).unwrap();
          const tracked = compileEditorIntent(document, {
            intent,
            mode: {
              type: "suggesting",
              revision: { id: ids.revisionId, author: "Reviewer", date: "2026-10-04T00:00:00Z" },
              newIds: ids.newIds,
            },
          });
          expect(tracked.isErr()).toBe(true);
          if (tracked.isErr()) {
            expect(tracked.error.reason).toBe("untrackable");
            expect(tracked.error.message).toContain("serializable wrapper review provenance");
          }
          expect(document).toStrictEqual(original(styleId));
          const expectedText =
            intent.type === "insertHyperlink"
              ? "ab😀cd".slice(0, from) + intent.text + "ab😀cd".slice(to)
              : "ab😀cd";
          for (const compiled of [direct]) {
            const applied = applyDocumentOps(document, compiled.ops).unwrap();
            const paragraph = applied.document.package.document.content.at(0);
            expect(
              paragraph?.type === "paragraph" ? paragraphVisibleText(paragraph) : undefined,
            ).toBe(expectedText);
            expect(paragraph?.type === "paragraph" ? tokens(paragraph.content) : undefined).toEqual(
              expectedTokens(
                intent,
                tokens(
                  original(styleId).package.document.content.flatMap((block) =>
                    block.type === "paragraph" ? block.content : [],
                  ),
                ),
              ),
            );
            const undone = applyDocumentOps(applied.document, applied.inverse).unwrap();
            expect(undone.document).toStrictEqual(document);
            expect(
              applyDocumentOps(undone.document, undone.inverse).unwrap().document,
            ).toStrictEqual(applied.document);
          }
        }
      },
    ),
    {
      numRuns: 25,
      id: "hyperlink intents have exact inverses and refuse unrepresentable suggestions",
    },
  );
});

test("cross-paragraph hyperlinks preserve formatting and marker identities through exact inverses", () => {
  assertProperty(
    fc.property(fc.constantFrom(...positions), fc.constantFrom(...positions), (left, right) => {
      for (const factory of Object.values(factories)) {
        const document = original();
        const first = document.package.document.content.at(0);
        if (first?.type !== "paragraph") panic("Missing hyperlink fixture paragraph");
        first.paraId = "1B123456";
        first.content.unshift({ type: "bookmarkStart", id: 7, name: "anchor" });
        const second = structuredClone(first);
        second.paraId = "2D123456";
        second.content.shift();
        second.content.push({ type: "bookmarkEnd", id: 7 });
        document.package.document.content.push(second);
        const base = factory({ from: left, to: right });
        const intent = {
          ...base,
          from: { ...base.from, blockId: "1b123456" },
          to: { ...base.to, blockId: "2d123456" },
        };
        const compiled = compileEditorIntent(document, {
          intent,
          mode: { type: "editing" },
        }).unwrap();
        const applied = applyDocumentOps(document, compiled.ops).unwrap();
        const content = applied.document.package.document.content.flatMap((block) =>
          block.type === "paragraph" ? block.content : [],
        );
        const before = tokens(first.content).concat(tokens(second.content));
        expect(tokens(content)).toEqual(
          expectedTokens({ ...intent, to: { ...intent.to, offset: 6 + right } }, before),
        );
        const markers = (items: readonly ParagraphContent[]): unknown[] =>
          items.flatMap((item) => {
            if (item.type === "hyperlink") return markers(item.children);
            if (item.type === "bookmarkStart" || item.type === "bookmarkEnd") return [item];
            return [];
          });
        expect(markers(content)).toEqual([
          { type: "bookmarkStart", id: 7, name: "anchor" },
          { type: "bookmarkEnd", id: 7 },
        ]);
        const undone = applyDocumentOps(applied.document, applied.inverse).unwrap();
        expect(undone.document).toStrictEqual(document);
        expect(applyDocumentOps(undone.document, undone.inverse).unwrap().document).toStrictEqual(
          applied.document,
        );
      }
    }),
    {
      numRuns: 25,
      id: "cross-paragraph hyperlinks preserve formatting and marker identities through exact inverses",
    },
  );
});

test("hyperlink boundaries refuse invalid gaps without changing authored content", () => {
  for (const invalid of [-1, 3, 7, 0.5, Number.NaN]) {
    for (const factory of Object.values(factories)) {
      const document = original();
      const intent = factory({ from: invalid, to: invalid });
      const compiled = compileEditorIntent(document, { intent, mode: { type: "editing" } });
      expect(compiled.isErr()).toBe(true);
      if (compiled.isErr())
        expect(compiled.error.reason).toBe(invalid === 3 ? "splitsSurrogatePair" : "invalidOffset");
      expect(document).toStrictEqual(original());
    }
  }
});

test("direct hyperlink intent targets follow the shared external URL policy atomically", () => {
  assertProperty(
    fc.property(
      fc.constantFrom("custom+folio:", "ftp://", "file:///"),
      fc.constantFrom("example.org/document", "record42"),
      (scheme, target) => {
        for (const factory of Object.values(factories)) {
          const base = factory({ from: 0, to: 1 });
          if (base.type === "removeHyperlink") continue;
          const document = original();
          const compiled = compileEditorIntent(document, {
            intent: { ...base, href: `${scheme}${target}` },
            mode: { type: "editing" },
          });
          expect(compiled.isErr()).toBe(true);
          if (compiled.isErr())
            expect(compiled.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION);
          expect(document).toStrictEqual(original());
        }
      },
    ),
    {
      numRuns: 24,
      id: "direct hyperlink intent targets follow the shared external URL policy atomically",
    },
  );
});

test.each([
  { input: " HTTPS://EXAMPLE.ORG/document ", target: { href: "https://example.org/document" } },
  { input: "mailto:clerk@example.org", target: { href: "mailto:clerk@example.org" } },
  { input: "tel:+420123456789", target: { href: "tel:+420123456789" } },
  { input: "#café東京", target: { anchor: "café東京" } },
  { input: "", target: {} },
])("direct hyperlink intents preserve supported target $input", ({ input, target }) => {
  for (const factory of Object.values(factories)) {
    const base = factory({ from: 0, to: 1 });
    if (base.type === "removeHyperlink") continue;
    const document = original();
    const compiled = compileEditorIntent(document, {
      intent: { ...base, href: input },
      mode: { type: "editing" },
    }).unwrap();
    const applied = applyDocumentOps(document, compiled.ops).unwrap();
    const paragraph = applied.document.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") panic("Hyperlink target fixture lost its paragraph");
    const link = paragraph.content.find((item) => item.type === "hyperlink");
    expect(link).toMatchObject({ type: "hyperlink", ...target });
    expect(applyDocumentOps(applied.document, applied.inverse).unwrap().document).toStrictEqual(
      document,
    );
  }
});
