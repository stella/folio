import { panic } from "better-result";
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { checkCoverageSample } from "../../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { allocateEditorIntentIds, compileEditorIntent } from "../editorIntent";
import { paragraphLength } from "../offsets";
import { OP_STORIES, type TextPosition } from "../types";
import {
  documentArbitrary,
  reviewDocumentArbitrary,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(240_000));

type Typed = { document: Document; caret: TextPosition; reused: boolean };

/**
 * Type `text` at `caret` as `author`, suggesting. The compiler's own
 * application, when it hands one back, must be what applying its ops gives.
 */
const typeAt = (
  document: Document,
  caret: TextPosition,
  text: string,
  author: string,
  modeType: "suggesting" | "editing",
): Typed | undefined => {
  const intent = { type: "replaceText", from: caret, to: caret, text } as const;
  const ids = allocateEditorIntentIds(document, intent);
  const compiled = compileEditorIntent(document, {
    intent,
    mode:
      modeType === "suggesting"
        ? {
            type: "suggesting",
            revision: {
              id: ids.revisionId,
              author,
              date: "2026-01-01T00:00:00Z",
            },
            newIds: ids.newIds,
          }
        : { type: "editing", newIds: ids.newIds },
  });
  if (compiled.isErr()) return undefined;
  const applied = applyDocumentOps(document, compiled.value.ops);
  const own = compiled.value.applied;
  // A generated caret can fall where no editor caret goes (inside a
  // surrogate pair); its ops then apply nowhere, and no application is
  // handed back.
  if (applied.isErr()) {
    expect(own).toBeUndefined();
    return undefined;
  }
  if (own !== undefined) {
    expect(own.document).toEqual(applied.value.document);
    expect(own.inverse).toEqual(applied.value.inverse);
    expect(own.touched).toEqual(applied.value.touched);
    expect(own.revisions).toEqual(applied.value.revisions);
  }
  return {
    document: applied.value.document,
    caret: compiled.value.selection,
    reused: own !== undefined,
  };
};

describe("a compiled intent's own application", () => {
  test("equals applying its ops, through generated typing runs in both modes", () => {
    const law = (reused: { suggesting: number; editing: number }) =>
      fc.property(
        fc.oneof(documentArbitrary, reviewDocumentArbitrary),
        fc.nat(),
        fc.nat(),
        fc.array(fc.constantFrom("x", "yz", " "), {
          minLength: 2,
          maxLength: 5,
        }),
        fc.constantFrom("Typist", "Author"),
        fc.constantFrom("suggesting", "editing"),
        (document, paragraphPick, offsetPick, keys, author, modeType) => {
          const paragraphs = storyParagraphs(document.package.document).map(
            ({ paragraph }) => paragraph,
          );
          const source = paragraphs.at(paragraphPick % paragraphs.length);
          if (source === undefined) panic("A generated body has a paragraph.");
          let typed: Typed | undefined = {
            document,
            caret: {
              story: OP_STORIES.MAIN,
              blockId: source.paraId ?? "",
              offset: offsetPick % (paragraphLength(source) + 1),
            },
            reused: false,
          };
          for (const key of keys) {
            if (typed === undefined) return;
            typed = typeAt(typed.document, typed.caret, key, author, modeType);
            if (typed?.reused) reused[modeType] += 1;
          }
        },
      );
    assertProperty(law({ suggesting: 0, editing: 0 }), {
      numRuns: 200,
      id: "a compiled intent's own application",
    });
    // The property's run count, the two modes taking turns.
    const COVERAGE_RUNS = 200;
    const reused = { suggesting: 0, editing: 0 };
    checkCoverageSample(law(reused), COVERAGE_RUNS, {
      map: ([document, paragraphPick, offsetPick, keys, author], index) => [
        document,
        paragraphPick,
        offsetPick,
        keys,
        author,
        index % 2 === 0 ? "suggesting" : "editing",
      ],
    });
    // Direct typing, and typing on beside one's own insertion, hand their
    // application back.
    expect(reused.suggesting).toBeGreaterThan(0);
    expect(reused.editing).toBeGreaterThan(0);
  });
});
