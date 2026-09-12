import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";

import { propertyConfig } from "../../../../../test/property-testing";
import type { Document, Paragraph, Run, Table } from "../../types/document";
import { schema } from "../schema";
import { stableProjectionIdentity } from "./__tests__/stableProjectionIdentity";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const generatedText = fc
  .array(fc.constantFrom("alpha", " ", "beta", "\n", "§", "😀", "č", "م", "e\u0301"), {
    maxLength: 8,
  })
  .map((parts) => parts.join(""));

const generatedNonEmptyText = fc
  .array(fc.constantFrom("alpha", " ", "beta", "§", "😀", "č", "م", "e\u0301"), {
    minLength: 1,
    maxLength: 8,
  })
  .map((parts) => parts.join(""));

const generatedParagraph = fc
  .record({
    text: generatedText,
    alignment: fc.option(fc.constantFrom("center", "right"), {
      nil: undefined,
    }),
  })
  .map(({ text, alignment }) => {
    const content = (
      text.length === 0 ? [] : [{ type: "run", content: [{ type: "text", text }] }]
    ) satisfies Paragraph["content"];
    if (alignment === undefined) {
      return { type: "paragraph", content } satisfies Paragraph;
    }
    return {
      type: "paragraph",
      content,
      formatting: { alignment },
    } satisfies Paragraph;
  });

const emptyTextBody = [] satisfies readonly Paragraph[];

const generatedTextBody = fc.oneof(
  fc.constant(emptyTextBody),
  fc.array(generatedParagraph, { minLength: 1, maxLength: 4 }),
);

type PlaceholderEdit =
  | { readonly type: "identity"; readonly paraId: string; readonly textId: string }
  | { readonly type: "text"; readonly text: string }
  | { readonly type: "formatting"; readonly alignment: "center" | "right" };

const generatedId = fc
  .integer({ min: 0, max: 0x7fff_ffff })
  .map((value) => value.toString(16).toUpperCase().padStart(8, "0"));

const generatedPlaceholderEdit = fc.oneof(
  fc.record({
    type: fc.constant("identity"),
    paraId: generatedId,
    textId: generatedId,
  }),
  fc.record({ type: fc.constant("text"), text: generatedNonEmptyText }),
  fc.record({
    type: fc.constant("formatting"),
    alignment: fc.constantFrom("center", "right"),
  }),
) satisfies fc.Arbitrary<PlaceholderEdit>;

const generatedIdentityClasses = fc.array(fc.integer({ min: 0, max: 4 }), {
  minLength: 1,
  maxLength: 10,
});

const generatedRepeatedIdentityClasses = generatedIdentityClasses.filter(
  (classes) => new Set(classes).size < classes.length,
);

describe("text-box content provenance properties", () => {
  test("document-model projection preserves empty and authored bodies to a fixed point", () => {
    fc.assert(
      fc.property(generatedTextBody, (body) => {
        const source = documentWithTextBody(body);
        const projected = toProseDoc(source);
        const restored = fromProseDoc(projected, source);
        const reprojected = toProseDoc(restored);
        const restoredAgain = fromProseDoc(reprojected, restored);

        expect(textBodyContent(restored)).toEqual(body);
        expect(textBodyContent(restoredAgain)).toEqual(body);
        expect(textBoxContentState(projected)).toEqual({
          type: body.length === 0 ? "source-empty" : "authored",
        });
        expect(stableProjectionIdentity(reprojected)).toBe(stableProjectionIdentity(projected));
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("only identity allocation may leave a source-empty placeholder source-empty", () => {
    fc.assert(
      fc.property(generatedPlaceholderEdit, (edit) => {
        const source = documentWithTextBody([]);
        const projected = toProseDoc(source);
        const textBox = projected.firstChild;
        const placeholder = textBox?.firstChild;
        if (textBox?.type.name !== "textBox" || placeholder?.type.name !== "paragraph") {
          throw new Error("Expected a source-empty text-box placeholder");
        }

        const editedParagraph = (() => {
          switch (edit.type) {
            case "identity":
              return placeholder.type.create(
                { ...placeholder.attrs, paraId: edit.paraId, textId: edit.textId },
                placeholder.content,
              );
            case "text":
              return placeholder.type.create(placeholder.attrs, [schema.text(edit.text)]);
            case "formatting":
              return placeholder.type.create(
                { ...placeholder.attrs, alignment: edit.alignment },
                placeholder.content,
              );
            default: {
              const exhaustive: never = edit;
              return exhaustive;
            }
          }
        })();
        const edited = schema.node("doc", projected.attrs, [
          textBox.type.create(textBox.attrs, [editedParagraph]),
        ]);
        const restored = fromProseDoc(edited, source);
        const reprojected = toProseDoc(restored);
        const body = textBodyContent(restored);

        switch (edit.type) {
          case "identity":
            expect(body).toEqual([]);
            expect(textBoxContentState(reprojected)).toEqual({ type: "source-empty" });
            break;
          case "text":
            expect(paragraphText(body.at(0))).toBe(edit.text);
            expect(textBoxContentState(reprojected)).toEqual({ type: "authored" });
            break;
          case "formatting":
            expect(body).toMatchObject([
              { type: "paragraph", formatting: { alignment: edit.alignment } },
            ]);
            expect(textBoxContentState(reprojected)).toEqual({ type: "authored" });
            break;
          default: {
            const exhaustive: never = edit;
            throw new Error(`Unexpected edit: ${String(exhaustive)}`);
          }
        }

        expect(stableProjectionIdentity(toProseDoc(fromProseDoc(reprojected, restored)))).toBe(
          stableProjectionIdentity(reprojected),
        );
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("canonical identity preserves arbitrary equality relationships", () => {
    fc.assert(
      fc.property(generatedIdentityClasses, (classes) => {
        const first = projectionWithIdentityClasses(classes, "first");
        const second = projectionWithIdentityClasses(classes, "second");

        expect(stableProjectionIdentity(first)).toBe(stableProjectionIdentity(second));
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("canonical identity detects a split equality relationship", () => {
    fc.assert(
      fc.property(generatedRepeatedIdentityClasses, (classes) => {
        const repeatedClass = classes.find(
          (candidate, index) => classes.indexOf(candidate) !== index,
        );
        if (repeatedClass === undefined) {
          throw new Error("Expected a repeated identity class");
        }
        const splitIndex = classes.lastIndexOf(repeatedClass);
        const splitClasses = classes.with(splitIndex, Math.max(...classes) + 1);

        expect(stableProjectionIdentity(projectionWithIdentityClasses(classes, "first"))).not.toBe(
          stableProjectionIdentity(projectionWithIdentityClasses(splitClasses, "second")),
        );
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });
});

const documentWithTextBody = (content: readonly Paragraph[]): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              content: [
                {
                  type: "shape",
                  shape: {
                    type: "shape",
                    shapeType: "textBox",
                    size: { width: 914_400, height: 457_200 },
                    textBody: { content: [...content] },
                  },
                },
              ],
            },
          ],
        },
      ],
    },
  },
});

const textBodyContent = (document: Document): readonly (Paragraph | Table)[] => {
  const host = document.package.document.content.at(0);
  if (host?.type !== "paragraph") {
    throw new Error("Expected text-box host paragraph");
  }
  const run = host.content.find((content): content is Run => content.type === "run");
  const shape = run?.content.find((content) => content.type === "shape");
  if (shape?.type !== "shape" || shape.shape.shapeType !== "textBox") {
    throw new Error("Expected text-box shape");
  }
  return shape.shape.textBody?.content ?? [];
};

const paragraphText = (block: Paragraph | Table | undefined): string | undefined => {
  if (block?.type !== "paragraph") {
    return undefined;
  }
  return block.content
    .flatMap((content) => (content.type === "run" ? content.content : []))
    .flatMap((content) => (content.type === "text" ? content.text : []))
    .join("");
};

const textBoxContentState = (document: PMNode): unknown =>
  document.firstChild?.attrs["_docxTextBodyContentState"];

const projectionWithIdentityClasses = (classes: readonly number[], namespace: string): PMNode =>
  schema.node(
    "doc",
    null,
    classes.map((identityClass, index) =>
      schema.node(
        "textBox",
        {
          _docxGroupId: `${namespace}-group-${String(identityClass)}`,
          _docxAnchorId: `${namespace}-anchor-${String(classes.at(-index - 1))}`,
        },
        [schema.node("paragraph")],
      ),
    ),
  );
