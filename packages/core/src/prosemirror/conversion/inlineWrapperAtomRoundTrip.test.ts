/**
 * A transparent wrapper belongs around every inline item it contains.
 *
 * ProseMirror stores that containment as an `inlineWrapper` mark. An inline
 * leaf's own mark set describes content inside the leaf, so consulting it to
 * decide whether the leaf may carry a mark drops wrappers around atomic fields
 * and equations even though their paragraph admits the mark.
 */

import { describe, expect, test } from "bun:test";

import { INLINE_WRAPPER_ELEMENTS } from "../../docx/inlineWrapperParser";
import { serializeParagraph } from "../../docx/serializer/paragraphSerializer";
import type {
  Document,
  InlineWrapper,
  MathEquation,
  Paragraph,
  SimpleField,
} from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { schema } from "../schema";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

type AtomicWrapperContent = MathEquation | SimpleField;

const WRAPPERS = {
  bdo: (content: AtomicWrapperContent): InlineWrapper => ({
    type: "inlineWrapper",
    kind: "bidi",
    control: "override",
    direction: "rtl",
    content: [content],
  }),
  dir: (content: AtomicWrapperContent): InlineWrapper => ({
    type: "inlineWrapper",
    kind: "bidi",
    control: "embedding",
    direction: "ltr",
    content: [content],
  }),
  smartTag: (content: AtomicWrapperContent): InlineWrapper => ({
    type: "inlineWrapper",
    kind: "smartTag",
    element: "place",
    content: [content],
  }),
  customXml: (content: AtomicWrapperContent): InlineWrapper => ({
    type: "inlineWrapper",
    kind: "customXml",
    element: "party",
    content: [content],
  }),
} as const satisfies Record<
  (typeof INLINE_WRAPPER_ELEMENTS)[number],
  (content: AtomicWrapperContent) => InlineWrapper
>;

const ATOMS = {
  fldSimple: (): SimpleField => ({
    type: "simpleField",
    instruction: " PAGE ",
    fieldType: "PAGE",
    content: [{ type: "run", content: [{ type: "text", text: "1" }] }],
  }),
  oMath: (): MathEquation => ({
    type: "mathEquation",
    display: "inline",
    ommlXml: "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>",
    plainText: "x",
  }),
  oMathPara: (): MathEquation => ({
    type: "mathEquation",
    display: "block",
    ommlXml: "<m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>",
    plainText: "x",
  }),
} as const;

const documentWith = (content: InlineWrapper): Document => {
  const document = createEmptyDocument();
  document.package.document.content = [{ type: "paragraph", content: [content] }];
  return document;
};

const firstParagraph = (document: Document): Paragraph => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    throw new Error("Expected an atomic-wrapper paragraph");
  }
  return paragraph;
};

const wrapperBounds = {
  bdo: { open: "<w:bdo", close: "</w:bdo>" },
  dir: { open: "<w:dir", close: "</w:dir>" },
  smartTag: { open: "<w:smartTag", close: "</w:smartTag>" },
  customXml: { open: "<w:customXml", close: "</w:customXml>" },
} as const satisfies Record<
  (typeof INLINE_WRAPPER_ELEMENTS)[number],
  { open: string; close: string }
>;

describe("transparent wrappers around atomic inline content", () => {
  test("the fixture covers every transparent wrapper the parser declares", () => {
    expect(Object.keys(WRAPPERS).toSorted()).toEqual([...INLINE_WRAPPER_ELEMENTS].toSorted());
    expect(Object.keys(wrapperBounds).toSorted()).toEqual([...INLINE_WRAPPER_ELEMENTS].toSorted());
  });

  test("inline leaves carry their parent's wrapper mark", () => {
    const wrapperMark = schema.marks["inlineWrapper"];
    const paragraphType = schema.nodes["paragraph"];
    if (!wrapperMark || !paragraphType) {
      throw new Error("Expected the paragraph and inline-wrapper schema members");
    }
    expect(paragraphType.allowsMarkType(wrapperMark)).toBe(true);

    for (const atom of Object.values(ATOMS)) {
      const source = documentWith(WRAPPERS.bdo(atom()));
      const projectedParagraph = toProseDoc(source).firstChild;
      const leaf = projectedParagraph?.firstChild;
      if (!leaf) {
        throw new Error("Expected the wrapper projection to keep its atomic leaf");
      }
      expect(leaf.isInline).toBe(true);
      expect(leaf.isLeaf).toBe(true);
      // The fixture reaches the old fault: this answers about child content,
      // not about a mark that the paragraph places on the leaf itself.
      expect(leaf.type.allowsMarkType(wrapperMark)).toBe(false);
      expect(leaf.marks.some(({ type }) => type === wrapperMark)).toBe(true);
    }
  });

  for (const element of INLINE_WRAPPER_ELEMENTS) {
    for (const [child, atom] of Object.entries(ATOMS)) {
      test(`keeps w:${element} around ${child} through editor projection and save`, () => {
        const source = documentWith(WRAPPERS[element](atom()));
        const roundTripped = fromProseDoc(toProseDoc(source), source);
        const paragraph = firstParagraph(roundTripped);
        const wrapper = paragraph.content.at(0);

        expect(wrapper?.type).toBe("inlineWrapper");
        if (wrapper?.type !== "inlineWrapper") {
          throw new Error("Expected the atomic content to remain in its wrapper");
        }
        expect(wrapper.content.at(0)?.type).toBe(atom().type);

        const xml = serializeParagraph(paragraph);
        const { open, close } = wrapperBounds[element];
        const childTag = `<${child === "fldSimple" ? "w" : "m"}:${child}`;
        const openAt = xml.indexOf(open);
        const childAt = xml.indexOf(childTag);
        const closeAt = xml.indexOf(close);
        expect(openAt).toBeGreaterThanOrEqual(0);
        expect(childAt).toBeGreaterThan(openAt);
        expect(closeAt).toBeGreaterThan(childAt);
      });
    }
  }
});
