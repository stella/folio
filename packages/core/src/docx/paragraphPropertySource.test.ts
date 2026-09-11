import { describe, expect, test } from "bun:test";
import { readFile } from "node:fs/promises";

import type { BlockContent, Document, Paragraph } from "../types/document";
import { parseDocx } from "./parser";
import { paragraphPropertySourceFingerprintFromParts } from "./paragraphPropertyDescriptor";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import {
  assignAbsentParagraphPropertySource,
  assignDocumentParagraphPropertySourceContract,
  assignParagraphPropertySource,
  cloneDocumentWithParagraphPropertySources,
  cloneParagraphWithPropertySource,
  cloneParagraphWithoutPropertySource,
  createParagraphPropertyTemplateStore,
  deriveDocumentWithParagraphPropertySources,
  getDocumentParagraphPropertySourceContract,
  getParagraphPropertySource,
  getParagraphPropertySourceToken,
  ParagraphPropertyStorySource,
  visitDocumentStoryParagraphs,
} from "./paragraphPropertySource";
import {
  ParagraphPropertySourceContract,
  ParagraphPropertySourceToken,
  ParagraphPropertyTransientTemplateHandle,
  ParagraphPropertyTransientTemplateStore,
  type ParagraphPropertySourceStory,
} from "./paragraphPropertySourceIdentity";

const LAYOUT_FIXTURE = new URL(
  "../../../../parity/fixtures/layout-kitchen-sink.docx",
  import.meta.url,
);
const BLOCK_SDT_FIXTURE = new URL(
  "./__tests__/__fixtures__/corpus/nested-block-sdt.docx",
  import.meta.url,
);
const LAYOUT_DIGEST = "96aef5ba6161127dc9b85ec487101c934ca1bc93f63fbb8caa53dde196a5fd5b";
const BLOCK_SDT_DIGEST = "2c4d3273683c5a7a059711c5d2a6ba2d040e914cb6b84d72fa809b7419b42e11";

const token = (digest: string, story: ParagraphPropertySourceStory, ordinal: number) =>
  ParagraphPropertySourceContract.fromDigest(digest)
    .bindStoryCensus(story, ordinal + 1)
    .tokenAt(ordinal).serialized;

const tokensIn = (content: BlockContent[]): string[] => {
  const tokens: string[] = [];
  visitDocumentStoryParagraphs(content, (paragraph) => {
    const sourceToken = getParagraphPropertySourceToken(paragraph);
    if (sourceToken) {
      tokens.push(sourceToken.serialized);
    }
  });
  return tokens;
};

const firstParagraphIn = (content: BlockContent[]): Paragraph | undefined => {
  for (const block of content) {
    if (block.type === "paragraph") {
      return block;
    }
    if (block.type === "blockSdt") {
      const paragraph = firstParagraphIn(block.content);
      if (paragraph) {
        return paragraph;
      }
      continue;
    }
    for (const row of block.rows) {
      for (const cell of row.cells) {
        const paragraph = firstParagraphIn(cell.content);
        if (paragraph) {
          return paragraph;
        }
      }
    }
  }
  return undefined;
};

describe("paragraph-property source identity", () => {
  test("only canonical serialized values cross into frozen trusted identities", () => {
    const digest = "a".repeat(64);
    const contract = ParagraphPropertySourceContract.fromDigest(digest);
    const story = { relationshipId: "rId:header/1", type: "header" } as const;
    const sourceToken = contract.bindStoryCensus(story, 13).tokenAt(12);

    expect(Object.isFrozen(contract)).toBe(true);
    expect(Object.isFrozen(sourceToken)).toBe(true);
    expect(ParagraphPropertySourceContract.read(contract.serialized)).toEqual({
      status: "valid",
      value: expect.any(ParagraphPropertySourceContract),
    });
    expect(ParagraphPropertySourceContract.read(`folio-ppr-v2:${"A".repeat(64)}`)).toEqual({
      raw: `folio-ppr-v2:${"A".repeat(64)}`,
      status: "invalid",
    });
    expect(contract.readToken(sourceToken.serialized)).toEqual({
      status: "valid",
      value: expect.any(ParagraphPropertySourceToken),
    });
    expect(contract.readToken("p2s:bad:0")).toEqual({
      raw: "p2s:bad:0",
      status: "invalid",
    });
    expect(sourceToken.belongsTo(contract)).toBe(true);
    expect(sourceToken.belongsToStory({ relationshipId: "rId:header/1", type: "header" })).toBe(
      true,
    );
    expect(
      contract.readToken("p2s:header:%72Id%3Aheader%2F1:c").status,
    ).toBe("invalid");
    expect(sourceToken.belongsToStory({ relationshipId: "other", type: "header" })).toBe(false);
    expect(() =>
      contract.bindStoryCensus({ relationshipId: "", type: "header" }, 1),
    ).toThrow();
    expect(() =>
      contract.bindStoryCensus({ noteId: Number.NaN, type: "footnote" }, 1),
    ).toThrow();
    const collidingPrefixContract = ParagraphPropertySourceContract.fromDigest(
      `${"a".repeat(32)}${"b".repeat(32)}`,
    );
    expect(sourceToken.belongsTo(collidingPrefixContract)).toBe(false);
    const sameLocalToken = collidingPrefixContract.bindStoryCensus(story, 13).tokenAt(12);
    expect(sameLocalToken.serialized).toBe(sourceToken.serialized);
    expect(sourceToken.belongsTo(collidingPrefixContract)).toBe(false);
    expect(sameLocalToken.belongsTo(contract)).toBe(false);
  });

  test("contract-local token construction and parsing agree over bounded story ordinals", () => {
    const contract = ParagraphPropertySourceContract.fromDigest("b".repeat(64));
    for (const ordinal of [0, 1, 35, 36, 1024]) {
      const sourceToken = contract.bindStoryCensus({ type: "document" }, ordinal + 1).tokenAt(ordinal);
      const parsed = contract.readToken(sourceToken.serialized);
      expect(parsed.status).toBe("valid");
      if (parsed.status === "valid") {
        expect(parsed.value.ordinal).toBe(ordinal);
        expect(parsed.value.belongsTo(contract)).toBe(true);
      }
    }
  });

  test("the private contract follows immutable derivations without entering JSON", async () => {
    const document = await parseDocx(await readFile(LAYOUT_FIXTURE), { preloadFonts: false });
    const contract = getDocumentParagraphPropertySourceContract(document);
    const weakened = { ...document, package: { ...document.package } };
    const derived = deriveDocumentWithParagraphPropertySources(document, {
      package: { ...document.package },
    });

    expect(() => getDocumentParagraphPropertySourceContract(weakened)).toThrow();
    expect(getDocumentParagraphPropertySourceContract(derived)).toBe(contract);
    expect(JSON.stringify(derived)).not.toContain("paragraphPropertySourceContract");
    expect(JSON.stringify(derived)).not.toContain("folio-ppr-v2");
  });

  test("the sanctioned structured clone explicitly transfers private source identity", async () => {
    const document = await parseDocx(await readFile(LAYOUT_FIXTURE), { preloadFonts: false });
    const rawClone = structuredClone(document);
    const ownedClone = cloneDocumentWithParagraphPropertySources(document);
    const sourceParagraph = firstParagraphIn(document.package.document.content);
    const rawClonedParagraph = firstParagraphIn(rawClone.package.document.content);
    const clonedParagraph = firstParagraphIn(ownedClone.package.document.content);

    expect(getDocumentParagraphPropertySourceContract(rawClone)).toBeUndefined();
    expect(getDocumentParagraphPropertySourceContract(ownedClone)).toBe(
      getDocumentParagraphPropertySourceContract(document),
    );
    expect(rawClonedParagraph && getParagraphPropertySource(rawClonedParagraph)).toBeUndefined();
    expect(rawClonedParagraph && getParagraphPropertySourceToken(rawClonedParagraph)).toBeUndefined();
    expect(sourceParagraph && getParagraphPropertySource(sourceParagraph)).toEqual(
      clonedParagraph && getParagraphPropertySource(clonedParagraph),
    );
  });

  test("the sanctioned clone retains comment captures and does not retoken derived sections", () => {
    const bodyParagraph: Paragraph = { content: [], type: "paragraph" };
    const commentParagraph: Paragraph = { content: [], type: "paragraph" };
    assignAbsentParagraphPropertySource(bodyParagraph);
    assignParagraphPropertySource(commentParagraph, {
      fingerprint: paragraphPropertySourceFingerprintFromParts({}, {}),
      xml: '<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:test"><x:owner/></w:pPr>',
    });
    const document: Document = {
      package: {
        document: {
          comments: [{ author: "Reviewer", content: [commentParagraph], id: 7 }],
          content: [bodyParagraph],
          sections: [{ content: [bodyParagraph], properties: {} }],
        },
      },
    };
    assignDocumentParagraphPropertySourceContract(document, "e".repeat(64));
    const source = getParagraphPropertySource(commentParagraph);
    const cloned = cloneDocumentWithParagraphPropertySources(document);
    const clonedComment = cloned.package.document.comments?.at(0)?.content.at(0);
    if (!clonedComment) {
      throw new Error("Cloned comment fixture lost its comment paragraph");
    }

    expect(getParagraphPropertySource(clonedComment)).toEqual(source);
    expect(getParagraphPropertySourceToken(clonedComment)?.story).toEqual({
      commentId: 7,
      type: "comment",
    });
    expect(serializeParagraph(clonedComment)).toBe(serializeParagraph(commentParagraph));

    const clonedBodyParagraph = firstParagraphIn(cloned.package.document.content);
    const sectionParagraph = cloned.package.document.sections?.at(0)?.content.at(0);
    if (sectionParagraph?.type !== "paragraph" || clonedBodyParagraph !== sectionParagraph) {
      throw new Error("Derived section must retain the body paragraph alias");
    }
    expect(getParagraphPropertySourceToken(sectionParagraph)).toBe(
      clonedBodyParagraph && getParagraphPropertySourceToken(clonedBodyParagraph),
    );
  });

  test("paragraph source bindings follow spreads, stay immutable, and never enter JSON", async () => {
    const document = await parseDocx(await readFile(LAYOUT_FIXTURE), { preloadFonts: false });
    let paragraph: Paragraph | undefined;
    visitDocumentStoryParagraphs(document.package.document.content, (candidate) => {
      if (!paragraph && getParagraphPropertySource(candidate)) {
        paragraph = candidate;
      }
    });
    if (!paragraph) {
      throw new Error("Layout fixture must contain a paragraph");
    }
    const source = getParagraphPropertySource(paragraph);
    const weakened = { ...paragraph };
    const derived = cloneParagraphWithPropertySource(paragraph, {});

    expect(() => getParagraphPropertySource(weakened)).toThrow();
    expect(getParagraphPropertySource(derived)).toBe(source);
    expect(getParagraphPropertySourceToken(derived)).toBe(
      getParagraphPropertySourceToken(paragraph),
    );
    expect(Object.isFrozen(source)).toBe(true);
    expect(source && Object.isFrozen(source.fingerprint)).toBe(true);
    expect(source && Object.isFrozen(source.fingerprint.pPrBase)).toBe(true);
    expect(JSON.stringify(derived)).not.toContain("p2s:");
    expect(JSON.stringify(derived)).not.toContain("<w:pPr");

    const detached = cloneParagraphWithoutPropertySource(paragraph, {});
    expect(getParagraphPropertySource(detached)).toBeUndefined();
    expect(getParagraphPropertySourceToken(detached)).toBeUndefined();
  });

  test("transient template captures are one-use and conversion scoped", async () => {
    const document = await parseDocx(await readFile(LAYOUT_FIXTURE), { preloadFonts: false });
    const source = firstParagraphIn(document.package.document.content);
    if (!source) {
      throw new Error("Layout fixture must contain a paragraph");
    }
    const store = createParagraphPropertyTemplateStore();
    const sourceStory = ParagraphPropertyStorySource.fromDocument(document, { type: "document" });
    const sourceToken = getParagraphPropertySourceToken(source);
    if (!sourceToken) {
      throw new Error("Layout fixture paragraph must carry its source token");
    }
    const capture = sourceStory.templateCapture(sourceToken);
    const handle = store.registerAll([capture]).at(0);
    if (!handle) {
      throw new Error("Template capture registration must issue one handle");
    }
    const target: Paragraph = { content: [], type: "paragraph" };

    expect(() => store.beginResolution([handle, handle])).toThrow();
    const registry = store.beginResolution([handle]);
    registry.consume(handle, target);
    registry.assertFullyConsumed();
    expect(getParagraphPropertySource(target)).toBe(getParagraphPropertySource(source));
    expect(getParagraphPropertySourceToken(target)).toBeUndefined();
    expect(() => registry.consume(handle, { content: [], type: "paragraph" })).toThrow();
    const repeatedResolution = store.beginResolution([handle]);
    repeatedResolution.consume(handle, { content: [], type: "paragraph" });
    repeatedResolution.assertFullyConsumed();

    const otherStore = createParagraphPropertyTemplateStore();
    expect(() => otherStore.beginResolution([handle])).toThrow();

    const unconsumedHandle = store.registerAll([capture]).at(0);
    if (!unconsumedHandle) {
      throw new Error("Template capture registration must issue one handle");
    }
    const unconsumedRegistry = store.beginResolution([unconsumedHandle]);
    expect(() => unconsumedRegistry.assertFullyConsumed()).toThrow();
    expect(handle).toBeInstanceOf(ParagraphPropertyTransientTemplateHandle);
    expect(handle).not.toBe(otherStore.registerAll([capture]).at(0));
  });

  test("template capacity preflight is atomic", () => {
    const store = new ParagraphPropertyTransientTemplateStore(1);

    expect(() => store.registerAll(["first", "second"])).toThrow(
      "template capture capacity was exceeded",
    );
    const handle = store.registerAll(["only"]).at(0);
    if (!handle) {
      throw new Error("Atomic registration must leave capacity untouched after refusal");
    }
    const registry = store.beginResolution([handle]);
    expect(registry.consume(handle)).toBe("only");
    registry.assertFullyConsumed();
  });

  test("v2 traversal fixes body, text-box, and table ordinals across parse options", async () => {
    const source = await readFile(LAYOUT_FIXTURE);
    const browser = await parseDocx(source, { preloadFonts: true });
    const materializer = await parseDocx(source, { preloadFonts: false });
    const browserTokens = tokensIn(browser.package.document.content);
    const materializerTokens = tokensIn(materializer.package.document.content);

    expect(getDocumentParagraphPropertySourceContract(browser)?.serialized).toBe(
      `folio-ppr-v2:${LAYOUT_DIGEST}`,
    );
    expect(browserTokens).toEqual(
      Array.from({ length: 41 }, (_, ordinal) =>
        token(LAYOUT_DIGEST, { type: "document" }, ordinal),
      ),
    );
    expect(materializerTokens).toEqual(browserTokens);

    const bodyParagraph = browser.package.document.content.at(0);
    if (bodyParagraph?.type !== "paragraph") {
      throw new Error("Layout fixture must start with a paragraph");
    }
    const textBoxParagraph = bodyParagraph.content
      .filter((content) => content.type === "run")
      .flatMap((run) => run.content)
      .find((content) => content.type === "shape" && content.shape.textBody)
      ?.shape.textBody?.content.at(0);
    const table = browser.package.document.content.find((block) => block.type === "table");
    const tableParagraph = table?.rows.at(0)?.cells.at(0)?.content.at(0);

    expect(getParagraphPropertySourceToken(bodyParagraph)?.serialized).toBe(
      token(LAYOUT_DIGEST, { type: "document" }, 0),
    );
    expect(textBoxParagraph?.type).toBe("paragraph");
    expect(
      textBoxParagraph?.type === "paragraph"
        ? getParagraphPropertySourceToken(textBoxParagraph)?.serialized
        : undefined,
    ).toBe(token(LAYOUT_DIGEST, { type: "document" }, 1));
    expect(tableParagraph?.type).toBe("paragraph");
    expect(
      tableParagraph?.type === "paragraph"
        ? getParagraphPropertySourceToken(tableParagraph)?.serialized
        : undefined,
    ).toBe(token(LAYOUT_DIGEST, { type: "document" }, 7));
  });

  test("v2 traversal includes nested block content controls", async () => {
    const source = await readFile(BLOCK_SDT_FIXTURE);
    const document = await parseDocx(source, { preloadFonts: false });
    const paragraph = firstParagraphIn(document.package.document.content);

    expect(tokensIn(document.package.document.content)).toEqual([
      token(BLOCK_SDT_DIGEST, { type: "document" }, 0),
    ]);
    expect(paragraph && getParagraphPropertySourceToken(paragraph)?.serialized).toBe(
      token(BLOCK_SDT_DIGEST, { type: "document" }, 0),
    );
  });

  test("v2 all-story ordinals ignore package map and note array order", () => {
    const paragraph = (text: string): Paragraph => ({
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text }] }],
    });
    const makeDocument = (reverse: boolean) => {
      const body = paragraph("body");
      const headerA = paragraph("header-a");
      const headerB = paragraph("header-b");
      const footerA = paragraph("footer-a");
      const footerB = paragraph("footer-b");
      const footnoteA = paragraph("footnote-a");
      const footnoteB = paragraph("footnote-b");
      const endnoteA = paragraph("endnote-a");
      const commentA = paragraph("comment-a");
      const commentB = paragraph("comment-b");
      const ordered = reverse
        ? [
            body,
            headerB,
            headerA,
            footerB,
            footerA,
            footnoteB,
            footnoteA,
            endnoteA,
            commentB,
            commentA,
          ]
        : [
            body,
            headerA,
            headerB,
            footerA,
            footerB,
            footnoteA,
            footnoteB,
            endnoteA,
            commentA,
            commentB,
          ];
      for (const source of ordered) {
        assignAbsentParagraphPropertySource(source);
      }
      const document: Document = {
        package: {
          document: {
            comments: reverse
              ? [
                  { author: "B", content: [commentB], id: 2 },
                  { author: "A", content: [commentA], id: 1 },
                ]
              : [
                  { author: "A", content: [commentA], id: 1 },
                  { author: "B", content: [commentB], id: 2 },
                ],
            content: [body],
            sections: [{ content: [body], properties: {} }],
          },
          headers: new Map(
            reverse
              ? [
                  ["rIdB", { type: "header", hdrFtrType: "default", content: [headerB] }],
                  ["rIdA", { type: "header", hdrFtrType: "default", content: [headerA] }],
                ]
              : [
                  ["rIdA", { type: "header", hdrFtrType: "default", content: [headerA] }],
                  ["rIdB", { type: "header", hdrFtrType: "default", content: [headerB] }],
                ],
          ),
          footers: new Map(
            reverse
              ? [
                  ["rIdD", { type: "footer", hdrFtrType: "default", content: [footerB] }],
                  ["rIdC", { type: "footer", hdrFtrType: "default", content: [footerA] }],
                ]
              : [
                  ["rIdC", { type: "footer", hdrFtrType: "default", content: [footerA] }],
                  ["rIdD", { type: "footer", hdrFtrType: "default", content: [footerB] }],
                ],
          ),
          footnotes: reverse
            ? [
                { type: "footnote", id: 2, content: [footnoteB] },
                { type: "footnote", id: 1, content: [footnoteA] },
              ]
            : [
                { type: "footnote", id: 1, content: [footnoteA] },
                { type: "footnote", id: 2, content: [footnoteB] },
              ],
          endnotes: [{ type: "endnote", id: 1, content: [endnoteA] }],
        },
      };
      assignDocumentParagraphPropertySourceContract(document, "c".repeat(64));
      return [
        body,
        headerA,
        headerB,
        footerA,
        footerB,
        footnoteA,
        footnoteB,
        endnoteA,
        commentA,
        commentB,
      ].map((source) => getParagraphPropertySourceToken(source)?.serialized);
    };

    expect(makeDocument(true)).toEqual(makeDocument(false));
    expect(makeDocument(false)).toEqual(
      [
        token("c".repeat(64), { type: "document" }, 0),
        token("c".repeat(64), { relationshipId: "rIdA", type: "header" }, 0),
        token("c".repeat(64), { relationshipId: "rIdB", type: "header" }, 0),
        token("c".repeat(64), { relationshipId: "rIdC", type: "footer" }, 0),
        token("c".repeat(64), { relationshipId: "rIdD", type: "footer" }, 0),
        token("c".repeat(64), { noteId: 1, type: "footnote" }, 0),
        token("c".repeat(64), { noteId: 2, type: "footnote" }, 0),
        token("c".repeat(64), { noteId: 1, type: "endnote" }, 0),
        token("c".repeat(64), { commentId: 1, type: "comment" }, 0),
        token("c".repeat(64), { commentId: 2, type: "comment" }, 0),
      ],
    );
  });

  test("contract assignment validates the whole story census before mutating provenance", () => {
    const body: Paragraph = { content: [], type: "paragraph" };
    const lateHeader: Paragraph = { content: [], type: "paragraph" };
    assignAbsentParagraphPropertySource(body);
    const document: Document = {
      package: {
        document: { content: [body] },
        headers: new Map([
          ["z", { content: [lateHeader], hdrFtrType: "default", type: "header" }],
        ]),
      },
    };

    expect(() => assignDocumentParagraphPropertySourceContract(document, "f".repeat(64))).toThrow();
    expect(getDocumentParagraphPropertySourceContract(document)).toBeUndefined();
    expect(getParagraphPropertySourceToken(body)).toBeUndefined();

    assignAbsentParagraphPropertySource(lateHeader);
    assignDocumentParagraphPropertySourceContract(document, "f".repeat(64));
    expect(getParagraphPropertySourceToken(body)?.belongsToStory({ type: "document" })).toBe(true);
    expect(
      getParagraphPropertySourceToken(lateHeader)?.belongsToStory({
        relationshipId: "z",
        type: "header",
      }),
    ).toBe(true);
  });

  test("a late non-transitionable capture cannot partially bind earlier stories", () => {
    const body: Paragraph = { content: [], type: "paragraph" };
    const lateHeader: Paragraph = { content: [], type: "paragraph" };
    assignAbsentParagraphPropertySource(body);
    assignAbsentParagraphPropertySource(lateHeader);
    Object.freeze(lateHeader);
    const document: Document = {
      package: {
        document: { content: [body] },
        headers: new Map([
          ["z", { content: [lateHeader], hdrFtrType: "default", type: "header" }],
        ]),
      },
    };

    expect(() => assignDocumentParagraphPropertySourceContract(document, "e".repeat(64))).toThrow();
    expect(getDocumentParagraphPropertySourceContract(document)).toBeUndefined();
    expect(getParagraphPropertySourceToken(body)).toBeUndefined();
    expect(() => getParagraphPropertySourceToken(lateHeader)).toThrow(
      "weakened paragraph-property source binding",
    );
  });

  test("derived sections cannot introduce a second paragraph ownership surface", () => {
    const body: Paragraph = { content: [], type: "paragraph" };
    const detachedSectionParagraph: Paragraph = { content: [], type: "paragraph" };
    assignAbsentParagraphPropertySource(body);
    assignAbsentParagraphPropertySource(detachedSectionParagraph);
    const document: Document = {
      package: {
        document: {
          content: [body],
          sections: [{ content: [detachedSectionParagraph], properties: {} }],
        },
      },
    };

    expect(() => assignDocumentParagraphPropertySourceContract(document, "d".repeat(64))).toThrow();
    expect(getDocumentParagraphPropertySourceContract(document)).toBeUndefined();
    expect(getParagraphPropertySourceToken(body)).toBeUndefined();
  });

  test.each(["comment", "footnote"] as const)(
    "duplicate %s story identities fail before any paragraph is bound",
    (storyType) => {
      const first: Paragraph = { content: [], type: "paragraph" };
      const second: Paragraph = { content: [], type: "paragraph" };
      assignAbsentParagraphPropertySource(first);
      assignAbsentParagraphPropertySource(second);
      const document: Document = {
        package: {
          document: {
            ...(storyType === "comment"
              ? {
                  comments: [
                    { author: "A", content: [first], id: 4 },
                    { author: "B", content: [second], id: 4 },
                  ],
                }
              : {}),
            content: [],
          },
          ...(storyType === "footnote"
            ? {
                footnotes: [
                  { content: [first], id: 4, type: "footnote" },
                  { content: [second], id: 4, type: "footnote" },
                ],
              }
            : {}),
        },
      };

      expect(() =>
        assignDocumentParagraphPropertySourceContract(document, "9".repeat(64)),
      ).toThrow();
      expect(getDocumentParagraphPropertySourceContract(document)).toBeUndefined();
      expect(getParagraphPropertySourceToken(first)).toBeUndefined();
      expect(getParagraphPropertySourceToken(second)).toBeUndefined();
    },
  );
});
