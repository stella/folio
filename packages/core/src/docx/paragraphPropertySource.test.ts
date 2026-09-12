import { describe, expect, test } from "bun:test";
import { readFile } from "node:fs/promises";

import type { BlockContent, Paragraph } from "../types/document";
import { parseDocx } from "./parser";
import {
  TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR,
  cloneDocumentWithParagraphPropertySources,
  copyDocumentParagraphPropertySources,
  copyParagraphPropertySource,
  getDocumentParagraphPropertySourceContract,
  getParagraphPropertySource,
  getParagraphPropertySourceToken,
  restoreTableCellsWithParagraphPropertySources,
  transportTableCellsWithParagraphPropertySources,
  visitDocumentStoryParagraphs,
} from "./paragraphPropertySource";

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

const token = (digest: string, ordinal: number) =>
  `p1d:${digest.slice(0, 32)}:${ordinal.toString(36)}`;

const tokensIn = (content: BlockContent[]): string[] => {
  const tokens: string[] = [];
  visitDocumentStoryParagraphs(content, (paragraph) => {
    const sourceToken = getParagraphPropertySourceToken(paragraph);
    if (sourceToken) {
      tokens.push(sourceToken);
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
  test("the private contract follows immutable derivations without entering JSON", async () => {
    const document = await parseDocx(await readFile(LAYOUT_FIXTURE), { preloadFonts: false });
    const contract = getDocumentParagraphPropertySourceContract(document);
    const derived = { ...document, package: { ...document.package } };

    expect(getDocumentParagraphPropertySourceContract(derived)).toBe(contract);
    expect(copyDocumentParagraphPropertySources(derived)?.size).toBe(41);
    expect(JSON.stringify(derived)).not.toContain("paragraphPropertySourceBinding");
    expect(JSON.stringify(derived)).not.toContain("folio-ppr-v1");
  });

  test("callers receive an isolated copy of the private source registry", async () => {
    const document = await parseDocx(await readFile(LAYOUT_FIXTURE), { preloadFonts: false });
    const sources = copyDocumentParagraphPropertySources(document);
    const sourceCount = sources?.size;

    sources?.clear();

    expect(sourceCount).toBe(41);
    expect(copyDocumentParagraphPropertySources(document)?.size).toBe(sourceCount);
  });

  test("the sanctioned structured clone explicitly transfers private source identity", async () => {
    const document = await parseDocx(await readFile(LAYOUT_FIXTURE), { preloadFonts: false });
    const rawClone = structuredClone(document);
    const ownedClone = cloneDocumentWithParagraphPropertySources(document);
    const sourceParagraph = firstParagraphIn(document.package.document.content);
    const clonedParagraph = firstParagraphIn(ownedClone.package.document.content);

    expect(getDocumentParagraphPropertySourceContract(rawClone)).toBeUndefined();
    expect(copyDocumentParagraphPropertySources(rawClone)).toBeUndefined();
    expect(getDocumentParagraphPropertySourceContract(ownedClone)).toBe(
      getDocumentParagraphPropertySourceContract(document),
    );
    expect(copyDocumentParagraphPropertySources(ownedClone)?.size).toBe(41);
    expect(sourceParagraph && getParagraphPropertySource(sourceParagraph)).toEqual(
      clonedParagraph && getParagraphPropertySource(clonedParagraph),
    );
  });

  test("v1 traversal fixes body, text-box, and table ordinals across parse options", async () => {
    const source = await readFile(LAYOUT_FIXTURE);
    const browser = await parseDocx(source, { preloadFonts: true });
    const materializer = await parseDocx(source, { preloadFonts: false });
    const browserTokens = tokensIn(browser.package.document.content);
    const materializerTokens = tokensIn(materializer.package.document.content);

    expect(getDocumentParagraphPropertySourceContract(browser)).toBe(
      `folio-ppr-v1:${LAYOUT_DIGEST}`,
    );
    expect(browserTokens).toEqual(
      Array.from({ length: 41 }, (_, ordinal) => token(LAYOUT_DIGEST, ordinal)),
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

    expect(getParagraphPropertySourceToken(bodyParagraph)).toBe(token(LAYOUT_DIGEST, 0));
    expect(textBoxParagraph?.type).toBe("paragraph");
    expect(
      textBoxParagraph?.type === "paragraph"
        ? getParagraphPropertySourceToken(textBoxParagraph)
        : undefined,
    ).toBe(token(LAYOUT_DIGEST, 1));
    expect(tableParagraph?.type).toBe("paragraph");
    expect(
      tableParagraph?.type === "paragraph"
        ? getParagraphPropertySourceToken(tableParagraph)
        : undefined,
    ).toBe(token(LAYOUT_DIGEST, 7));
  });

  test("v1 traversal includes nested block content controls", async () => {
    const source = await readFile(BLOCK_SDT_FIXTURE);
    const document = await parseDocx(source, { preloadFonts: false });
    const paragraph = firstParagraphIn(document.package.document.content);

    expect(tokensIn(document.package.document.content)).toEqual([token(BLOCK_SDT_DIGEST, 0)]);
    expect(paragraph && getParagraphPropertySourceToken(paragraph)).toBe(
      token(BLOCK_SDT_DIGEST, 0),
    );
  });

  test("hidden table paragraphs retain one source identity across transport clones", async () => {
    const document = await parseDocx(await readFile(LAYOUT_FIXTURE), { preloadFonts: false });
    const paragraph = firstParagraphIn(document.package.document.content);
    if (!paragraph) {
      throw new Error("Layout fixture must contain a paragraph");
    }
    const sourceToken = getParagraphPropertySourceToken(paragraph);
    const firstTransport = transportTableCellsWithParagraphPropertySources([
      { type: "tableCell", content: [paragraph] },
    ]);
    const secondTransport = transportTableCellsWithParagraphPropertySources(firstTransport);
    const restored = restoreTableCellsWithParagraphPropertySources(secondTransport);
    const restoredParagraph = restored.at(0)?.content.at(0);
    if (typeof sourceToken !== "string" || restoredParagraph?.type !== "paragraph") {
      throw new Error("Transport fixture lost its paragraph source");
    }

    expect(getParagraphPropertySourceToken(restoredParagraph)).toBe(sourceToken);
    expect(Object.hasOwn(restoredParagraph, TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR)).toBe(false);
  });

  test("transport explicitly marks a newly authored hidden paragraph", () => {
    const paragraph: Paragraph = { type: "paragraph", content: [] };
    const transported = transportTableCellsWithParagraphPropertySources([
      { type: "tableCell", content: [paragraph] },
    ]);
    const transportedParagraph = transported.at(0)?.content.at(0);
    if (transportedParagraph?.type !== "paragraph") {
      throw new Error("Transport fixture lost its authored paragraph");
    }

    expect(Reflect.get(transportedParagraph, TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR)).toEqual({
      type: "authored",
    });
    const restored = restoreTableCellsWithParagraphPropertySources(transported);
    const restoredParagraph = restored.at(0)?.content.at(0);
    expect(restoredParagraph?.type).toBe("paragraph");
    expect(
      restoredParagraph?.type === "paragraph"
        ? Object.hasOwn(restoredParagraph, TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR)
        : true,
    ).toBe(false);
  });

  test.each([
    { type: "source" },
    { token: "p1d:00000000000000000000000000000000:0", type: "source" },
    { type: "authored" },
  ] as const)("rejects conflicting hidden transport binding %#", async (binding) => {
    const document = await parseDocx(await readFile(LAYOUT_FIXTURE), { preloadFonts: false });
    const paragraph = firstParagraphIn(document.package.document.content);
    if (!paragraph) {
      throw new Error("Layout fixture must contain a paragraph");
    }
    const tampered = { ...paragraph };
    copyParagraphPropertySource(tampered, paragraph);
    if (!Reflect.set(tampered, TABLE_CELL_PARAGRAPH_SOURCE_BINDING_ATTR, binding)) {
      throw new Error("Transport fixture could not attach its invalid binding");
    }

    expect(() =>
      transportTableCellsWithParagraphPropertySources([
        { type: "tableCell", content: [tampered] },
      ]),
    ).toThrow();
  });
});
