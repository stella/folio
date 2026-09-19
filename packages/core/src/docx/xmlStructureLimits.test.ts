/**
 * What a package may allocate, not what it may weigh.
 *
 * The byte bounds these tests sit beside cannot express the exposure they are
 * asked to cover: `<w:r/>` is seven bytes of markup and roughly fifty bytes of
 * parsed tree, so markup that satisfies every byte bound still buys an
 * arbitrary tree. Each case here is a shape, not a file: element-dense markup,
 * attribute-dense markup, a package spread thin across parts, repetitive
 * nesting, and a package that inflates. Every one asserts the refusal is
 * typed, names the part, and lands with the scan stopped at the bound rather
 * than at the end of the input.
 */
import { describe, expect, mock, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "./parser";
import { DocxSecurityError, unzipDocx } from "./unzip";
import { loadDocxArchive } from "./server/boundedArchive";
import {
  assertXmlResourceLimits,
  createXmlPackageBudget,
  FOLIO_XML_RESOURCE_LIMITS,
  XmlResourceLimitError,
} from "./xmlResourceLimits";
import * as xmlParser from "./xmlParser";

const CONTENT_TYPES = "<Types />";

type Part = { path: string; xml: string };

const packageOf = async (parts: readonly Part[]): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  for (const { path, xml } of parts) {
    zip.file(path, xml);
  }
  return await zip.generateAsync({ compression: "DEFLATE", type: "arraybuffer" });
};

const document = (body: string): Part => ({
  path: "word/document.xml",
  xml: `<w:document><w:body>${body}</w:body></w:document>`,
});

const rejectedError = async (work: Promise<unknown>): Promise<unknown> => {
  try {
    await work;
  } catch (error) {
    return error;
  }
  return expect.unreachable("expected a refusal");
};

const asLimitError = (error: unknown): XmlResourceLimitError => {
  expect(error).toBeInstanceOf(XmlResourceLimitError);
  return error as XmlResourceLimitError;
};

describe("element-dense markup", () => {
  test("a part of millions of empty runs is refused at the shipped default", async () => {
    const runs = FOLIO_XML_RESOURCE_LIMITS.maxElementsPerPart + 1000;
    const buffer = await packageOf([document("<w:r/>".repeat(runs))]);

    const error = asLimitError(await rejectedError(unzipDocx(buffer)));

    expect(error.limit).toBe("elements");
    expect(error.partPath).toBe("word/document.xml");
    expect(error.allowed).toBe(FOLIO_XML_RESOURCE_LIMITS.maxElementsPerPart);
  });

  test("the scan stops at the bound, so its work is the limit's and not the input's", () => {
    // `observed` is the counter's value where it refused. That it is exactly
    // one past the bound is the proof that the remaining input was never read:
    // a scan that ran to the end would report the whole count.
    const limits = { ...FOLIO_XML_RESOURCE_LIMITS, maxElementsPerPart: 10 };
    const error = (() => {
      try {
        assertXmlResourceLimits({ xml: `<root>${"<a/>".repeat(1_000_000)}</root>`, limits });
        return expect.unreachable("expected a refusal");
      } catch (thrown) {
        return asLimitError(thrown);
      }
    })();

    expect(error.observed).toBe(11);
    expect(error.allowed).toBe(10);
  });

  test("no tree is built for a part the preflight refuses", async () => {
    // Bound before mocking: after `mock.module` the namespace's `parseXml` is
    // the wrapper, so a wrapper that reached for it again would call itself.
    const { parseXml } = xmlParser;
    let parseCalls = 0;
    mock.module("./xmlParser", () => ({
      ...xmlParser,
      parseXml: (...args: Parameters<typeof parseXml>) => {
        parseCalls += 1;
        return parseXml(...args);
      },
    }));
    try {
      // The spy has to be able to see a tree being built, or observing none
      // proves nothing: a package folio accepts must run it up first.
      await parseDocx(await packageOf([document("<w:p><w:r><w:t>a</w:t></w:r></w:p>")]));
      expect(parseCalls).toBeGreaterThan(0);

      parseCalls = 0;
      await rejectedError(
        parseDocx(await packageOf([document("<w:r/>".repeat(50))]), {
          unzipLimits: { maxXmlElementsPerPart: 10 },
        }),
      );

      expect(parseCalls).toBe(0);
    } finally {
      mock.restore();
    }
  });
});

describe("attribute-dense markup", () => {
  test("one paragraph of attribute-bearing runs is refused on attributes", async () => {
    // Few enough elements to clear the element bound, enough attributes to
    // cross the attribute bound: an element budget alone would accept this.
    const runs = Math.ceil(FOLIO_XML_RESOURCE_LIMITS.maxAttributesPerPart / 5) + 100;
    const run = `<w:r w:a="1" w:b="2" w:c="3" w:d="4" w:e="5"/>`;
    const buffer = await packageOf([document(`<w:p>${run.repeat(runs)}</w:p>`)]);

    const error = asLimitError(await rejectedError(unzipDocx(buffer)));

    expect(error.limit).toBe("attributes");
    expect(error.allowed).toBe(FOLIO_XML_RESOURCE_LIMITS.maxAttributesPerPart);
    expect(runs + 2).toBeLessThan(FOLIO_XML_RESOURCE_LIMITS.maxElementsPerPart);
  });
});

describe("a package spread thin across parts", () => {
  test("many parts, each within the per-part bound, are refused together", async () => {
    // The shape the per-part bound cannot see: no part is remarkable, the
    // package is. Before the package budget existed, only three parts were
    // counted at all and the rest of these would have been parsed.
    const perPart = 50_000;
    const parts = Math.ceil(FOLIO_XML_RESOURCE_LIMITS.maxElementsPerPackage / perPart) + 2;
    const body = "<w:r/>".repeat(perPart);
    const buffer = await packageOf([
      document(""),
      ...Array.from({ length: parts }, (_, index) => ({
        path: `word/header${String(index)}.xml`,
        xml: `<w:hdr>${body}</w:hdr>`,
      })),
    ]);

    const error = asLimitError(await rejectedError(unzipDocx(buffer)));

    expect(error.limit).toBe("package-elements");
    expect(error.allowed).toBe(FOLIO_XML_RESOURCE_LIMITS.maxElementsPerPackage);
    expect(perPart).toBeLessThan(FOLIO_XML_RESOURCE_LIMITS.maxElementsPerPart);
  });

  test("the budget is the package's, so a second part inherits the first part's spend", () => {
    const limits = { ...FOLIO_XML_RESOURCE_LIMITS, maxElementsPerPackage: 100 };
    const budget = createXmlPackageBudget();
    const part = `<root>${"<a/>".repeat(60)}</root>`;

    assertXmlResourceLimits({ xml: part, limits, budget });
    const error = (() => {
      try {
        assertXmlResourceLimits({ xml: part, limits, budget, partPath: "word/footer1.xml" });
        return expect.unreachable("expected a refusal");
      } catch (thrown) {
        return asLimitError(thrown);
      }
    })();

    expect(error.limit).toBe("package-elements");
    expect(error.partPath).toBe("word/footer1.xml");
    expect(error.observed).toBe(101);
  });
});

describe("repetitive nesting", () => {
  test("a deeply repeated table is refused on depth", async () => {
    let body = "<w:p/>";
    for (let level = 0; level < FOLIO_XML_RESOURCE_LIMITS.maxDepth; level += 1) {
      body = `<w:tbl><w:tr><w:tc>${body}</w:tc></w:tr></w:tbl>`;
    }
    const buffer = await packageOf([document(body)]);

    const error = asLimitError(await rejectedError(unzipDocx(buffer)));

    expect(error.limit).toBe("depth");
    expect(error.allowed).toBe(FOLIO_XML_RESOURCE_LIMITS.maxDepth);
    expect(error.observed).toBe(FOLIO_XML_RESOURCE_LIMITS.maxDepth + 1);
  });
});

describe("a package that inflates", () => {
  test("a small archive that expands past the ceiling is refused", async () => {
    // Four parts of 2 MiB of one repeated byte. The ceiling is lowered rather
    // than the payload raised: the shape under test is the ratio, and proving
    // it at 8 MiB costs the same assertion and 250 MiB of work.
    const filler = "x".repeat(2 * 1024 * 1024);
    const buffer = await packageOf(
      Array.from({ length: 4 }, (_, index) => ({
        path: `word/header${String(index)}.xml`,
        xml: `<w:hdr>${filler}</w:hdr>`,
      })),
    );

    expect(buffer.byteLength * 400).toBeLessThan(8 * 1024 * 1024);
    await expect(
      unzipDocx(buffer, { maxTotalUncompressedBytes: 1024 * 1024 }),
    ).rejects.toBeInstanceOf(DocxSecurityError);
  });
});

describe("the server archive reader", () => {
  test("bounds every XML part it hands out, not only the ones a caller parses", async () => {
    const buffer = await packageOf([document(`${"<w:r/>".repeat(200)}`)]);
    const archive = await loadDocxArchive(buffer, {
      xmlLimits: { maxElementsPerPart: 10 },
    });

    const error = asLimitError(await rejectedError(archive.readEntryString("word/document.xml")));

    expect(error.limit).toBe("elements");
    expect(error.partPath).toBe("word/document.xml");
  });

  test("charges a re-read part once, so reading twice is not a second package", async () => {
    const buffer = await packageOf([document("<w:r/>".repeat(40))]);
    const archive = await loadDocxArchive(buffer, {
      xmlLimits: { maxElementsPerPackage: 60 },
    });

    await archive.readEntryString("word/document.xml");
    expect(await archive.readEntryString("word/document.xml")).toContain("<w:r/>");
  });
});

describe("host configuration", () => {
  test("a host may tighten the bounds below the shipped defaults", async () => {
    const buffer = await packageOf([document("<w:r/>".repeat(100))]);

    const error = asLimitError(
      await rejectedError(unzipDocx(buffer, { maxXmlElementsPerPackage: 10 })),
    );

    expect(error.limit).toBe("package-elements");
    expect(error.allowed).toBe(10);
  });
});
