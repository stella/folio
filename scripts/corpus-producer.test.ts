import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import {
  PRODUCER_FAMILIES,
  producerFamilyOf,
  producerLabel,
  readCorpusProducer,
} from "./lib/corpus-producer";

const EXTENDED_PROPERTIES_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

const extendedProperties = (inner: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Properties xmlns="${EXTENDED_PROPERTIES_NAMESPACE}">${inner}</Properties>`;

const packageWith = async (parts: Record<string, string>): Promise<Uint8Array> => {
  const zip = new JSZip();
  for (const [path, content] of Object.entries(parts)) {
    zip.file(path, content);
  }
  return zip.generateAsync({ type: "uint8array" });
};

describe("producerFamilyOf", () => {
  test("places the producers the corpus actually carries", () => {
    expect(producerFamilyOf("Microsoft Office Word")).toBe(PRODUCER_FAMILIES.word);
    // The lower-case spelling is written by tools that copied Word's part.
    expect(producerFamilyOf("Microsoft word")).toBe(PRODUCER_FAMILIES.word);
    expect(producerFamilyOf("Microsoft Macintosh Word")).toBe(PRODUCER_FAMILIES.wordMac);
    expect(producerFamilyOf("Microsoft Office Outlook")).toBe(PRODUCER_FAMILIES.outlook);
    expect(producerFamilyOf("LibreOffice/7.0.6.2$Linux_X86_64 LibreOffice_project/144abb84")).toBe(
      PRODUCER_FAMILIES.libreoffice,
    );
    expect(producerFamilyOf("WPS Office_11.1.0.11664_F1E327BC")).toBe(PRODUCER_FAMILIES.wps);
    expect(producerFamilyOf("ONLYOFFICE/9.1.0.173")).toBe(PRODUCER_FAMILIES.onlyoffice);
    expect(producerFamilyOf("Apache POI")).toBe(PRODUCER_FAMILIES.apachePoi);
  });

  test("a Macintosh string does not fall through to plain Word", () => {
    // Both patterns match the string; order decides, and the order is the contract.
    expect(producerFamilyOf("Microsoft Macintosh Word")).not.toBe(PRODUCER_FAMILIES.word);
  });

  test("an empty application is unnamed, an unrecognised one is other", () => {
    expect(producerFamilyOf("")).toBe(PRODUCER_FAMILIES.unnamed);
    expect(producerFamilyOf("   ")).toBe(PRODUCER_FAMILIES.unnamed);
    expect(producerFamilyOf("Some Unknown Suite")).toBe(PRODUCER_FAMILIES.other);
  });
});

describe("producerLabel", () => {
  test("carries the major version and nothing finer", () => {
    expect(producerLabel(PRODUCER_FAMILIES.word, "16.0000")).toBe("word/16");
    expect(producerLabel(PRODUCER_FAMILIES.libreoffice, "7.0.6.2$Linux")).toBe("libreoffice/7");
  });

  test("normalises a zero-padded major so one producer has one label", () => {
    expect(producerLabel(PRODUCER_FAMILIES.word, "00.0000")).toBe(
      producerLabel(PRODUCER_FAMILIES.word, "0.0000"),
    );
  });

  test("never invents a version for a family that is an absence of evidence", () => {
    expect(producerLabel(PRODUCER_FAMILIES.unnamed, "16.0000")).toBe("unnamed");
    expect(producerLabel(PRODUCER_FAMILIES.noExtendedProperties, "16.0000")).toBe(
      "no-extended-properties",
    );
    expect(producerLabel(PRODUCER_FAMILIES.unknown, "16.0000")).toBe("unknown");
  });

  test("falls back to the family alone when no version is stated", () => {
    expect(producerLabel(PRODUCER_FAMILIES.word, "")).toBe("word");
  });
});

describe("readCorpusProducer", () => {
  test("reads Application and AppVersion", async () => {
    const bytes = await packageWith({
      "docProps/app.xml": extendedProperties(
        "<Application>Microsoft Office Word</Application><AppVersion>16.0000</AppVersion>",
      ),
    });
    expect(await readCorpusProducer({ bytes, documentPart: "word/document.xml" })).toEqual({
      family: PRODUCER_FAMILIES.word,
      label: "word/16",
    });
  });

  test("a package without the part is its own producer class", async () => {
    const bytes = await packageWith({ "word/document.xml": "<w:document/>" });
    expect(await readCorpusProducer({ bytes, documentPart: "word/document.xml" })).toEqual({
      family: PRODUCER_FAMILIES.noExtendedProperties,
      label: "no-extended-properties",
    });
  });

  test("the main part's name separates Word Online from desktop Word", async () => {
    const parts = {
      "docProps/app.xml": extendedProperties(
        "<Application>Microsoft Office Word</Application><AppVersion>16.0000</AppVersion>",
      ),
    };
    const bytes = await packageWith(parts);
    expect(await readCorpusProducer({ bytes, documentPart: "word/document2.xml" })).toEqual({
      family: PRODUCER_FAMILIES.wordOnline,
      label: "word-online/16",
    });
  });

  test("a prefixed extended-properties part is read the same way", async () => {
    const bytes = await packageWith({
      "docProps/app.xml":
        `<?xml version="1.0"?><ep:Properties xmlns:ep="${EXTENDED_PROPERTIES_NAMESPACE}">` +
        `<ep:Application>LibreOffice/7.4.7.2$Linux_X86_64</ep:Application></ep:Properties>`,
    });
    expect(await readCorpusProducer({ bytes, documentPart: "word/document.xml" })).toEqual({
      family: PRODUCER_FAMILIES.libreoffice,
      label: "libreoffice/7",
    });
  });

  test("bytes that are not an archive are unknown rather than a throw", async () => {
    const bytes = new Uint8Array([1, 2, 3, 4]);
    expect(await readCorpusProducer({ bytes, documentPart: "word/document.xml" })).toEqual({
      family: PRODUCER_FAMILIES.unknown,
      label: "unknown",
    });
  });
});
