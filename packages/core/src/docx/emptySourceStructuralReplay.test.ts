import { expect, spyOn, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { getDocumentSourceBaseline } from "./headerFooterVerbatim";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";
import * as serializer from "./serializer/documentSerializer";
import { readRootNamespaceBindings } from "./serializer/partNamespaces";
import * as replay from "./storyBlockReplay";
import * as structural from "./structuralXmlPatch";
import * as xmlReader from "./xmlParser";

const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const WORD_14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const seed = () => {
  const template = createEmptyDocument();
  return createDocx({
    ...template,
    package: {
      ...template.package,
      document: { ...template.package.document, content: [] },
    },
  });
};

const openXml = async (xml: string): Promise<Document> => {
  const zip = await ZipArchive.read(await seed());
  zip.set("word/document.xml", xml);
  return parseDocx((await zip.generate({ compression: "store" })).buffer);
};

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const captured = (document: Document) => {
  const source = getDocumentSourceBaseline(document);
  if (source.type !== "captured") throw new Error("No captured source");
  return source;
};

const render = (document: Document) => {
  const { xml } = captured(document);
  const bindings = readRootNamespaceBindings(xml);
  return (path?: "empty-source") =>
    serializer.serializeDocumentPart(document, bindings, {
      xml,
      bodyAuthority: "canonical",
      structuralReplay: path,
      fragmentValidation: replay.createReplayFragmentValidation(),
    });
};

test("a canonical edit of a physically empty body needs no serializer block replay", async () => {
  const clock = Date.parse("2026-01-11T10:15:30Z");
  const originalDate = Date;
  const fixedDate = new Proxy(originalDate, {
    construct(target, args, newTarget) {
      return Reflect.construct(
        target,
        args.length === 0 ? [clock] : args,
        newTarget,
      );
    },
    apply() {
      return new originalDate(clock).toString();
    },
    get(target, key, receiver) {
      return key === "now" ? () => clock : Reflect.get(target, key, receiver);
    },
  });
  const input = await seed();
  const apply = async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(input, {
      session: "canonical",
    });
    const snapshot = reviewer.snapshot();
    const anchor = snapshot.blocks.at(0);
    if (!anchor) throw new Error("Missing virtual paragraph");
    const operations = Array.from({ length: 16 }, (_, index) => ({
      id: `redline-${index}`,
      type:
        index === 0 ? ("replaceBlock" as const) : ("insertAfterBlock" as const),
      blockId: anchor.id,
      text: `Inserted paragraph ${index}`,
    }));
    const result = reviewer.applyOperations(operations, {
      mode: "tracked-changes",
      snapshot,
      undefinedReferences: "refuse",
      revisionStamp: { idSeed: 1, date: "2026-01-11T10:15:30Z" },
    });
    expect(result.skipped).toEqual([]);
    expect(result.applied).toHaveLength(operations.length);
    return { result, saved: await reviewer.save({ repack: "refuse" }) };
  };

  const replays = spyOn(replay, "buildDocumentBlockReplay");
  const patches = spyOn(structural, "buildStructuralDocumentPatch");
  const serialize = serializer.serializeDocumentPart;
  const access = spyOn(serializer, "serializeDocumentPart");
  try {
    globalThis.Date = fixedDate;
    const actual = await apply();
    expect(replays).toHaveBeenCalledTimes(0);
    expect(patches).toHaveBeenCalledTimes(1);
    expect(patches.mock.calls.at(0)?.at(0)?.emptySourceParagraphCount).toBe(16);
    expect(patches.mock.results.at(0)?.value.type).toBe("patched");

    access.mockImplementation((document, bindings, options) =>
      serialize(document, bindings, {
        ...options,
        structuralReplay: undefined,
      }),
    );
    patches.mockClear();
    const expected = await apply();
    expect(replays).toHaveBeenCalledTimes(1);
    expect(patches).toHaveBeenCalledTimes(1);
    expect(
      patches.mock.calls.at(0)?.at(0)?.emptySourceParagraphCount,
    ).toBeUndefined();
    expect(actual.result).toEqual(expected.result);
    expect(actual.saved.type).toBe("selective");
    expect(expected.saved.type).toBe("selective");
    if (
      actual.saved.type === "selective" &&
      expected.saved.type === "selective"
    ) {
      expect(
        Buffer.compare(
          Buffer.from(actual.saved.buffer),
          Buffer.from(expected.saved.buffer),
        ),
      ).toBe(0);
    }
  } finally {
    globalThis.Date = originalDate;
    access.mockRestore();
    patches.mockRestore();
    replays.mockRestore();
  }
});

test.each([
  [
    "authored paragraph",
    WORD,
    "",
    '<w:p w14:paraId="00000001"><w:r><w:t>Kept</w:t></w:r></w:p>',
    undefined,
  ],
  ["Strict part", STRICT, "", "", "structuralUnreadableBody"],
  [
    "authored Ignorable order",
    WORD,
    ' mc:Ignorable="w16 w14" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"',
    "",
    "structuralBodyBlock",
  ],
] as const)(
  "%s keeps the original complete replay",
  async (_, namespace, extraRoot, content, reason) => {
    const document = await openXml(
      `<w:document xmlns:w="${namespace}" xmlns:w14="${WORD_14}" xmlns:mc="${MC}"${extraRoot}><w:body>${content}<w:sectPr><w:pgMar w:top="1440"/></w:sectPr></w:body></w:document>`,
    );
    document.package.document.content.push(paragraph("00000002", "Inserted"));
    const serialized = render(document);
    const baseline = serialized();
    const replays = spyOn(replay, "buildDocumentBlockReplay");
    const patches = spyOn(structural, "buildStructuralDocumentPatch");
    try {
      const result = serialized("empty-source");
      expect(result).toEqual(baseline);
      expect(result.structuralPatch).toBeUndefined();
      expect(replays).toHaveBeenCalledTimes(1);
      if (reason === undefined) expect(patches).toHaveBeenCalledTimes(0);
      else {
        expect(patches).toHaveBeenCalledTimes(1);
        expect(patches.mock.results.at(0)?.value).toEqual({
          type: "declined",
          reason,
        });
      }
    } finally {
      patches.mockRestore();
      replays.mockRestore();
    }
  },
);

test("an authored body namespace not present in generated scope keeps the full replay", async () => {
  const document = await openXml(
    `<w:document xmlns:w="${WORD}" xmlns:w14="${WORD_14}"><w:body xmlns:local="urn:authored:local"><w:sectPr><w:pgMar w:top="1440"/></w:sectPr></w:body></w:document>`,
  );
  document.package.document.content = [paragraph("00000001", "Inserted")];
  const serialize = render(document);
  const baseline = serialize();
  const replays = spyOn(replay, "buildDocumentBlockReplay");
  const patches = spyOn(structural, "buildStructuralDocumentPatch");
  try {
    expect(serialize("empty-source")).toEqual(baseline);
    expect(patches).toHaveBeenCalledTimes(1);
    expect(patches.mock.results.at(0)?.value).toEqual({
      type: "declined",
      reason: "structuralBodyBlock",
    });
    expect(replays).toHaveBeenCalledTimes(1);
  } finally {
    patches.mockRestore();
    replays.mockRestore();
  }
});

test("a direct structural refusal falls back to the original replay", async () => {
  const document = await parseDocx(await seed());
  document.package.document.content = [
    paragraph("01000001", "First"),
    paragraph("01000001", "Repeated"),
  ];
  const serialized = render(document);
  const baseline = serialized();
  const replays = spyOn(replay, "buildDocumentBlockReplay");
  const patches = spyOn(structural, "buildStructuralDocumentPatch");
  try {
    expect(serialized("empty-source")).toEqual(baseline);
    expect(patches).toHaveBeenCalledTimes(1);
    expect(patches.mock.results.at(0)?.value).toEqual({
      type: "declined",
      reason: "structuralParagraphIds",
    });
    expect(replays).toHaveBeenCalledTimes(1);
    const { xml } = captured(document);
    expect(
      structural.buildStructuralDocumentPatch({
        originalXml: xml,
        serializedXml: baseline.xml,
        changedIds: new Set(),
      }),
    ).toEqual({ type: "declined", reason: "structuralParagraphIds" });
  } finally {
    patches.mockRestore();
    replays.mockRestore();
  }
});

test("a refusal of the expanded body still reaches the original path", async () => {
  const document = await parseDocx(await seed());
  document.package.document.content = [
    paragraph("01000001", "First"),
    paragraph("01000002", "Second"),
  ];
  const serialized = render(document);
  const baseline = serialized();
  const { xml } = captured(document);
  const read = xmlReader.parseXmlDocument;
  // Reject only the full expanded body, as an output limit could. The compact
  // generated document and the individually validated fragments still read.
  // Written for the source part, the generated document can spell the
  // expanded body exactly; it is read first, so only later reads refuse it.
  let generatedRead = false;
  const budget = spyOn(xmlReader, "parseXmlDocument").mockImplementation(
    (value) => {
      if (value !== baseline.xml) return read(value);
      if (generatedRead) return null;
      generatedRead = true;
      return read(value);
    },
  );
  const replays = spyOn(replay, "buildDocumentBlockReplay");
  const patches = spyOn(structural, "buildStructuralDocumentPatch");
  try {
    expect(serialized("empty-source")).toEqual(baseline);
    expect(patches).toHaveBeenCalledTimes(1);
    expect(patches.mock.results.at(0)?.value).toEqual({
      type: "declined",
      reason: "structuralSpliceRefused",
    });
    expect(replays).toHaveBeenCalledTimes(1);
    expect(budget).toHaveBeenCalledWith(baseline.xml);
    expect(
      structural.buildStructuralDocumentPatch({
        originalXml: xml,
        serializedXml: baseline.xml,
        changedIds: new Set(),
      }),
    ).toEqual({
      type: "declined",
      reason: "structuralUnreadableBody",
    });
  } finally {
    patches.mockRestore();
    replays.mockRestore();
    budget.mockRestore();
  }
});
