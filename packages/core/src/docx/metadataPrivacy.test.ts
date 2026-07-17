import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createEmptyDocument } from "../utils/createDocument";
import { parseCoreProperties } from "./corePropertiesParser";
import {
  FolioDocumentPrivacyArchiveError,
  InvalidFolioDocumentPrivacyOptionsError,
  rewriteDocxMetadataPrivacy,
} from "./metadataPrivacy";
import { createDocx } from "./rezip";

const CORE_PROPERTIES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="urn:properties" xmlns:dc="urn:descriptive" xmlns:dcterms="urn:terms" xmlns:custom="urn:custom"><dc:title>Title</dc:title><dc:subject>Subject</dc:subject><dc:creator>Creator</dc:creator><cp:keywords>keywords</cp:keywords><dc:description>Description</dc:description><cp:lastModifiedBy>Modifier</cp:lastModifiedBy><cp:revision>7</cp:revision><dcterms:created>2026-07-01T10:30:00Z</dcterms:created><dcterms:modified>2026-07-02T11:45:00Z</dcterms:modified><custom:retained>keep me</custom:retained></cp:coreProperties>`;

const buildDocument = async (): Promise<ArrayBuffer> => {
  const template = createEmptyDocument();
  const buffer = await createDocx({
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [{ type: "run", content: [{ type: "text", text: "Body text." }] }],
          },
        ],
      },
    },
  });
  const zip = await JSZip.loadAsync(buffer);
  zip.file("docProps/core.xml", CORE_PROPERTIES_XML);
  return zip.generateAsync({ type: "arraybuffer" });
};

const readCorePropertiesXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file("docProps/core.xml");
  if (!file) {
    throw new Error("expected core properties");
  }
  return file.async("text");
};

const expectPrivateMetadataAbsent = async (buffer: ArrayBuffer): Promise<void> => {
  const xml = await readCorePropertiesXml(buffer);
  for (const property of [
    "title",
    "subject",
    "creator",
    "keywords",
    "description",
    "lastModifiedBy",
    "created",
    "modified",
  ]) {
    expect(xml).not.toContain(`<${property}>`);
    expect(xml).not.toContain(`:${property}>`);
  }
  expect(xml).toContain("<cp:revision>7</cp:revision>");
  expect(xml).toContain("<custom:retained>keep me</custom:retained>");
};

describe("rewriteDocxMetadataPrivacy", () => {
  test("removes selected metadata and reports fields that existed", async () => {
    const result = await rewriteDocxMetadataPrivacy(await buildDocument(), {
      transforms: ["remove-attribution", "remove-timestamps", "remove-descriptive-metadata"],
    });

    expect(result.privacyReport).toEqual({
      appliedTransforms: ["remove-attribution", "remove-timestamps", "remove-descriptive-metadata"],
      removedMetadataProperties: [
        "title",
        "subject",
        "creator",
        "keywords",
        "description",
        "lastModifiedBy",
        "created",
        "modified",
      ],
    });
    await expectPrivateMetadataAbsent(result.buffer);
    expect(parseCoreProperties(await readCorePropertiesXml(result.buffer))).toEqual({
      revision: 7,
    });
  });

  test("scrubs core properties written under a differently-cased part name", async () => {
    // OPC part names are case-insensitive; a producer that wrote
    // "docProps/Core.xml" instead of the conventional lowercase path is
    // still valid and must not be missed by the privacy rewrite.
    const template = createEmptyDocument();
    const buffer = await createDocx({
      ...template,
      package: {
        ...template.package,
        document: {
          ...template.package.document,
          content: [
            {
              type: "paragraph",
              paraId: "00000001",
              content: [{ type: "run", content: [{ type: "text", text: "Body text." }] }],
            },
          ],
        },
      },
    });
    const zip = await JSZip.loadAsync(buffer);
    zip.remove("docProps/core.xml");
    zip.file("docProps/Core.xml", CORE_PROPERTIES_XML);
    const casedBuffer = await zip.generateAsync({ type: "arraybuffer" });

    const result = await rewriteDocxMetadataPrivacy(casedBuffer, {
      transforms: ["remove-attribution", "remove-timestamps", "remove-descriptive-metadata"],
    });

    expect(result.privacyReport.removedMetadataProperties).toEqual([
      "title",
      "subject",
      "creator",
      "keywords",
      "description",
      "lastModifiedBy",
      "created",
      "modified",
    ]);

    const rewrittenZip = await JSZip.loadAsync(result.buffer);
    // Written back under the original casing, not a new lowercase entry.
    expect(rewrittenZip.file("docProps/core.xml")).toBeNull();
    const rewrittenFile = rewrittenZip.file("docProps/Core.xml");
    if (!rewrittenFile) {
      throw new Error("expected the cased core properties part to survive the rewrite");
    }
    const xml = await rewrittenFile.async("text");
    expect(xml).not.toContain("<dc:creator>");
    expect(xml).not.toContain("<cp:lastModifiedBy>");
    expect(xml).toContain("<cp:revision>7</cp:revision>");
  });

  test("keeps removed timestamps absent through selective and structural saves", async () => {
    const rewritten = await rewriteDocxMetadataPrivacy(await buildDocument(), {
      transforms: ["remove-attribution", "remove-timestamps", "remove-descriptive-metadata"],
    });
    const selectiveReviewer = await FolioDocxReviewer.fromBuffer(rewritten.buffer);
    const target = selectiveReviewer.snapshot().blocks.at(0);
    if (!target) {
      throw new Error("expected a body block");
    }
    selectiveReviewer.applyOperations(
      [{ id: "replace", type: "replaceBlock", blockId: target.id, text: "Updated body." }],
      { mode: "direct" },
    );
    const selectivelySaved = await selectiveReviewer.toBuffer();
    await expectPrivateMetadataAbsent(selectivelySaved);

    const structuralReviewer = await FolioDocxReviewer.fromBuffer(selectivelySaved);
    const structuralTarget = structuralReviewer.snapshot().blocks.at(0);
    if (!structuralTarget) {
      throw new Error("expected a body block");
    }
    structuralReviewer.applyOperations(
      [
        {
          id: "insert",
          type: "insertAfterBlock",
          blockId: structuralTarget.id,
          text: "Additional body.",
        },
      ],
      { mode: "direct" },
    );
    await expectPrivateMetadataAbsent(await structuralReviewer.toBuffer());
  });

  test("rejects unknown transforms at the public boundary", async () => {
    const buffer = await buildDocument();

    await expect(
      Reflect.apply(rewriteDocxMetadataPrivacy, undefined, [buffer, { transforms: ["unknown"] }]),
    ).rejects.toBeInstanceOf(InvalidFolioDocumentPrivacyOptionsError);
  });

  test("rejects an unreadable package with a tagged error", async () => {
    await expect(
      rewriteDocxMetadataPrivacy(new TextEncoder().encode("not a package").buffer, {
        transforms: ["remove-attribution"],
      }),
    ).rejects.toBeInstanceOf(FolioDocumentPrivacyArchiveError);
  });
});
