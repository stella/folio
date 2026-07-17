import { createHash } from "node:crypto";

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { createEmptyDocx } from "../rezip";
import {
  applyDocxXmlPatchProposal,
  FOLIO_DOCX_XML_PATCH_APPLICATION_PROFILE,
  FOLIO_DOCX_XML_PATCH_APPLICATION_VERSION,
  UnsupportedFolioDocxXmlPatchApplicationProfileError,
} from "./applyDocxXmlPatchProposal";
import { FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION } from "./evaluateDocxXmlPatchProposal";
import { FOLIO_DOCX_CONFORMANCE_PROFILE } from "./validateDocxConformance";

const DOCUMENT_PATH = "word/document.xml";
const STYLES_PATH = "word/styles.xml";
const REPLACEMENT_XML =
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Updated</w:t></w:r></w:p></w:body></w:document>';

const sha256 = (bytes: Uint8Array | string): string =>
  createHash("sha256").update(bytes).digest("hex");

type PackageFixture = {
  bytes: Uint8Array;
  documentXml: string;
  documentSha256: string;
  stylesXml: string;
  stylesSha256: string;
};

const packageFixture = async (): Promise<PackageFixture> => {
  const bytes = new Uint8Array(await createEmptyDocx());
  const zip = await JSZip.loadAsync(bytes);
  const document = zip.file(DOCUMENT_PATH);
  const styles = zip.file(STYLES_PATH);
  if (document === null || styles === null) {
    throw new Error("Expected document and styles package parts");
  }
  const documentBytes = await document.async("uint8array");
  const stylesBytes = await styles.async("uint8array");
  return {
    bytes,
    documentXml: new TextDecoder().decode(documentBytes),
    documentSha256: sha256(documentBytes),
    stylesXml: new TextDecoder().decode(stylesBytes),
    stylesSha256: sha256(stylesBytes),
  };
};

const proposalFor = (baseSha256: string, replacementXml = REPLACEMENT_XML) => ({
  version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
  replacements: [{ path: DOCUMENT_PATH, baseSha256, replacementXml }],
});

describe("applyDocxXmlPatchProposal", () => {
  test("applies all replacements and returns a content-addressed receipt", async () => {
    const fixture = await packageFixture();
    const replacementStylesXml = `${fixture.stylesXml}\n`;

    const result = await applyDocxXmlPatchProposal({
      bytes: fixture.bytes,
      proposal: {
        version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
        replacements: [
          {
            path: DOCUMENT_PATH,
            baseSha256: fixture.documentSha256,
            replacementXml: REPLACEMENT_XML,
          },
          {
            path: STYLES_PATH,
            baseSha256: fixture.stylesSha256,
            replacementXml: replacementStylesXml,
          },
        ],
      },
      allowedParts: [DOCUMENT_PATH, STYLES_PATH],
      validationProfile: FOLIO_DOCX_CONFORMANCE_PROFILE,
    });

    expect(result).toMatchObject({
      version: FOLIO_DOCX_XML_PATCH_APPLICATION_VERSION,
      profile: FOLIO_DOCX_XML_PATCH_APPLICATION_PROFILE,
      status: "applied",
      producesOutput: true,
      conformance: {
        profile: FOLIO_DOCX_CONFORMANCE_PROFILE,
        status: "conformant",
      },
    });
    if (result.status !== "applied") {
      throw new Error("Expected applied XML patch result");
    }
    expect(result.receipt).toMatchObject({
      input: {
        sha256: sha256(fixture.bytes),
        byteLength: fixture.bytes.byteLength,
      },
      output: {
        sha256: sha256(result.bytes),
        byteLength: result.bytes.byteLength,
      },
      replacements: [
        {
          path: DOCUMENT_PATH,
          baseSha256: fixture.documentSha256,
          replacementSha256: sha256(REPLACEMENT_XML),
        },
        {
          path: STYLES_PATH,
          baseSha256: fixture.stylesSha256,
          replacementSha256: sha256(replacementStylesXml),
        },
      ],
    });

    const outputZip = await JSZip.loadAsync(result.bytes);
    expect(await outputZip.file(DOCUMENT_PATH)?.async("text")).toBe(REPLACEMENT_XML);
    expect(await outputZip.file(STYLES_PATH)?.async("text")).toBe(replacementStylesXml);
    const inputZip = await JSZip.loadAsync(fixture.bytes);
    expect(await inputZip.file(DOCUMENT_PATH)?.async("text")).toBe(fixture.documentXml);
  });

  test("returns no output when proposal preconditions fail", async () => {
    const fixture = await packageFixture();

    const result = await applyDocxXmlPatchProposal({
      bytes: fixture.bytes,
      proposal: proposalFor("0".repeat(64)),
      allowedParts: [DOCUMENT_PATH],
      validationProfile: FOLIO_DOCX_CONFORMANCE_PROFILE,
    });

    expect(result.status).toBe("proposal-rejected");
    expect(result.producesOutput).toBe(false);
    expect(result.receipt).toBeNull();
    expect(result).not.toHaveProperty("bytes");
  });

  test("withholds output that fails the complete validation profile", async () => {
    const fixture = await packageFixture();
    const invalidDocument =
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:notBody/></w:document>';

    const result = await applyDocxXmlPatchProposal({
      bytes: fixture.bytes,
      proposal: proposalFor(fixture.documentSha256, invalidDocument),
      allowedParts: [DOCUMENT_PATH],
      validationProfile: FOLIO_DOCX_CONFORMANCE_PROFILE,
    });

    expect(result.status).toBe("output-rejected");
    expect(result.producesOutput).toBe(false);
    expect(result.conformance?.status).toBe("invalid");
    expect(result.receipt).toBeNull();
    expect(result).not.toHaveProperty("bytes");
  });

  test("rejects validation profiles that are not implemented", async () => {
    const fixture = await packageFixture();

    await expect(
      Reflect.apply(applyDocxXmlPatchProposal, undefined, [
        {
          bytes: fixture.bytes,
          proposal: proposalFor(fixture.documentSha256),
          allowedParts: [DOCUMENT_PATH],
          validationProfile: "unavailable-profile",
        },
      ]),
    ).rejects.toBeInstanceOf(UnsupportedFolioDocxXmlPatchApplicationProfileError);
  });
});
