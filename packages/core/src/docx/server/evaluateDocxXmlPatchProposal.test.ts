import { createHash } from "node:crypto";

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { createEmptyDocx } from "../rezip";
import {
  evaluateDocxXmlPatchProposal,
  FOLIO_DOCX_XML_PATCH_PROPOSAL_PROFILE,
  FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
  InvalidFolioDocxXmlPatchProposalError,
  InvalidFolioDocxXmlPatchProposalOptionsError,
  parseFolioDocxXmlPatchProposal,
} from "./evaluateDocxXmlPatchProposal";

const DOCUMENT_PATH = "word/document.xml";
const REPLACEMENT_XML =
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>';

type PackageFixture = {
  bytes: Uint8Array;
  documentSha256: string;
};

const packageFixture = async (): Promise<PackageFixture> => {
  const bytes = await createEmptyDocx();
  const zip = await JSZip.loadAsync(bytes);
  const documentBytes = await zip.file(DOCUMENT_PATH)?.async("uint8array");
  if (documentBytes === undefined) {
    throw new Error("Expected document package part");
  }
  return {
    bytes,
    documentSha256: createHash("sha256").update(documentBytes).digest("hex"),
  };
};

const proposalFor = (baseSha256: string) => ({
  version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
  replacements: [{ path: DOCUMENT_PATH, baseSha256, replacementXml: REPLACEMENT_XML }],
});

describe("parseFolioDocxXmlPatchProposal", () => {
  test("accepts the strict versioned envelope", () => {
    const proposal = proposalFor("0".repeat(64));

    expect(parseFolioDocxXmlPatchProposal(proposal)).toEqual(proposal);
  });

  test("rejects unsupported versions and unexpected properties", () => {
    expect(() =>
      parseFolioDocxXmlPatchProposal({ ...proposalFor("0".repeat(64)), version: 2 }),
    ).toThrow(InvalidFolioDocxXmlPatchProposalError);
    expect(() =>
      parseFolioDocxXmlPatchProposal({
        ...proposalFor("0".repeat(64)),
        outputPath: "modified.docx",
      }),
    ).toThrow("unexpected property");
    expect(() =>
      parseFolioDocxXmlPatchProposal({
        version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
        replacements: Array.from({ length: 17 }),
      }),
    ).toThrow("expected at most 16 replacements");
  });
});

describe("evaluateDocxXmlPatchProposal", () => {
  test("accepts an allowed replacement with an exact base hash without producing output", async () => {
    const { bytes, documentSha256 } = await packageFixture();

    const evaluation = await evaluateDocxXmlPatchProposal({
      bytes,
      proposal: proposalFor(documentSha256),
      allowedParts: [DOCUMENT_PATH],
    });

    expect(evaluation).toMatchObject({
      version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
      profile: FOLIO_DOCX_XML_PATCH_PROPOSAL_PROFILE,
      status: "accepted",
      producesOutput: false,
      issues: [],
    });
    expect(evaluation.replacements).toEqual([
      {
        path: DOCUMENT_PATH,
        baseSha256: documentSha256,
        currentSha256: documentSha256,
        replacementSha256: createHash("sha256").update(REPLACEMENT_XML).digest("hex"),
        replacementByteLength: new TextEncoder().encode(REPLACEMENT_XML).byteLength,
        encoding: "utf-8",
      },
    ]);
    expect(evaluation).not.toHaveProperty("bytes");
  });

  test("rejects stale base hashes", async () => {
    const { bytes } = await packageFixture();

    const evaluation = await evaluateDocxXmlPatchProposal({
      bytes,
      proposal: proposalFor("0".repeat(64)),
      allowedParts: [DOCUMENT_PATH],
    });

    expect(evaluation.status).toBe("rejected");
    expect(evaluation.replacements).toEqual([]);
    expect(evaluation.issues).toContainEqual({
      code: "base-hash-mismatch",
      message: "The package part changed after the proposal base was inspected.",
      proposalPath: "$.replacements[0].baseSha256",
      part: DOCUMENT_PATH,
    });
  });

  test("enforces the server-owned allowlist before package inspection", async () => {
    const evaluation = await evaluateDocxXmlPatchProposal({
      bytes: new Uint8Array(),
      proposal: proposalFor("0".repeat(64)),
      allowedParts: [],
    });

    expect(evaluation.status).toBe("rejected");
    expect(evaluation.issues.at(0)).toMatchObject({
      code: "part-not-allowed",
      part: DOCUMENT_PATH,
    });
  });

  test("rejects duplicate, unsafe-path, and invalid-hash proposals deterministically", async () => {
    const replacement = {
      path: "../document.xml",
      baseSha256: "INVALID",
      replacementXml: REPLACEMENT_XML,
    };
    const evaluation = await evaluateDocxXmlPatchProposal({
      bytes: new Uint8Array(),
      proposal: {
        version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
        replacements: [replacement, replacement],
      },
      allowedParts: [],
    });

    expect(evaluation.issues.map(({ code }) => code)).toEqual([
      "invalid-part-path",
      "invalid-base-sha256",
      "duplicate-part",
      "invalid-part-path",
      "invalid-base-sha256",
    ]);
  });

  test("rejects unsafe XML before inspection", async () => {
    const doctype = await evaluateDocxXmlPatchProposal({
      bytes: new Uint8Array(),
      proposal: {
        version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
        replacements: [
          {
            path: "custom.xml",
            baseSha256: "0".repeat(64),
            replacementXml: "<!DOCTYPE root><root/>",
          },
        ],
      },
      allowedParts: ["custom.xml"],
    });
    const malformed = await evaluateDocxXmlPatchProposal({
      bytes: new Uint8Array(),
      proposal: {
        version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
        replacements: [
          {
            path: "custom.xml",
            baseSha256: "0".repeat(64),
            replacementXml: "<root>",
          },
        ],
      },
      allowedParts: ["custom.xml"],
    });
    const encodingMismatch = await evaluateDocxXmlPatchProposal({
      bytes: new Uint8Array(),
      proposal: {
        version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
        replacements: [
          {
            path: "custom.xml",
            baseSha256: "0".repeat(64),
            replacementXml: '<?xml version="1.0" encoding="UTF-16"?><root/>',
          },
        ],
      },
      allowedParts: ["custom.xml"],
    });

    expect(doctype.issues.map(({ code }) => code)).toEqual(["xml-doctype-forbidden"]);
    expect(malformed.issues.map(({ code }) => code)).toEqual(["xml-not-well-formed"]);
    expect(encodingMismatch.issues.map(({ code }) => code)).toEqual(["xml-encoding-mismatch"]);
  });

  test("rejects replacement byte-limit violations without parsing oversized XML", async () => {
    const evaluation = await evaluateDocxXmlPatchProposal({
      bytes: new Uint8Array(),
      proposal: {
        version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
        replacements: [
          {
            path: "custom.xml",
            baseSha256: "0".repeat(64),
            replacementXml: "<!DOCTYPE root><root/>",
          },
        ],
      },
      allowedParts: ["custom.xml"],
      limits: { maxPartBytes: 4, maxTotalBytes: 4 },
    });

    expect(evaluation.issues.map(({ code }) => code)).toEqual([
      "replacement-too-large",
      "replacements-too-large",
    ]);

    const multibyteXml = "<r>é</r>";
    const multibyte = await evaluateDocxXmlPatchProposal({
      bytes: new Uint8Array(),
      proposal: {
        version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
        replacements: [
          {
            path: "custom.xml",
            baseSha256: "0".repeat(64),
            replacementXml: multibyteXml,
          },
        ],
      },
      allowedParts: ["custom.xml"],
      limits: {
        maxPartBytes: multibyteXml.length,
        maxTotalBytes: multibyteXml.length,
      },
    });

    expect(multibyte.issues.map(({ code }) => code)).toEqual([
      "replacement-too-large",
      "replacements-too-large",
    ]);
  });

  test("rejects oversized replacement arrays before parsing their entries", async () => {
    const evaluation = await evaluateDocxXmlPatchProposal({
      bytes: new Uint8Array(),
      proposal: {
        version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
        replacements: Array.from({ length: 17 }),
      },
      allowedParts: [],
    });

    expect(evaluation.issues.at(0)).toMatchObject({
      code: "too-many-replacements",
      proposalPath: "$.replacements",
    });
  });

  test("maps missing and non-XML package parts to typed rejections", async () => {
    const { bytes } = await packageFixture();
    const missing = await evaluateDocxXmlPatchProposal({
      bytes,
      proposal: {
        version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
        replacements: [
          {
            path: "missing.xml",
            baseSha256: "0".repeat(64),
            replacementXml: "<root/>",
          },
        ],
      },
      allowedParts: ["missing.xml"],
    });
    const zip = await JSZip.loadAsync(bytes);
    zip.file("word/media/payload.bin", "binary");
    const withBinary = await zip.generateAsync({ type: "uint8array" });
    const binary = await evaluateDocxXmlPatchProposal({
      bytes: withBinary,
      proposal: {
        version: FOLIO_DOCX_XML_PATCH_PROPOSAL_VERSION,
        replacements: [
          {
            path: "word/media/payload.bin",
            baseSha256: "0".repeat(64),
            replacementXml: "<root/>",
          },
        ],
      },
      allowedParts: ["word/media/payload.bin"],
    });

    expect(missing.issues.at(0)).toMatchObject({ code: "part-not-found" });
    expect(binary.issues.at(0)).toMatchObject({ code: "part-not-xml" });
  });

  test("rejects unsupported versions as data and invalid policy options as configuration", async () => {
    const unsupported = await evaluateDocxXmlPatchProposal({
      bytes: new Uint8Array(),
      proposal: { version: 2, replacements: [] },
      allowedParts: [],
    });
    const malformed = await evaluateDocxXmlPatchProposal({
      bytes: new Uint8Array(),
      proposal: { version: "2", replacements: [] },
      allowedParts: [],
    });

    expect(unsupported.issues.at(0)).toMatchObject({
      code: "unsupported-version",
      proposalPath: "$.version",
    });
    expect(malformed.issues.at(0)).toMatchObject({
      code: "invalid-proposal",
      proposalPath: "$.version",
    });
    await expect(
      evaluateDocxXmlPatchProposal({
        bytes: new Uint8Array(),
        proposal: proposalFor("0".repeat(64)),
        allowedParts: ["../document.xml"],
      }),
    ).rejects.toBeInstanceOf(InvalidFolioDocxXmlPatchProposalOptionsError);
    await expect(
      evaluateDocxXmlPatchProposal({
        bytes: new Uint8Array(),
        proposal: proposalFor("0".repeat(64)),
        allowedParts: [],
        limits: { maxReplacements: 17 },
      }),
    ).rejects.toBeInstanceOf(InvalidFolioDocxXmlPatchProposalOptionsError);
  });
});
