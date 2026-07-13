import { describe, expect, test } from "bun:test";
import path from "node:path";
import { DOCX_CONFORMANCE_CLASSES } from "@stll/docx-core/model";

import type { Document } from "../types/document";
import {
  FOLIO_DOCX_CAPABILITY_IDS,
  FOLIO_DOCX_CAPABILITY_MANIFEST,
  FOLIO_DOCX_CAPABILITY_MANIFEST_VERSION,
  getFolioDocxCapability,
  InvalidFolioDocxCapabilityIdError,
} from "./capabilities";
import { inspectDocxCompatibility } from "./compatibility";

const createDocument = (rawXml?: string, paraId?: string): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          ...(paraId === undefined ? {} : { paraId }),
          content: [
            {
              type: "run",
              content: [
                { type: "text", text: "Body" },
                {
                  type: "drawing",
                  image: {
                    type: "image",
                    rId: "rId1",
                    size: { width: 9525, height: 9525 },
                    wrap: { type: "inline" },
                  },
                  ...(rawXml !== undefined ? { rawXml } : {}),
                },
              ],
            },
          ],
        },
      ],
    },
  },
});

describe("DOCX compatibility inspection", () => {
  test("allows editing ordinary parsed content", () => {
    expect(inspectDocxCompatibility(createDocument())).toEqual({
      schemaVersion: 1,
      context: { host: "unknown", profile: "unknown" },
      canSafelyEdit: true,
      issues: [],
      reasons: [],
      unsupportedContentCount: 0,
    });
  });

  test("blocks editing when the document contains opaque drawing XML", () => {
    expect(inspectDocxCompatibility(createDocument("<w:drawing/>", "A1B2C3D4"))).toEqual({
      schemaVersion: 1,
      context: { host: "unknown", profile: "unknown" },
      canSafelyEdit: false,
      issues: [
        {
          code: "opaqueDrawing",
          capability: getFolioDocxCapability("opaqueDrawing"),
          coverage: { host: "unknown", profile: "unknown" },
          location: {
            blockId: "A1B2C3D4",
            part: { type: "document" },
            path: "package.document.content[0].content[0].content[1]",
          },
        },
      ],
      reasons: ["opaqueDrawing"],
      unsupportedContentCount: 1,
    });
  });

  test("reports the requested profile, host, and non-body part", () => {
    const document = createDocument("<w:drawing/>", "A1B2C3D4");
    const content = document.package.document.content;
    document.package.document.content = [];
    document.package.headers = new Map([
      [
        "rId7",
        {
          type: "header",
          hdrFtrType: "default",
          content,
        },
      ],
    ]);

    const compatibility = inspectDocxCompatibility(document, {
      host: "browser",
      profile: "transitional",
    });

    expect(compatibility.context).toEqual({ host: "browser", profile: "transitional" });
    expect(compatibility.issues.at(0)?.coverage).toEqual({
      host: "covered",
      profile: "covered",
    });
    expect(compatibility.issues.at(0)?.location).toEqual({
      blockId: "A1B2C3D4",
      part: { type: "header", relationshipId: "rId7" },
      path: 'package.headers.get("rId7").content[0].content[0].content[1]',
    });
  });

  test("marks profiles outside an issue's evidence scope as unverified", () => {
    const compatibility = inspectDocxCompatibility(createDocument("<w:drawing/>"), {
      host: "server",
      profile: "strict",
    });

    expect(compatibility.issues.at(0)?.coverage).toEqual({
      host: "covered",
      profile: "unverified",
    });
  });

  test("uses parsed package metadata unless the caller overrides it", () => {
    const document = createDocument("<w:drawing/>");
    document.package.conformanceClass = DOCX_CONFORMANCE_CLASSES.STRICT;

    expect(inspectDocxCompatibility(document).context.profile).toBe(
      DOCX_CONFORMANCE_CLASSES.STRICT,
    );
    expect(
      inspectDocxCompatibility(document, {
        profile: DOCX_CONFORMANCE_CLASSES.TRANSITIONAL,
      }).context.profile,
    ).toBe(DOCX_CONFORMANCE_CLASSES.TRANSITIONAL);
  });
});

describe("DOCX capability manifest", () => {
  test("describes operation support for structured compatibility issues", async () => {
    expect(FOLIO_DOCX_CAPABILITY_MANIFEST.version).toBe(FOLIO_DOCX_CAPABILITY_MANIFEST_VERSION);
    expect(getFolioDocxCapability("opaqueDrawing")).toEqual({
      id: "opaqueDrawing",
      feature: "drawings",
      hosts: ["browser", "server"],
      profiles: ["transitional"],
      support: {
        create: "unsupported",
        edit: "unsupported",
        preserve: "supported",
        read: "supported",
        render: "partial",
      },
      evidence: [
        {
          type: "test",
          path: "packages/core/src/docx/compatibility.test.ts",
        },
        {
          type: "test",
          path: "packages/core/src/docx/rezip.test.ts",
        },
      ],
    });
  });

  test("covers core document and review features with repository evidence", async () => {
    expect(FOLIO_DOCX_CAPABILITY_MANIFEST.capabilities.paragraphs.support).toEqual({
      create: "supported",
      edit: "supported",
      preserve: "supported",
      read: "supported",
      render: "partial",
    });
    expect(FOLIO_DOCX_CAPABILITY_MANIFEST.capabilities.tables.feature).toBe("tables");
    expect(FOLIO_DOCX_CAPABILITY_MANIFEST.capabilities.comments.feature).toBe("comments");
    expect(FOLIO_DOCX_CAPABILITY_MANIFEST.capabilities.trackedChanges.feature).toBe("revisions");

    const repositoryRoot = path.resolve(import.meta.dir, "../../../..");
    for (const capability of Object.values(FOLIO_DOCX_CAPABILITY_MANIFEST.capabilities)) {
      for (const evidence of capability.evidence) {
        expect(await Bun.file(path.join(repositoryRoot, evidence.path)).exists()).toBe(true);
      }
    }
  });

  test("describes document structure support conservatively", () => {
    expect(FOLIO_DOCX_CAPABILITY_MANIFEST.capabilities.headersFooters.support.edit).toBe(
      "supported",
    );
    expect(FOLIO_DOCX_CAPABILITY_MANIFEST.capabilities.notes.support.create).toBe("partial");
    expect(FOLIO_DOCX_CAPABILITY_MANIFEST.capabilities.numbering.support.create).toBe("partial");
    expect(FOLIO_DOCX_CAPABILITY_MANIFEST.capabilities.sections.support.create).toBe("supported");
    expect(FOLIO_DOCX_CAPABILITY_MANIFEST.capabilities.styles.support).toEqual({
      create: "partial",
      edit: "partial",
      preserve: "supported",
      read: "supported",
      render: "partial",
    });
  });

  test("keeps the public id list aligned with manifest keys", () => {
    expect(Object.keys(FOLIO_DOCX_CAPABILITY_MANIFEST.capabilities)).toEqual(
      FOLIO_DOCX_CAPABILITY_IDS,
    );
  });

  test("rejects unknown ids at the public lookup boundary", () => {
    expect(() => getFolioDocxCapability("__proto__")).toThrow(InvalidFolioDocxCapabilityIdError);
  });
});
