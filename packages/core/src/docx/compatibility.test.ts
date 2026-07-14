import { describe, expect, test } from "bun:test";
import { DOCX_CONFORMANCE_CLASSES } from "@stll/docx-core/model";

import type { Document } from "../types/document";
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

  test("reports opaque drawing XML at its document location", () => {
    expect(inspectDocxCompatibility(createDocument("<w:drawing/>", "A1B2C3D4"))).toEqual({
      schemaVersion: 1,
      context: { host: "unknown", profile: "unknown" },
      canSafelyEdit: false,
      issues: [
        {
          code: "opaqueDrawing",
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
    expect(compatibility.issues.at(0)?.location).toEqual({
      blockId: "A1B2C3D4",
      part: { type: "header", relationshipId: "rId7" },
      path: 'package.headers.get("rId7").content[0].content[0].content[1]',
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
