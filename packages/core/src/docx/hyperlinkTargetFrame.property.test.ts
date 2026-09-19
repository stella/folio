/**
 * `w:hyperlink/@w:tgtFrame` says what the document says.
 *
 * A frame name outside `{_blank,_self,_parent,_top}` is legal OOXML, and the
 * parser used to map every one of them to `_blank`, so a saved file no longer
 * said what the source said. The allow-list is a DOM concern: a named frame
 * addresses another browsing context, which `anchorTargetAttrs` clamps where an
 * anchor or a navigation is produced, not where a package is read.
 *
 * Two laws, one for each side of the split:
 * - whatever frame a document authors survives parse and save; and
 * - whatever frame a document authors, no anchor folio produces targets it.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { anchorTargetAttrs } from "../utils/urlSecurity";

import { parseDocumentBody } from "./documentParser";
import { serializeParagraph } from "./serializer/paragraphSerializer";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const ALLOWED_TARGETS = ["_blank", "_self", "_parent", "_top"];

const documentXml = (frame: string): string =>
  `${XML_DECLARATION}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:hyperlink r:id="rId1" w:tgtFrame="${frame}"><w:r><w:t>Schedule 2</w:t></w:r></w:hyperlink></w:p></w:body></w:document>`;

const relationships = new Map([
  [
    "rId1",
    {
      id: "rId1",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      target: "https://example.test/schedule",
      targetMode: "External" as const,
    },
  ],
]);

const parsedFrame = (frame: string): string | undefined => {
  const body = parseDocumentBody(documentXml(frame), null, null, null, relationships);
  const paragraph = body.content.at(0);
  if (paragraph?.type !== "paragraph") {
    throw new Error("expected a paragraph");
  }
  const hyperlink = paragraph.content.find((child) => child.type === "hyperlink");
  return hyperlink?.type === "hyperlink" ? hyperlink.target : undefined;
};

/**
 * A frame name Word accepts: `w:tgtFrame` is `ST_String`, so anything but the
 * XML metacharacters the fixture would have to escape.
 */
const frameArbitrary = fc.oneof(
  fc.constantFrom(...ALLOWED_TARGETS),
  fc
    .stringMatching(/^[A-Za-z_][A-Za-z0-9_-]{0,23}$/u)
    .filter((frame) => !ALLOWED_TARGETS.includes(frame)),
);

describe("authored hyperlink target frames", () => {
  test(
    "survive parse and re-serialization verbatim",
    () => {
      fc.assert(
        fc.property(frameArbitrary, (frame) => {
          expect(parsedFrame(frame)).toBe(frame);

          // Re-serialized from the model, with no captured markup to replay:
          // what the serializer writes is what the model kept.
          const body = parseDocumentBody(documentXml(frame), null, null, null, relationships);
          const paragraph = body.content.at(0);
          if (paragraph?.type !== "paragraph") {
            throw new Error("expected a paragraph");
          }
          expect(serializeParagraph({ ...paragraph, verbatimXml: undefined })).toContain(
            `w:tgtFrame="${frame}"`,
          );
        }),
        propertyConfig({ numRuns: 60 }),
      );
    },
    propertyTestTimeout(30_000),
  );

  test(
    "never reach a DOM anchor outside the allow-list",
    () => {
      fc.assert(
        fc.property(frameArbitrary, (frame) => {
          const { target, rel } = anchorTargetAttrs(parsedFrame(frame));
          expect(ALLOWED_TARGETS).toContain(target);
          expect(rel).toBe("noopener noreferrer");
        }),
        propertyConfig({ numRuns: 60 }),
      );
    },
    propertyTestTimeout(30_000),
  );
});
