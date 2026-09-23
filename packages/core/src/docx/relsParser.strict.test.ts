import { describe, expect, test } from "bun:test";

import {
  getHeaders,
  getImages,
  getRelationshipTypeName,
  isExternalHyperlink,
  parseRelationships,
  RELATIONSHIP_TYPES,
} from "./relsParser";

const STRICT_RELATIONSHIP_PREFIX = "http://purl.oclc.org/ooxml/officeDocument/relationships/";

describe("Strict relationship types", () => {
  test("normalizes every known office relationship type", () => {
    const transitionalPrefix =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    for (const type of Object.values(RELATIONSHIP_TYPES)) {
      if (!type.startsWith(transitionalPrefix)) {
        continue;
      }

      const strictType = STRICT_RELATIONSHIP_PREFIX + type.slice(transitionalPrefix.length);
      const relationships = parseRelationships(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="known" Type="${strictType}" Target="part.xml"/>` +
          `</Relationships>`,
      );
      expect(relationships.get("known")?.type).toBe(type);
    }
  });

  test("resolves known Strict relationships through the document readers", () => {
    const relationships = parseRelationships(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="header" Type="${STRICT_RELATIONSHIP_PREFIX}header" Target="header1.xml"/>` +
        `<Relationship Id="image" Type="${STRICT_RELATIONSHIP_PREFIX}image" Target="media/image1.png"/>` +
        `<Relationship Id="link" Type="${STRICT_RELATIONSHIP_PREFIX}hyperlink" Target="https://example.com" TargetMode="External"/>` +
        `</Relationships>`,
    );

    expect(getHeaders(relationships).map(({ id }) => id)).toEqual(["header"]);
    expect(getImages(relationships).map(({ id }) => id)).toEqual(["image"]);
    expect(isExternalHyperlink(relationships.get("link")!)).toBe(true);
    expect(relationships.get("header")?.type).toBe(RELATIONSHIP_TYPES.header);
    expect(getRelationshipTypeName(relationships.get("image")!.type)).toBe("image");
  });

  test("leaves unknown Strict relationship types intact", () => {
    const unknown = `${STRICT_RELATIONSHIP_PREFIX}futureType`;
    const relationships = parseRelationships(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="future" Type="${unknown}" Target="future.xml"/>` +
        `</Relationships>`,
    );

    expect(relationships.get("future")?.type).toBe(unknown);
  });
});
