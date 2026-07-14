import { describe, expect, test } from "bun:test";

import { parseCoreProperties } from "./corePropertiesParser";

describe("parseCoreProperties", () => {
  test("parses recognized values independently of namespace prefixes", () => {
    const properties = parseCoreProperties(`<?xml version="1.0" encoding="UTF-8"?>
      <meta:coreProperties xmlns:meta="urn:properties" xmlns:terms="urn:terms">
        <meta:title>Agreement</meta:title>
        <meta:subject>Review</meta:subject>
        <meta:creator>Initial author</meta:creator>
        <meta:keywords>contract, draft</meta:keywords>
        <meta:description>Working copy</meta:description>
        <meta:lastModifiedBy>Reviewer</meta:lastModifiedBy>
        <meta:revision>12</meta:revision>
        <terms:created>2026-07-01T10:30:00Z</terms:created>
        <terms:modified>2026-07-02T11:45:00Z</terms:modified>
      </meta:coreProperties>`);

    expect(properties).toEqual({
      title: "Agreement",
      subject: "Review",
      creator: "Initial author",
      keywords: "contract, draft",
      description: "Working copy",
      lastModifiedBy: "Reviewer",
      revision: 12,
      created: new Date("2026-07-01T10:30:00Z"),
      modified: new Date("2026-07-02T11:45:00Z"),
    });
  });

  test("omits malformed typed values and absent metadata", () => {
    expect(
      parseCoreProperties(`
        <cp:coreProperties xmlns:cp="urn:properties" xmlns:dcterms="urn:terms">
          <cp:revision>1.5</cp:revision>
          <dcterms:created>not-a-date</dcterms:created>
        </cp:coreProperties>`),
    ).toBeUndefined();
    expect(parseCoreProperties(null)).toBeUndefined();
    expect(parseCoreProperties("not xml")).toBeUndefined();
  });
});
