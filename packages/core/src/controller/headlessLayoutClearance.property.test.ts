import { describe, expect, test } from "bun:test";
import { writeZipPart } from "@stll/docx-core/zip";
import { panic } from "better-result";
import fc from "fast-check";
import JSZip from "jszip";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { installParityMeasurement, layoutBothPaths } from "./__tests__/headlessLayoutParity";
import type { Layout } from "../layout-engine/types";
import type { HeaderFooterVariant } from "../paged-layout/headerFooterMarginClearance";
import { getMargins } from "../paged-layout/sectionGeometry";

installParityMeasurement();

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const SIDES = ["header", "footer"] as const;

const storyHeights = fc.record({
  // Twips: 12px through 240px, crossing both authored body margins.
  header: fc.integer({ min: 180, max: 3600 }),
  footer: fc.integer({ min: 180, max: 3600 }),
});
const heightVariants = {
  default: storyHeights,
  first: storyHeights,
  even: storyHeights,
} as const satisfies Record<HeaderFooterVariant, typeof storyHeights>;

const clearanceCase = fc.record({
  heights: fc.record(heightVariants),
  titlePage: fc.boolean(),
  evenAndOdd: fc.boolean(),
  secondTop: fc.integer({ min: 1440, max: 2160 }),
  secondBottom: fc.integer({ min: 1440, max: 2160 }),
});

type ClearanceCase = fc.ArbitraryValue<typeof clearanceCase>;

const makePackage = async (spec: ClearanceCase): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  const parts = new Map<string, string>();
  const overrides = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>',
  ];
  const relationships = [
    `<Relationship Id="rIdSettings" Type="${R}/settings" Target="settings.xml"/>`,
  ];
  for (const [variant, heights] of Object.entries(spec.heights)) {
    for (const side of SIDES) {
      const name = `${side}-${variant}`;
      overrides.push(
        `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${side}+xml"/>`,
      );
      relationships.push(
        `<Relationship Id="rId-${name}" Type="${R}/${side}" Target="${name}.xml"/>`,
      );
      const tag = side === "header" ? "hdr" : "ftr";
      parts.set(
        `word/${name}.xml`,
        `<w:${tag} xmlns:w="${W}"><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${heights[side]}" w:lineRule="exact"/></w:pPr><w:r><w:t>${name}</w:t></w:r></w:p></w:${tag}>`,
      );
    }
  }
  const refs = Object.keys(spec.heights)
    .map(
      (variant) =>
        `<w:headerReference w:type="${variant}" r:id="rId-header-${variant}"/><w:footerReference w:type="${variant}" r:id="rId-footer-${variant}"/>`,
    )
    .join("");
  const section = (index: number) =>
    `<w:sectPr>${refs}<w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="${index === 0 ? 1080 : spec.secondTop}" w:bottom="${index === 0 ? 1080 : spec.secondBottom}" w:left="1080" w:right="1080" w:header="360" w:footer="360"/>${spec.titlePage ? "<w:titlePg/>" : ""}</w:sectPr>`;
  const body: string[] = [];
  for (let sectionIndex = 0; sectionIndex < 2; sectionIndex++) {
    // Enough body text for first, even, and subsequent odd pages in each
    // section, even with the shortest furniture and largest available body.
    for (let index = 0; index < 96; index++) {
      const id = (sectionIndex * 1000 + index + 1).toString(16).padStart(8, "0");
      const boundary = sectionIndex === 0 && index === 95 ? section(0) : "";
      body.push(
        `<w:p w14:paraId="${id}"><w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="exact"/>${boundary}</w:pPr><w:r><w:t>Section ${sectionIndex}, paragraph ${index}</w:t></w:r></w:p>`,
      );
    }
  }
  parts.set(
    "word/document.xml",
    `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="${W14}"><w:body>${body.join("")}${section(1)}</w:body></w:document>`,
  );
  parts.set(
    "word/settings.xml",
    `<w:settings xmlns:w="${W}">${spec.evenAndOdd ? "<w:evenAndOddHeaders/>" : ""}</w:settings>`,
  );
  parts.set(
    "_rels/.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDocument" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  parts.set(
    "word/_rels/document.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join("")}</Relationships>`,
  );
  parts.set(
    "[Content_Types].xml",
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.join("")}</Types>`,
  );
  for (const [path, data] of parts) writeZipPart({ zip, path, data });
  return zip.generateAsync({ type: "arraybuffer" });
};

const assertClearance = async (spec: ClearanceCase): Promise<void> => {
  const { headless, editor } = await layoutBothPaths(await makePackage(spec));
  const sections = headless.document.package.document.sections;
  expect(sections).toHaveLength(2);
  const assertPath = (layout: Layout) => {
    const exercised = new Set<string>();
    for (const page of layout.pages) {
      const sectionIndex = page.sectionIndex;
      if (sectionIndex === undefined) return panic("Expected a section index on every page");
      const section = sections?.at(sectionIndex);
      if (!section) return panic("Page addresses an absent section", { sectionIndex });
      let variant = "default";
      if (spec.titlePage && page.sectionPageNumber === 1) variant = "first";
      else if (spec.evenAndOdd && page.logicalNumber % 2 === 0) variant = "even";
      exercised.add(`${sectionIndex}:${variant}`);
      const margins = getMargins(section.properties);
      const header = headless.furniture.headerContentByRId?.get(`rId-header-${variant}`);
      const footer = headless.furniture.footerContentByRId?.get(`rId-footer-${variant}`);
      if (!header || !footer) return panic("Generated furniture must be measured", { variant });
      // The oracle reads the measured story bounds directly; it never calls
      // the clearance helper whose integration this property guards.
      const top = Math.max(margins.top, 24 + (header.marginPushBottom ?? header.height));
      const bottom = Math.max(margins.bottom, 24 + (footer.marginPushBottom ?? footer.height));
      expect(page.margins.top).toBe(top);
      expect(page.margins.bottom).toBe(bottom);
      expect(page.fragments.length).toBeGreaterThan(0);
      for (const fragment of page.fragments) {
        expect(fragment.y).toBeGreaterThanOrEqual(top);
        expect(fragment.y + fragment.height).toBeLessThanOrEqual(page.size.h - bottom + 0.001);
      }
    }
    const expected: string[] = [];
    for (const index of [0, 1]) {
      expected.push(`${index}:default`);
      if (spec.titlePage) expected.push(`${index}:first`);
      if (spec.evenAndOdd) expected.push(`${index}:even`);
    }
    expect([...exercised].sort()).toEqual(expected.sort());
  };
  assertPath(headless.layout);
  assertPath(editor);
};

// Helper-only clearance tests and headless-only pagination tests did not
// compare the two composition roots. This generated cross-path oracle adds
// furniture height, variant activation, and section geometry dimensions.
describe("headless/editor header and footer clearance", () => {
  test.each([180, 720, 3600])(
    "clears all variants and a changed section at %i twips",
    async (height) => {
      const heights = { header: height, footer: height };
      await assertClearance({
        heights: { default: heights, first: heights, even: heights },
        titlePage: true,
        evenAndOdd: true,
        secondTop: 1440,
        secondBottom: 1800,
      });
    },
    30_000,
  );

  test(
    "keeps body fragments clear and page breaks equal across section and variant geometry",
    async () => {
      await assertProperty(fc.asyncProperty(clearanceCase, assertClearance), {
        id: "headless-editor-section-variant-margin-clearance",
        numRuns: 50,
      });
    },
    propertyTestTimeout(90_000),
  );
});
