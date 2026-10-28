import { panic } from "better-result";
import { describe, expect, setDefaultTimeout, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { propertyTestTimeout } from "../../../../test/property-testing";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type {
  BlockContent,
  DocumentBody,
  HeaderFooter,
  Paragraph,
  Run,
} from "../types/document";
import { buildSections } from "./documentParser";
import {
  BLOCK_TREE_DESCENT,
  visitBlockTreeRecords,
  visitDocxContentSlots,
  visitParagraphRuns,
} from "./paragraphTraversal";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";
import { getDefaultSectionProperties } from "./sectionParser";
import { normalizeTrackedMoveRanges } from "./trackedMoveRangeNormalization";

setDefaultTimeout(propertyTestTimeout(30_000));

const PROFILES = {
  transitional: {
    namespace: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    prefix: "word",
  },
  strict: {
    namespace: "http://purl.oclc.org/ooxml/wordprocessingml/main",
    prefix: "strictWord",
  },
} as const;

const HOSTS = {
  body: (content: string) => content,
  cell: (content: string) =>
    `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${content}<w:p/></w:tc></w:tr></w:tbl><w:p/>`,
  blockSdt: (content: string) =>
    `<w:sdt><w:sdtPr/><w:sdtContent>${content}</w:sdtContent></w:sdt><w:p/>`,
  blockCustomXml: (content: string) =>
    `<w:customXml w:element="example">${content}</w:customXml><w:p/>`,
} as const;

const PLACEMENTS = {
  inlineStart: ["inline", "block"],
  blockStart: ["block", "inline"],
} as const;

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const paragraph = (content: string) => `<w:p>${content}</w:p>`;

const shapeRun = (content: Paragraph) =>
  ({
    type: "run",
    content: [
      {
        type: "shape",
        shape: {
          type: "shape",
          shapeType: "rect",
          size: { width: 1, height: 1 },
          textBody: { content: [content] },
        },
      },
    ],
  }) satisfies Run;

type MixedBodyOptions = {
  move: "moveFrom" | "moveTo";
  placement: keyof typeof PLACEMENTS;
};

const mixedBody = ({ move, placement }: MixedBodyOptions) => {
  const start = `<w:${move}RangeStart w:id="7" w:name="range" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"/>`;
  const end = `<w:${move}RangeEnd w:id="7"/>`;
  switch (placement) {
    case "inlineStart":
      return `${paragraph(`${run("before")}${start}${run("inside")}`)}${end}${paragraph(run("after"))}`;
    case "blockStart":
      return `${paragraph(run("before"))}${start}${paragraph(`${run("inside")}${end}${run("after")}`)}`;
    default:
      return placement satisfies never;
  }
};

type BuildDocxOptions = {
  body: string;
  profile: (typeof PROFILES)[keyof typeof PROFILES];
};

const buildDocx = async ({ body, profile }: BuildDocxOptions) => {
  const zip = await ZipArchive.read(await createEmptyDocx());
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><${profile.prefix}:document xmlns:${profile.prefix}="${profile.namespace}"><${profile.prefix}:body>${body.replaceAll("w:", `${profile.prefix}:`)}<${profile.prefix}:sectPr/></${profile.prefix}:body></${profile.prefix}:document>`,
  );
  return zip.generate({ compression: "store" }).then((bytes) => bytes.buffer);
};

const markers = (documentBody: DocumentBody) => {
  const found: {
    ownership: "block" | "inline";
    marker: BlockContent | Paragraph["content"][number];
  }[] = [];
  visitDocxContentSlots({ documentBody }, (slot) => {
    switch (slot.item.type) {
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
        found.push({ ownership: slot.type, marker: slot.item });
        return;
      default:
        return;
    }
  });
  return found;
};

describe("mixed block and inline move range normalization", () => {
  test("preserves both families at their authored slots through editor and save", async () => {
    const exercisedHosts = new Set<string>();
    const exercisedProfiles = new Set<string>();
    const exercisedPlacements = new Set<string>();
    for (const [profileName, profile] of Object.entries(PROFILES)) {
      for (const [hostName, host] of Object.entries(HOSTS)) {
        for (const placement of Object.keys(PLACEMENTS)) {
          if (placement !== "inlineStart" && placement !== "blockStart") {
            panic("Unexpected move range placement");
          }
          for (const move of ["moveFrom", "moveTo"] as const) {
            const parsed = await parseDocx(
              await buildDocx({
                body: host(mixedBody({ move, placement })),
                profile,
              }),
              { preloadFonts: false },
            );
            const originalMarkers = markers(parsed.package.document);
            expect(originalMarkers).toEqual([
              {
                ownership: PLACEMENTS[placement][0],
                marker: {
                  type: `${move}RangeStart`,
                  id: 7,
                  name: "range",
                  author: "Reviewer",
                  date: "2024-01-01T00:00:00Z",
                },
              },
              {
                ownership: PLACEMENTS[placement][1],
                marker: { type: `${move}RangeEnd`, id: 7 },
              },
            ]);
            expect(
              normalizeTrackedMoveRanges({
                documentBody: parsed.package.document,
              }),
            ).toEqual({
              removedUnbalancedMoveRangeMarkers: 0,
            });
            const edited = fromProseDoc(toProseDoc(parsed), parsed, {
              stylesheetSource: { type: "package" },
            });
            expect(markers(edited.package.document)).toEqual(originalMarkers);
            const saved = await repackDocx(edited, {
              updateModifiedDate: false,
            });
            const reparsed = await parseDocx(saved, { preloadFonts: false });
            expect(markers(reparsed.package.document)).toEqual(originalMarkers);
            expect(reparsed.warnings ?? []).not.toContain(
              "Removed 2 unbalanced tracked move range markers.",
            );
            exercisedHosts.add(hostName);
            exercisedProfiles.add(profileName);
            exercisedPlacements.add(placement);
          }
        }
      }
    }
    expect(exercisedHosts).toEqual(new Set(Object.keys(HOSTS)));
    expect(exercisedProfiles).toEqual(new Set(Object.keys(PROFILES)));
    expect(exercisedPlacements).toEqual(new Set(Object.keys(PLACEMENTS)));
  });

  test("removes malformed block endpoints and consolidates only changed inline content", () => {
    const material = {
      type: "paragraph",
      content: [
        { type: "run", content: [{ type: "text", text: "a" }] },
        { type: "moveFromRangeStart", id: 8, name: "orphan" },
        { type: "run", content: [{ type: "text", text: "b" }] },
      ],
    } satisfies Paragraph;
    const survivor = { type: "paragraph", content: [] } satisfies Paragraph;
    const finalSectionProperties = getDefaultSectionProperties();
    const sectionEnd = {
      type: "paragraph",
      content: [],
      sectionProperties: getDefaultSectionProperties(),
    } satisfies Paragraph;
    const content: BlockContent[] = [
      { type: "moveFromRangeEnd", id: 8 },
      material,
      sectionEnd,
      { type: "moveToRangeEnd", id: 9 },
      survivor,
    ];
    const sections = buildSections(content, finalSectionProperties);
    const firstSection = sections.at(0);
    const lastSection = sections.at(1);
    if (!firstSection || !lastSection) {
      panic("Expected two source sections");
    }
    const headers = new Map([
      [
        "default" as const,
        {
          type: "header",
          hdrFtrType: "default",
          content: [],
        } satisfies HeaderFooter,
      ],
    ]);
    firstSection.headers = headers;
    const originalProperties = firstSection.properties;
    const documentBody = {
      content,
      sections,
      finalSectionProperties,
    } satisfies DocumentBody;

    expect(normalizeTrackedMoveRanges({ documentBody })).toEqual({
      removedUnbalancedMoveRangeMarkers: 3,
    });
    expect(documentBody.content).toEqual([material, sectionEnd, survivor]);
    expect(material.content).toEqual([
      { type: "run", content: [{ type: "text", text: "ab" }] },
    ]);
    expect(documentBody.sections).toBe(sections);
    expect(documentBody.sections.at(0)).toBe(firstSection);
    expect(documentBody.sections.at(1)).toBe(lastSection);
    expect(firstSection.properties).toBe(originalProperties);
    expect(firstSection.headers).toBe(headers);
    expect(firstSection.content).toEqual([material, sectionEnd]);
    expect(lastSection.content).toEqual([survivor]);
    expect(documentBody.content.at(-1)).toBe(survivor);
    expect(normalizeTrackedMoveRanges({ documentBody })).toEqual({
      removedUnbalancedMoveRangeMarkers: 0,
    });
  });

  test("drops only duplicate starts across ownership kinds and preserves balanced occurrences", () => {
    const sharedStart = {
      type: "moveToRangeStart",
      id: 7,
      name: "range",
    } as const;
    const sharedEnd = { type: "moveToRangeEnd", id: 7 } as const;
    const middle = {
      type: "paragraph",
      content: [
        { type: "run", content: [{ type: "text", text: "a" }] },
        { type: "moveToRangeStart", id: 7, name: "duplicate" },
        { type: "run", content: [{ type: "text", text: "b" }] },
      ],
    } satisfies Paragraph;
    const documentBody = {
      content: [sharedStart, middle, sharedEnd, sharedStart, sharedEnd],
    } satisfies DocumentBody;
    const sections = buildSections(documentBody.content, undefined);
    const withSections = { ...documentBody, sections };

    expect(normalizeTrackedMoveRanges({ documentBody: withSections })).toEqual({
      removedUnbalancedMoveRangeMarkers: 1,
    });
    expect(withSections.content).toEqual([
      sharedStart,
      middle,
      sharedEnd,
      sharedStart,
      sharedEnd,
    ]);
    expect(middle.content).toEqual([
      { type: "run", content: [{ type: "text", text: "ab" }] },
    ]);
    expect(markers(withSections).map(({ marker }) => marker)).toEqual([
      sharedStart,
      sharedEnd,
      sharedStart,
      sharedEnd,
    ]);
    expect(normalizeTrackedMoveRanges({ documentBody: withSections })).toEqual({
      removedUnbalancedMoveRangeMarkers: 0,
    });
    const header = {
      type: "header",
      hdrFtrType: "default",
      content: [sharedStart, sharedEnd],
    } satisfies HeaderFooter;
    const headers = new Map([["header", header]]);
    const occurrences: string[] = [];
    visitDocxContentSlots(
      { documentBody: withSections, headers },
      ({ item }) => {
        if (
          item.type === "moveToRangeStart" ||
          item.type === "moveToRangeEnd"
        ) {
          occurrences.push(item.type);
        }
      },
    );
    expect(occurrences).toEqual([
      "moveToRangeStart",
      "moveToRangeEnd",
      "moveToRangeStart",
      "moveToRangeEnd",
      "moveToRangeStart",
      "moveToRangeEnd",
    ]);
    expect(
      normalizeTrackedMoveRanges({ documentBody: withSections, headers }),
    ).toEqual({
      removedUnbalancedMoveRangeMarkers: 0,
    });
  });

  test("leaves balanced same-owner pairs and their arrays untouched", () => {
    const source = {
      type: "paragraph",
      content: [
        { type: "moveFromRangeStart", id: 1, name: "source" },
        { type: "run", content: [{ type: "text", text: "source" }] },
        { type: "moveFromRangeEnd", id: 1 },
      ],
    } satisfies Paragraph;
    const target = {
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text: "target" }] }],
    } satisfies Paragraph;
    const content = [
      source,
      { type: "moveToRangeStart", id: 2, name: "target" },
      target,
      { type: "moveToRangeEnd", id: 2 },
    ] satisfies BlockContent[];
    const sourceContent = source.content;
    const targetContent = target.content;
    const before = structuredClone(content);
    const documentBody = { content };
    expect(normalizeTrackedMoveRanges({ documentBody })).toEqual({
      removedUnbalancedMoveRangeMarkers: 0,
    });
    expect(content).toEqual(before);
    expect(documentBody.content).toBe(content);
    expect(source.content).toBe(sourceContent);
    expect(target.content).toBe(targetContent);
  });

  test("retains readonly record traversal order and prune boundaries", () => {
    const boxed = {
      type: "paragraph",
      paraId: "00000002",
      content: [],
    } satisfies Paragraph;
    const cell = {
      type: "paragraph",
      paraId: "00000003",
      content: [],
    } satisfies Paragraph;
    const controlled = {
      type: "paragraph",
      paraId: "00000004",
      content: [],
    } satisfies Paragraph;
    const codeBox = {
      type: "paragraph",
      paraId: "00000005",
      content: [],
    } satisfies Paragraph;
    const resultBox = {
      type: "paragraph",
      paraId: "00000006",
      content: [],
    } satisfies Paragraph;
    const directRun = shapeRun(boxed);
    const codeRun = shapeRun(codeBox);
    const resultRun = shapeRun(resultBox);
    const blocks = [
      {
        type: "paragraph",
        paraId: "00000001",
        content: [
          directRun,
          {
            type: "complexField",
            instruction: "REF target",
            fieldType: "REF",
            fieldCode: [codeRun],
            fieldResult: [resultRun],
          },
        ],
      },
      {
        type: "table",
        rows: [
          { type: "tableRow", cells: [{ type: "tableCell", content: [cell] }] },
        ],
      },
      {
        type: "blockSdt",
        properties: { sdtType: "richText" },
        content: [controlled],
      },
      { type: "moveToRangeEnd", id: 9 },
    ] satisfies readonly BlockContent[];
    const cases = {
      none: [
        "00000001",
        "00000002",
        "00000005",
        "00000006",
        "table",
        "tableRow",
        "00000003",
        "00000004",
      ],
      paragraph: ["00000001", "table", "tableRow", "00000003", "00000004"],
      table: [
        "00000001",
        "00000002",
        "00000005",
        "00000006",
        "table",
        "00000004",
      ],
      tableRow: [
        "00000001",
        "00000002",
        "00000005",
        "00000006",
        "table",
        "tableRow",
        "00000004",
      ],
    } as const;
    const exercised = new Set<string>();
    for (const [pruned, expected] of Object.entries(cases)) {
      const found: string[] = [];
      visitBlockTreeRecords(blocks, (record) => {
        const name = record.type === "paragraph" ? record.paraId : record.type;
        if (!name) {
          panic("Missing traversal fixture paragraph identifier");
        }
        found.push(name);
        return record.type === pruned
          ? BLOCK_TREE_DESCENT.prune
          : BLOCK_TREE_DESCENT.descend;
      });
      expect(found).toEqual(expected);
      exercised.add(pruned);
    }
    expect(exercised).toEqual(new Set(Object.keys(cases)));
    const outer = blocks.at(0);
    if (outer?.type !== "paragraph") {
      panic("Missing traversal fixture outer paragraph");
    }
    const runs: Run[] = [];
    visitParagraphRuns(outer, (item) => runs.push(item));
    expect(runs).toEqual([directRun, codeRun, resultRun]);
    expect(runs.at(0)).toBe(directRun);
    expect(runs.at(1)).toBe(codeRun);
    expect(runs.at(2)).toBe(resultRun);
  });

  test("pairs in source order around shape runs and complex-field results", () => {
    const directStart = {
      type: "moveToRangeStart",
      id: 7,
      name: "direct",
    } as const;
    const fieldStart = {
      type: "moveFromRangeStart",
      id: 8,
      name: "field",
    } as const;
    const outer = {
      type: "paragraph",
      content: [
        shapeRun({ type: "paragraph", content: [directStart] }),
        { type: "moveToRangeEnd", id: 7 },
        {
          type: "complexField",
          instruction: "REF target",
          fieldType: "REF",
          fieldCode: [],
          fieldResult: [shapeRun({ type: "paragraph", content: [fieldStart] })],
        },
      ],
    } satisfies Paragraph;
    const documentBody = {
      content: [outer, { type: "moveFromRangeEnd", id: 8 }],
    } satisfies DocumentBody;
    expect(markers(documentBody).map(({ marker }) => marker.type)).toEqual([
      "moveToRangeStart",
      "moveToRangeEnd",
      "moveFromRangeStart",
      "moveFromRangeEnd",
    ]);
    const before = structuredClone(documentBody);
    expect(normalizeTrackedMoveRanges({ documentBody })).toEqual({
      removedUnbalancedMoveRangeMarkers: 0,
    });
    expect(documentBody).toEqual(before);
  });
});
