import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  Document,
  HeaderFooter,
  Paragraph,
  Section,
} from "../../model/document";
import { paragraphNumberingReference } from "../../model/paragraphNumbering";
import { serializeDocumentToDocx } from "../../serialize/docx";
import { ZipArchive } from "../../zip/archive";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  captureSectionView,
  captureSectionViewState,
  rebuildSections,
  replaceParagraphs,
  sectionsInStep,
} from "../blocks";
import { contractViolation } from "../contract";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  SECTION_BOUNDARY_POLICIES,
} from "../types";
import type { DocumentOp, SectionPropertiesState } from "../types";
import { documentArbitrary } from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(30_000));

const paragraph: Paragraph = {
  type: "paragraph",
  paraId: "00000001",
  content: [{ type: "run", content: [{ type: "text", text: "item" }] }],
};

const base: Document = { package: { document: { content: [paragraph] } } };

test("ordinary paragraph splices reuse the checked trailing section", () => {
  let untouchedBoundaryReads = 0;
  const untouched: Paragraph = {
    ...paragraph,
    get sectionProperties() {
      untouchedBoundaryReads += 1;
      return undefined;
    },
  };
  const editable = { ...paragraph, paraId: "00000002" };
  const content = [untouched, editable];
  const properties = {};
  const document = {
    package: {
      document: {
        content,
        finalSectionProperties: properties,
        sections: [{ content, properties }],
      },
    },
  } satisfies Document;
  const splice = (current: Document, replacement: readonly Paragraph[]) => {
    const result = replaceParagraphs({
      document: current,
      story: OP_STORIES.MAIN,
      at: { list: [], index: 1 },
      count: 1,
      replacement,
    });
    if (result.isErr()) throw result.error;
    return result.value;
  };
  const checked = splice(document, [editable]);
  expect(untouchedBoundaryReads).toBeGreaterThan(0);
  const section = checked.package.document.sections?.at(0);
  if (!section) throw new Error("The splice must retain the trailing section");

  untouchedBoundaryReads = 0;
  expect(splice(checked, [editable]).package.document.sections?.at(0)).toBe(
    section,
  );
  const edited = splice(checked, [
    editable,
    { ...paragraph, paraId: "00000003" },
  ]).package.document;
  expect(untouchedBoundaryReads).toBe(0);

  const rebuilt = rebuildSections({
    document: checked,
    content: edited.content,
    previous: [section],
  });
  if (rebuilt.isErr()) throw rebuilt.error;
  expect(edited.sections).toStrictEqual(rebuilt.value);
  expect(sectionsInStep(edited)).toBe(true);
});

test("numbering instance creation and its inverse preserve absent numbering and exact ids", () => {
  fc.assert(
    fc.property(
      fc.constantFrom(
        base,
        { ...base, package: { ...base.package, numbering: undefined } },
        {
          ...base,
          package: {
            ...base.package,
            numbering: {
              abstractNums: [
                {
                  abstractNumId: 7,
                  levels: [
                    { ilvl: 0, numFmt: "decimal" as const, lvlText: "%1." },
                  ],
                },
              ],
              nums: [],
            },
          },
        },
      ),
      fc.integer({ min: 1, max: 2_000_000_000 }),
      fc.integer({ min: 1, max: 2_000_000_000 }),
      (document, abstractNumId, numId) => {
        fc.pre(abstractNumId !== numId);
        const existingAbstract = document.package.numbering?.abstractNums.at(0);
        const abstractNumIdToUse =
          existingAbstract?.abstractNumId ?? abstractNumId;
        fc.pre(abstractNumIdToUse !== numId);
        const abstractNum =
          existingAbstract === undefined
            ? {
                abstractNumId: abstractNumIdToUse,
                levels: [
                  { ilvl: 0, numFmt: "decimal" as const, lvlText: "%1." },
                ],
              }
            : undefined;
        const num = { numId, abstractNumId: abstractNumIdToUse };
        const forward = applyDocumentOp(document, {
          type: DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
          ...(abstractNum === undefined ? {} : { abstractNum }),
          num,
        });
        expect(forward.isOk()).toBe(true);
        if (forward.isErr()) return;
        expect(forward.value.document.package.numbering?.nums).toEqual([num]);
        const undo = applyDocumentOps(
          forward.value.document,
          forward.value.inverse,
        );
        expect(undo.isOk()).toBe(true);
        if (undo.isErr()) return;
        expect(undo.value.document).toStrictEqual(document);
        const redo = applyDocumentOps(undo.value.document, undo.value.inverse);
        expect(redo.isOk()).toBe(true);
        if (redo.isErr()) return;
        expect(redo.value.document).toStrictEqual(forward.value.document);
      },
    ),
    propertyConfig(),
  );
});

test("generated section endpoint edits restore paragraph, final, and section-view state", () => {
  fc.assert(
    fc.property(
      documentArbitrary,
      fc.integer({ min: 10_000, max: 20_000 }),
      (document, width) => {
        const topLevel = document.package.document.content.filter(
          (block): block is Paragraph => block.type === "paragraph",
        );
        const target = topLevel.at(0);
        if (target?.paraId === undefined) return;
        const endpoint = {
          type: "paragraph" as const,
          blockId: target.paraId.toLowerCase(),
        };
        const current = target.sectionProperties;
        const properties = { ...current, pageWidth: width };
        const original = structuredClone(document);
        const paragraphEdit = applyDocumentOp(document, {
          type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
          endpoint,
          expected:
            current === undefined
              ? {
                  type: Object.hasOwn(target, "sectionProperties")
                    ? ("undefined" as const)
                    : ("omitted" as const),
                }
              : { type: "present" as const, value: current },
          properties: { type: "present", value: properties },
        });
        expect(paragraphEdit.isOk()).toBe(true);
        if (paragraphEdit.isErr()) return;
        const paragraphUndo = applyDocumentOps(
          paragraphEdit.value.document,
          paragraphEdit.value.inverse,
        );
        expect(paragraphUndo.isOk()).toBe(true);
        if (paragraphUndo.isErr()) return;
        expect(paragraphUndo.value.document).toStrictEqual(original);
        const paragraphRedo = applyDocumentOps(
          paragraphUndo.value.document,
          paragraphUndo.value.inverse,
        );
        expect(paragraphRedo.isOk()).toBe(true);
        if (paragraphRedo.isErr()) return;
        expect(paragraphRedo.value.document).toStrictEqual(
          paragraphEdit.value.document,
        );

        const finalProperties =
          document.package.document.finalSectionProperties;
        const finalEdit = applyDocumentOp(document, {
          type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
          endpoint: { type: "final" },
          expected:
            finalProperties === undefined
              ? {
                  type: Object.hasOwn(
                    document.package.document,
                    "finalSectionProperties",
                  )
                    ? ("undefined" as const)
                    : ("omitted" as const),
                }
              : { type: "present" as const, value: finalProperties },
          properties: {
            type: "present",
            value: { ...finalProperties, pageHeight: width + 1 },
          },
        });
        expect(finalEdit.isOk()).toBe(true);
        if (finalEdit.isErr()) return;
        const finalUndo = applyDocumentOps(
          finalEdit.value.document,
          finalEdit.value.inverse,
        );
        expect(finalUndo.isOk()).toBe(true);
        if (finalUndo.isErr()) return;
        expect(finalUndo.value.document).toStrictEqual(original);
        const finalRedo = applyDocumentOps(
          finalUndo.value.document,
          finalUndo.value.inverse,
        );
        expect(finalRedo.isOk()).toBe(true);
        if (finalRedo.isErr()) return;
        expect(finalRedo.value.document).toStrictEqual(
          finalEdit.value.document,
        );
      },
    ),
    propertyConfig({ numRuns: 100 }),
  );
});

test("section endpoint edits preserve exact inverse for explicit and absent properties", () => {
  for (const endpoint of [
    { type: "paragraph" as const, blockId: "00000001" },
    { type: "final" as const },
  ]) {
    const forward = applyDocumentOp(base, {
      type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
      endpoint,
      expected: { type: "omitted" },
      properties: {
        type: "present",
        value: { pageWidth: 12_240, marginTop: 720 },
      },
    });
    expect(forward.isOk()).toBe(true);
    if (forward.isErr()) continue;
    const undo = applyDocumentOps(
      forward.value.document,
      forward.value.inverse,
    );
    expect(undo.isOk()).toBe(true);
    if (undo.isErr()) continue;
    expect(undo.value.document).toStrictEqual(base);
    const redo = applyDocumentOps(undo.value.document, undo.value.inverse);
    expect(redo.isOk()).toBe(true);
    if (redo.isErr()) continue;
    expect(redo.value.document).toStrictEqual(forward.value.document);
  }
});

test("numbering deletion refuses live paragraph and shared abstract references", () => {
  const abstractNum = {
    abstractNumId: 1,
    levels: [{ ilvl: 0, numFmt: "decimal" as const, lvlText: "%1." }],
  };
  const num = { numId: 1, abstractNumId: 1 };
  const referenced = {
    ...base,
    package: {
      ...base.package,
      document: {
        content: [
          {
            ...paragraph,
            formatting: { numPr: paragraphNumberingReference({ numId: 1 }) },
          },
        ],
      },
      numbering: { abstractNums: [abstractNum], nums: [num] },
    },
  };
  const liveReference = applyDocumentOp(referenced, {
    type: DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE,
    num,
    abstractNum,
  });
  expect(liveReference.isErr()).toBe(true);
  if (liveReference.isErr()) expect(liveReference.error.reason).toBe("stale");

  const sharedAbstract = {
    ...base,
    package: {
      ...base.package,
      numbering: {
        abstractNums: [abstractNum],
        nums: [num, { numId: 2, abstractNumId: 1 }],
      },
    },
  };
  const stillUsed = applyDocumentOp(sharedAbstract, {
    type: DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE,
    num,
    abstractNum,
  });
  expect(stillUsed.isErr()).toBe(true);
  if (stillUsed.isErr()) expect(stillUsed.error.reason).toBe("stale");
});

test("numbering and section inverses restore byte-equivalent DOCX parts", async () => {
  const originalZip = await ZipArchive.read(
    await serializeDocumentToDocx(base),
  );
  const originalDocumentXml = await originalZip
    .entry("word/document.xml")
    ?.bytes();
  const operations = [
    {
      type: DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
      num: { numId: 8, abstractNumId: 8 },
      abstractNum: {
        abstractNumId: 8,
        levels: [{ ilvl: 0, numFmt: "decimal" as const, lvlText: "%1." }],
      },
    },
    {
      type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
      endpoint: { type: "paragraph" as const, blockId: "00000001" },
      properties: {
        type: "present",
        value: { pageWidth: 12_240, marginTop: 720 },
      },
    },
  ] as const;
  let current = base;
  const inverses: DocumentOp[][] = [];
  for (const operation of operations) {
    const applied = applyDocumentOp(current, operation);
    expect(applied.isOk()).toBe(true);
    if (applied.isErr()) return;
    current = applied.value.document;
    inverses.push([...applied.value.inverse]);
  }
  const restored = applyDocumentOps(current, inverses.toReversed().flat());
  expect(restored.isOk()).toBe(true);
  if (restored.isErr()) return;
  const restoredZip = await ZipArchive.read(
    await serializeDocumentToDocx(restored.value.document),
  );
  expect(await restoredZip.entry("word/document.xml")?.bytes()).toEqual(
    originalDocumentXml,
  );
  expect(restoredZip.entry("word/numbering.xml")).toBeUndefined();
});

const header: HeaderFooter = {
  type: "header",
  hdrFtrType: "default",
  content: [],
};
const footer: HeaderFooter = {
  type: "footer",
  hdrFtrType: "default",
  content: [],
};
const sectionPresence = ["omitted", "undefined", "empty", "populated"] as const;

test("generated section boundary changes preserve exact map presence through wire undo and redo", () => {
  fc.assert(
    fc.property(
      fc.constantFrom(...sectionPresence),
      fc.constantFrom(...sectionPresence),
      fc.integer({ min: 8_000, max: 20_000 }),
      (headerPresence, footerPresence, width) => {
        const boundary: Paragraph = {
          ...paragraph,
          sectionProperties: { pageWidth: width },
        };
        const last: Paragraph = { ...paragraph, paraId: "00000002" };
        const firstSection: Section = {
          properties: boundary.sectionProperties ?? {},
          content: [boundary],
        };
        const finalSection: Section = {
          properties: { pageHeight: width },
          content: [last],
        };
        const fill = (
          section: Section,
          key: "headers" | "footers",
          presence: typeof headerPresence,
        ) => {
          switch (presence) {
            case "omitted":
              return;
            case "undefined":
              section[key] = undefined;
              return;
            case "empty":
              section[key] = new Map();
              return;
            case "populated":
              section[key] = new Map([
                ["default", key === "headers" ? header : footer],
              ]);
              return;
            default: {
              const unreachable: never = presence;
              return unreachable;
            }
          }
        };
        fill(firstSection, "headers", headerPresence);
        fill(firstSection, "footers", footerPresence);
        fill(finalSection, "headers", footerPresence);
        fill(finalSection, "footers", headerPresence);
        const original: Document = {
          package: {
            document: {
              content: [boundary, last],
              finalSectionProperties: finalSection.properties,
              sections: [firstSection, finalSection],
            },
          },
        };
        const forward = applyDocumentOp(original, {
          type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
          story: "main",
          blockId: "00000001",
          nextBlockId: "00000002",
          sectionBoundary: SECTION_BOUNDARY_POLICIES.REMOVE,
        }).unwrap();
        expect(contractViolation(forward.document)).toBeUndefined();
        const inverse: DocumentOp[] = JSON.parse(
          JSON.stringify(forward.inverse),
        );
        const undo = applyDocumentOps(forward.document, inverse).unwrap();
        expect(undo.document).toStrictEqual(original);
        expect(
          captureSectionViewState(undo.document.package.document),
        ).toStrictEqual(captureSectionViewState(original.package.document));
        for (const [
          index,
          originalSection,
        ] of original.package.document.sections?.entries() ?? []) {
          const restored = undo.document.package.document.sections?.at(index);
          expect(Object.hasOwn(restored ?? {}, "headers")).toBe(
            Object.hasOwn(originalSection, "headers"),
          );
          expect(Object.hasOwn(restored ?? {}, "footers")).toBe(
            Object.hasOwn(originalSection, "footers"),
          );
        }
        const redo = applyDocumentOps(undo.document, undo.inverse).unwrap();
        expect(redo.document).toStrictEqual(forward.document);
      },
    ),
    propertyConfig({ numRuns: 40 }),
  );
});

test("section reconstruction refuses unresolved parts and missing canonical properties atomically", () => {
  const boundary: Paragraph = {
    ...paragraph,
    sectionProperties: { pageWidth: 10_000 },
  };
  const original: Document = {
    package: {
      document: {
        content: [boundary],
        sections: [
          { properties: boundary.sectionProperties ?? {}, content: [boundary] },
        ],
      },
    },
  };
  for (const properties of [
    undefined,
    {
      pageWidth: 11_000,
      headerReferences: [{ type: "default" as const, rId: "rIdMissing" }],
    },
  ]) {
    const replacement: Paragraph = {
      ...paragraph,
      ...(properties === undefined ? {} : { sectionProperties: properties }),
    };
    const result = applyDocumentOps(original, [
      {
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
        story: "main",
        blockId: "00000001",
        patch: { alignment: "end" },
      },
      {
        type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
        story: "main",
        expected: [{ ...boundary, formatting: { alignment: "end" } }],
        blocks: [replacement],
        sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE,
      },
    ]);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.reason).toBe("sectionBoundary");
    expect(original.package.document.content.at(0)).toBe(boundary);
    expect(original.package.document.sections?.at(0)?.properties).toBe(
      boundary.sectionProperties,
    );
  }
});

test("same-count replacements rebuild changed section properties and stale snapshots refuse", () => {
  const boundary: Paragraph = {
    ...paragraph,
    sectionProperties: { pageWidth: 10_000 },
  };
  const original: Document = {
    package: {
      document: {
        content: [boundary],
        sections: [
          {
            properties: boundary.sectionProperties ?? {},
            content: [boundary],
            headers: undefined,
          },
        ],
      },
    },
  };
  const changed: Paragraph = {
    ...boundary,
    sectionProperties: { pageWidth: 11_000 },
  };
  const forward = applyDocumentOp(original, {
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: "main",
    expected: [boundary],
    blocks: [changed],
    sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE,
  }).unwrap();
  expect(forward.document.package.document.sections?.at(0)?.properties).toEqual(
    changed.sectionProperties,
  );
  expect(contractViolation(forward.document)).toBeUndefined();
  const undo = applyDocumentOps(forward.document, forward.inverse).unwrap();
  expect(captureSectionViewState(undo.document.package.document)).toStrictEqual(
    captureSectionViewState(original.package.document),
  );
  const invalid = applyDocumentOp(original, {
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: "main",
    expected: [boundary],
    blocks: [changed],
    sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE,
    sectionView: {
      expected: captureSectionView(original.package.document.sections ?? []),
      restore: [],
    },
  });
  expect(invalid.isErr()).toBe(true);
  if (invalid.isErr()) expect(invalid.error.reason).toBe("stale");
  expect(original.package.document.content.at(0)).toBe(boundary);
});

test("metadata-only endpoint updates and view creation have exact inverses", () => {
  for (const view of ["omitted", "undefined", "sections"] as const) {
    const original = structuredClone(base);
    const body = original.package.document;
    if (view === "undefined") body.sections = undefined;
    if (view === "sections")
      body.sections = [
        { properties: {}, content: body.content, headers: undefined },
      ];
    const before = captureSectionViewState(body);
    const forward = applyDocumentOp(original, {
      type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
      endpoint: { type: "final" },
      properties: { type: "omitted" },
      expectedSectionMetadata: before,
      sectionMetadata: {
        type: "sections",
        value: [
          {
            properties: {},
            headers: { type: "entries", value: [["default", header]] },
            footers: { type: "undefined" },
          },
        ],
      },
    }).unwrap();
    expect(forward.inverse.length).toBe(1);
    expect(
      forward.document.package.document.sections
        ?.at(0)
        ?.headers?.get("default"),
    ).toEqual(header);
    const undo = applyDocumentOps(
      forward.document,
      JSON.parse(JSON.stringify(forward.inverse)),
    ).unwrap();
    expect(undo.document).toStrictEqual(original);
    expect(Object.hasOwn(undo.document.package.document, "sections")).toBe(
      Object.hasOwn(body, "sections"),
    );
    expect(
      captureSectionViewState(undo.document.package.document),
    ).toStrictEqual(before);
    expect(
      applyDocumentOps(undo.document, undo.inverse).unwrap().document,
    ).toStrictEqual(forward.document);
    const stale = applyDocumentOp(original, {
      type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
      endpoint: { type: "final" },
      properties: { type: "omitted" },
      expectedSectionMetadata: { type: "sections", value: [] },
    });
    expect(stale.isErr()).toBe(true);
    if (stale.isErr()) expect(stale.error.reason).toBe("stale");
  }
});

test("endpoint property states preserve exact omitted and undefined keys through wire history", () => {
  fc.assert(
    fc.property(fc.integer({ min: 8_000, max: 20_000 }), (width) => {
      for (const endpoint of [
        { type: "paragraph", blockId: "00000001" },
        { type: "final" },
      ] as const) {
        for (const beforeType of ["omitted", "undefined", "present"] as const) {
          for (const afterType of [
            "omitted",
            "undefined",
            "present",
          ] as const) {
            const original = structuredClone(base);
            const body = original.package.document;
            const first = body.content.at(0);
            if (first?.type !== "paragraph")
              throw new TypeError("Missing fixture paragraph.");
            const beforeValue = { pageWidth: width };
            switch (beforeType) {
              case "omitted":
                break;
              case "undefined":
                if (endpoint.type === "final")
                  body.finalSectionProperties = undefined;
                else first.sectionProperties = undefined;
                break;
              case "present":
                if (endpoint.type === "final")
                  body.finalSectionProperties = beforeValue;
                else first.sectionProperties = beforeValue;
                break;
              default: {
                const unreachable: never = beforeType;
                return unreachable;
              }
            }
            const before = (
              beforeType === "present"
                ? ({ type: beforeType, value: beforeValue } as const)
                : ({ type: beforeType } as const)
            ) satisfies SectionPropertiesState;
            const after = (
              afterType === "present"
                ? ({
                    type: afterType,
                    value: { pageWidth: width + 1 },
                  } as const)
                : ({ type: afterType } as const)
            ) satisfies SectionPropertiesState;
            const forward = applyDocumentOp(original, {
              type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
              endpoint,
              expected: before,
              properties: after,
            }).unwrap();
            const nextBody = forward.document.package.document;
            const nextParagraph = nextBody.content.at(0);
            const nextRecord =
              endpoint.type === "final" ? nextBody : nextParagraph;
            const key =
              endpoint.type === "final"
                ? "finalSectionProperties"
                : "sectionProperties";
            expect(Object.hasOwn(nextRecord ?? {}, key)).toBe(
              afterType !== "omitted",
            );
            const undo = applyDocumentOps(
              forward.document,
              JSON.parse(JSON.stringify(forward.inverse)),
            ).unwrap();
            expect(undo.document).toStrictEqual(original);
            const undoRecord =
              endpoint.type === "final"
                ? undo.document.package.document
                : undo.document.package.document.content.at(0);
            expect(Object.hasOwn(undoRecord ?? {}, key)).toBe(
              beforeType !== "omitted",
            );
            expect(
              applyDocumentOps(undo.document, undo.inverse).unwrap().document,
            ).toStrictEqual(forward.document);
            if (beforeType !== "present") {
              const stale = applyDocumentOp(original, {
                type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
                endpoint,
                expected: {
                  type: beforeType === "omitted" ? "undefined" : "omitted",
                },
                properties: after,
              });
              expect(stale.isErr()).toBe(true);
              if (stale.isErr()) expect(stale.error.reason).toBe("stale");
            }
          }
        }
      }
    }),
    propertyConfig({ numRuns: 10 }),
  );
});

/** Numbering with instances and definitions in a sampled order. */
const unorderedNumbering = fc
  .uniqueArray(fc.integer({ min: 1, max: 40 }), { minLength: 1, maxLength: 5 })
  .map((ids) => ({
    abstractNums: ids.map((abstractNumId) => ({
      abstractNumId,
      levels: [{ ilvl: 0, numFmt: "decimal" as const, lvlText: "%1." }],
    })),
    nums: ids.map((id) => ({ numId: id, abstractNumId: id })),
  }));

const numberedDocument = (
  numbering: Document["package"]["numbering"],
): Document => ({ ...base, package: { ...base.package, numbering } });

const creation = (id: number): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
  num: { numId: id, abstractNumId: id },
  abstractNum: {
    abstractNumId: id,
    levels: [{ ilvl: 0, numFmt: "decimal" as const, lvlText: "%1." }],
  },
});

/**
 * Two creations applied in either order leave one numbering part;
 * `appending` is the test-only break, which moves each created entry to the
 * end of its list as an appending creation would.
 */
const concurrentCreations = (appending: boolean) =>
  fc.property(
    unorderedNumbering,
    fc.uniqueArray(fc.integer({ min: 41, max: 80 }), {
      minLength: 2,
      maxLength: 2,
    }),
    (numbering, [first = 41, second = 42]) => {
      const create = (document: Document, id: number): Document => {
        const applied = applyDocumentOp(document, creation(id));
        if (applied.isErr()) throw applied.error;
        const created = applied.value.document;
        const part = created.package.numbering;
        if (!appending || part === undefined) return created;
        const last = <Entry>(
          list: readonly Entry[],
          is: (each: Entry) => boolean,
        ) => list.filter((each) => !is(each)).concat(list.filter(is));
        return numberedDocument({
          ...part,
          nums: last(part.nums, ({ numId }) => numId === id),
          abstractNums: last(
            part.abstractNums,
            ({ abstractNumId }) => abstractNumId === id,
          ),
        });
      };
      const start = numberedDocument(numbering);
      expect(create(create(start, first), second).package.numbering).toEqual(
        create(create(start, second), first).package.numbering,
      );
    },
  );

test(
  "concurrent numbering creations leave one order whichever applies first",
  () => {
    fc.assert(concurrentCreations(false), propertyConfig());
  },
  propertyTestTimeout(30_000),
);

test(
  "appending numbering creations are caught",
  () => {
    const result = fc.check(
      concurrentCreations(true),
      // The break is a counterexample on purpose: report nothing.
      propertyConfig({ numRuns: 1, reporter: () => undefined }),
    );
    expect(result.failed).toBe(true);
  },
  propertyTestTimeout(30_000),
);

test("undoing the deletion of a numbering instance puts it back where it was", () => {
  fc.assert(
    fc.property(unorderedNumbering, fc.nat(), (numbering, pick) => {
      const removed = numbering.nums[pick % numbering.nums.length];
      const abstractNum = numbering.abstractNums.find(
        ({ abstractNumId }) => abstractNumId === removed?.abstractNumId,
      );
      if (removed === undefined || abstractNum === undefined)
        throw new Error("The sampled instance is there.");
      const start = numberedDocument(numbering);
      const deleted = applyDocumentOp(start, {
        type: DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE,
        num: removed,
        abstractNum,
      });
      if (deleted.isErr()) throw deleted.error;
      const undone = applyDocumentOps(
        deleted.value.document,
        deleted.value.inverse,
      );
      if (undone.isErr()) throw undone.error;
      expect(undone.value.document).toStrictEqual(start);
    }),
    propertyConfig(),
  );
});
