import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type {
  Document,
  HeaderFooter,
  HeaderFooterType,
  Paragraph,
  Section,
  Watermark,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { DOCUMENT_OP_TYPES, type DocumentOp, type SetDocumentWatermarkOp } from "../types";
import {
  ensureDocumentWatermarkHeaderCoverage,
  planDocumentWatermarkCoverage,
  planDocumentWatermarkHosts,
} from "../watermark";

setDefaultTimeout(propertyTestTimeout(30_000));
const paragraph = (paraId: string, text = ""): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
  preservedAttributes: [{ namespace: "urn:source", name: "stamp", value: "authored" }],
});
const fixture = (titlePage: boolean, evenPages: boolean): Document => ({
  package: {
    document: {
      content: [
        { ...paragraph("00000001", "First"), sectionProperties: { titlePg: titlePage } },
        paragraph("00000002", "Last"),
      ],
      finalSectionProperties: {
        titlePg: titlePage,
        headerReferences: [{ type: "default", rId: "rIdSource" }],
      },
    },
    headers: new Map([
      [
        "rIdSource",
        {
          type: "header",
          hdrFtrType: "default",
          content: [
            paragraph("00000003", "Before"),
            paragraph("00000004", " "),
            paragraph("00000005", "After"),
          ],
          watermark: { kind: "text", text: "Source" },
          watermarkBlockIndex: 1,
          rawWatermarkXml: '<w:p><v:shape id="authored"/></w:p>',
          verbatimXml: '<w:hdr xmlns:w="urn:fixture"/>',
          verbatimFingerprint: "captured",
        },
      ],
    ]),
    footers: new Map([
      [
        "rIdFooter",
        { type: "footer", hdrFtrType: "default", content: [paragraph("00000006", "Foreign")] },
      ],
    ]),
    settings: { defaultTabStop: 720, evenAndOddHeaders: evenPages },
    relationships: new Map([
      [
        "rIdForeign",
        { id: "rIdForeign", type: "urn:foreign", target: "opaque.bin", targetMode: "Internal" },
      ],
    ]),
  },
});
type OperationOptions = { document: Document; watermark: Watermark; serial: number };
const operation = ({ document, watermark, serial }: OperationOptions): SetDocumentWatermarkOp => ({
  type: DOCUMENT_OP_TYPES.SET_DOCUMENT_WATERMARK,
  change: { kind: "set", watermark },
  hosts: planDocumentWatermarkHosts(document).map((rId, index) => ({
    rId,
    paraId: (0x3000 + serial * 32 + index).toString(16).padStart(8, "0"),
  })),
  coverage: planDocumentWatermarkCoverage(document).map((type, index) => ({
    type,
    rId: `rIdCoverage${serial}_${type}`,
    paraId: (0x1000 + serial * 4 + index).toString(16).padStart(8, "0"),
  })),
});

test("generated watermark histories retain raw source, placement, coverage and exact undo/redo", () => {
  assertProperty(
    fc.property(
      fc.boolean(),
      fc.boolean(),
      fc.boolean(),
      fc.array(
        fc.oneof(
          fc
            .string({ minLength: 1, maxLength: 12 })
            .map((text): Watermark => ({ kind: "text", text, diagonal: false, opacity: 0.4 })),
          fc
            .record({ scale: fc.double({ min: 0.1, max: 2, noNaN: true }), washout: fc.boolean() })
            .map(
              ({ scale, washout }): Watermark => ({
                kind: "picture",
                imageRId: "rIdWatermarkImage",
                imageTarget: "media/watermark.png",
                scale,
                washout,
              }),
            ),
        ),
        { minLength: 1, maxLength: 5 },
      ),
      (titlePage, evenPages, hostless, watermarks) => {
        const original = fixture(titlePage, evenPages);
        const sourceHeader = original.package.headers?.get("rIdSource");
        if (hostless && sourceHeader) delete sourceHeader.watermarkBlockIndex;
        let current = original;
        const inverse: (readonly DocumentOp[])[] = [];
        for (const [index, watermark] of watermarks.entries()) {
          const op = operation({ document: current, watermark, serial: index });
          const applied = applyDocumentOp(current, op).unwrap();
          expect(applied.document.package.relationships).toBe(original.package.relationships);
          expect(applied.document.package.footers).toBe(original.package.footers);
          expect(planDocumentWatermarkCoverage(applied.document)).toEqual([]);
          expect(applied.document.package.headers?.get("rIdSource")?.watermarkBlockIndex).toBe(
            hostless ? 0 : 1,
          );
          expect(
            applied.document.package.headers
              ?.get("rIdSource")
              ?.content.map((block) => (block.type === "paragraph" ? block.paraId : "")),
          ).toEqual(
            hostless
              ? ["00003000", "00000003", "00000004", "00000005"]
              : ["00000003", "00000004", "00000005"],
          );
          expect(
            applied.document.package.headers?.get("rIdSource")?.rawWatermarkXml,
          ).toBeUndefined();
          inverse.unshift(applied.inverse);
          current = applied.document;
        }
        const removed = applyDocumentOp(current, {
          type: DOCUMENT_OP_TYPES.SET_DOCUMENT_WATERMARK,
          change: { kind: "remove" },
          coverage: [],
          hosts: [],
        }).unwrap();
        inverse.unshift(removed.inverse);
        current = removed.document;
        for (const header of current.package.headers?.values() ?? []) {
          expect(header.watermark).toBeUndefined();
          expect(header.watermarkBlockIndex).toBeUndefined();
        }
        const final = current;
        const redo: (readonly DocumentOp[])[] = [];
        for (const step of inverse) {
          const applied = applyDocumentOps(current, step).unwrap();
          redo.unshift(applied.inverse);
          current = applied.document;
        }
        expect(current).toStrictEqual(original);
        for (const step of redo) current = applyDocumentOps(current, step).unwrap().document;
        expect(current).toStrictEqual(final);
      },
    ),
    { seed: 20261014, numRuns: 40 },
  );
});

test("watermark coverage only adds missing forward variants and leaves authored VML intact", () => {
  const original = fixture(true, true);
  const types = planDocumentWatermarkCoverage(original);
  expect(types).toEqual(["default", "first", "even"]);
  const extended = ensureDocumentWatermarkHeaderCoverage({
    document: original,
    authority: "legacy",
    watermark: { kind: "text", text: "New" },
    coverage: types.map((type, index) => ({
      type,
      rId: `rIdNew${type}`,
      content: [paragraph((0x2000 + index).toString(16).padStart(8, "0"))],
    })),
  }).unwrap();
  expect(extended.package.headers?.get("rIdSource")).toBe(
    original.package.headers?.get("rIdSource"),
  );
  const first = extended.package.document.content.at(0);
  expect(first?.type === "paragraph" && first.sectionProperties?.headerReferences).toHaveLength(3);
  expect(planDocumentWatermarkCoverage(extended)).toEqual([]);
  const invalid = operation({
    document: original,
    watermark: { kind: "text", text: "New" },
    serial: 0,
  });
  expect(
    applyDocumentOp(original, {
      ...invalid,
      coverage: invalid.coverage.map((entry) => Object.assign({}, entry, { rId: "rIdForeign" })),
    }).isErr(),
  ).toBe(true);
  expect(
    applyDocumentOp(original, {
      ...invalid,
      coverage: invalid.coverage.map((entry) => Object.assign({}, entry, { paraId: "00000001" })),
    }).isErr(),
  ).toBe(true);
});

test("canonical picture watermark normalization omits internal target mode and resolves uniform dimensions", () => {
  const targetModes: readonly {
    watermark: Watermark;
    expectedExternal: boolean | undefined;
  }[] = [
    {
      watermark: {
        kind: "picture",
        imageRId: "rIdWatermarkImage",
        imageTarget: "media/watermark.png",
        widthPt: 249,
        heightPt: 124.2,
      },
      expectedExternal: undefined,
    },
    {
      watermark: {
        kind: "picture",
        imageRId: "rIdWatermarkImage",
        imageTarget: "media/watermark.png",
        imageTargetExternal: false,
        widthPt: 249,
        heightPt: 124.2,
      },
      expectedExternal: undefined,
    },
    {
      watermark: {
        kind: "picture",
        imageRId: "rIdWatermarkImage",
        imageTarget: "https://example.test/watermark.png",
        imageTargetExternal: true,
        widthPt: 249,
        heightPt: 124.2,
      },
      expectedExternal: true,
    },
  ];

  for (const [index, { watermark, expectedExternal }] of targetModes.entries()) {
    const original = fixture(true, false);
    const applied = applyDocumentOp(
      original,
      operation({ document: original, watermark, serial: index }),
    ).unwrap();
    const normalized = applied.document.package.headers?.get("rIdSource")?.watermark;

    expect(normalized).toMatchObject({
      kind: "picture",
      imageTargetExternal: expectedExternal,
      widthPt: 249,
      heightPt: 124.2,
      scale: 0.6,
    });
    if (expectedExternal === undefined) {
      expect(normalized && "imageTargetExternal" in normalized).toBe(false);
    }
  }

  const dimensions = [
    {
      watermark: {
        kind: "picture",
        imageRId: "rIdWatermarkImage",
        widthPt: 415,
        heightPt: 208,
      },
      expected: { widthPt: 415, heightPt: 207, scale: 1 },
    },
    {
      watermark: {
        kind: "picture",
        imageRId: "rIdWatermarkImage",
        widthPt: 300,
        heightPt: 123,
      },
      expected: { widthPt: 300, heightPt: 123 },
    },
    {
      watermark: {
        kind: "picture",
        imageRId: "rIdWatermarkImage",
        scale: 0.6,
        widthPt: 300,
        heightPt: 123,
      },
      expected: { widthPt: 249, heightPt: 124.2, scale: 0.6 },
    },
  ] satisfies readonly { watermark: Watermark; expected: Record<string, number> }[];

  for (const [index, { watermark, expected }] of dimensions.entries()) {
    const original = fixture(true, false);
    const applied = applyDocumentOp(
      original,
      operation({ document: original, watermark, serial: targetModes.length + index }),
    ).unwrap();
    const normalized = applied.document.package.headers?.get("rIdSource")?.watermark;

    expect(normalized).toMatchObject({ kind: "picture", ...expected });
    expect(normalized && "scale" in normalized).toBe("scale" in expected);
  }
});

test("generated malformed watermark payloads refuse without model mutation", () => {
  assertProperty(
    fc.property(
      fc.constantFrom(
        -1,
        0,
        Number.MIN_VALUE,
        1e308,
        Number.NaN,
        Number.POSITIVE_INFINITY,
        Number.NEGATIVE_INFINITY,
      ),
      fc.constantFrom(-1, 2, Number.NaN, Number.POSITIVE_INFINITY),
      fc.constantFrom("\u0000", "\u0001", "\ud800"),
      (dimension, opacity, illegal) => {
        const original = fixture(true, true);
        const snapshot = structuredClone(original);
        const invalidText: readonly (readonly [string, unknown])[] = [
          ["text", ""],
          ["text", illegal],
          ["text", 3],
          ["font", 3],
          ["font", illegal],
          ["font", "Calibri;rotation:0"],
          ["color", 'red" stroked="t'],
          ["color", 3],
          ["diagonal", "false"],
          ["opacity", opacity],
        ];
        for (const [field, value] of invalidText) {
          const op = operation({
            document: original,
            watermark: { kind: "text", text: "Valid" },
            serial: 0,
          });
          if (op.change.kind !== "set") throw new Error("Expected set fixture");
          Reflect.set(op.change.watermark, field, value);
          expect(applyDocumentOp(original, op).isErr()).toBe(true);
          expect(original).toStrictEqual(snapshot);
        }
        const invalidPicture: readonly (readonly [string, unknown])[] = [
          ["widthPt", dimension],
          ["heightPt", dimension],
          ["scale", dimension],
          ["widthPt", "415"],
          ["imageRId", 3],
          ["imageRId", illegal],
          ["imageTarget", undefined],
          ["imageTarget", 3],
          ["imageTarget", illegal],
          ["imageTargetExternal", "true"],
          ["washout", "false"],
        ];
        for (const [field, value] of invalidPicture) {
          const op = operation({
            document: original,
            watermark: {
              kind: "picture",
              imageRId: "rIdImage",
              imageTarget: "word/media/image.png",
            },
            serial: 0,
          });
          if (op.change.kind !== "set") throw new Error("Expected set fixture");
          Reflect.set(op.change.watermark, field, value);
          expect(applyDocumentOp(original, op).isErr()).toBe(true);
          expect(original).toStrictEqual(snapshot);
        }
        for (const field of ["coverage", "hosts"]) {
          for (const value of [null, [null], [3], [{}]]) {
            const op = operation({
              document: original,
              watermark: { kind: "text", text: "Valid" },
              serial: 0,
            });
            Reflect.set(op, field, value);
            expect(applyDocumentOp(original, op).isErr()).toBe(true);
            expect(original).toStrictEqual(snapshot);
          }
        }
      },
    ),
    { seed: 20261017, numRuns: 30 },
  );
});

test("generated annotated watermark hosts survive removal and exact inverse", () => {
  assertProperty(
    fc.property(fc.boolean(), fc.boolean(), (titlePage, evenPages) => {
      const original = fixture(titlePage, evenPages);
      const source = original.package.headers?.get("rIdSource");
      const host = source?.content.at(1);
      if (!source || host?.type !== "paragraph") throw new Error("Missing host fixture");
      host.content.push(
        { type: "bookmarkStart", id: 40, name: "Owned elsewhere" },
        { type: "bookmarkEnd", id: 40 },
      );
      const removed = applyDocumentOp(original, {
        type: DOCUMENT_OP_TYPES.SET_DOCUMENT_WATERMARK,
        change: { kind: "remove" },
        coverage: [],
        hosts: [],
      }).unwrap();
      expect(removed.document.package.headers?.get("rIdSource")?.content).toStrictEqual(
        source.content,
      );
      const restored = applyDocumentOps(removed.document, removed.inverse).unwrap();
      expect(restored.document).toStrictEqual(original);
      const set = applyDocumentOp(
        original,
        operation({ document: original, watermark: { kind: "text", text: "Updated" }, serial: 0 }),
      ).unwrap();
      const content = set.document.package.headers?.get("rIdSource")?.content;
      expect(content?.slice(0, 1)).toStrictEqual(source.content.slice(0, 1));
      expect(content?.slice(2)).toStrictEqual(source.content.slice(1));
      expect(applyDocumentOps(set.document, set.inverse).unwrap().document).toStrictEqual(original);
    }),
    { seed: 20261018, numRuns: 20 },
  );
});

test("generated canonical watermark mounts resolve cloned references and inherit new coverage", () => {
  assertProperty(
    fc.property(
      fc.integer({ min: 2, max: 5 }),
      fc.boolean(),
      fc.boolean(),
      fc.boolean(),
      (sectionCount, hasReferences, titlePage, evenPages) => {
        const original = fixture(titlePage, evenPages);
        const source = original.package.headers?.get("rIdSource");
        if (!source) throw new Error("Missing source header fixture");
        const alternate = {
          type: "header",
          hdrFtrType: "default",
          content: [paragraph("70000020", "Alternate")],
        } satisfies HeaderFooter;
        original.package.headers?.set("rIdAlternate", alternate);
        const sections = Array.from({ length: sectionCount }, (_, index) => {
          const properties = {
            titlePg: titlePage,
            headerReferences:
              hasReferences && (index === 0 || index === sectionCount - 1)
                ? [{ type: "default", rId: index === 0 ? "rIdSource" : "rIdAlternate" } as const]
                : [],
          };
          const block = {
            ...paragraph((0x70000000 + index).toString(16).padStart(8, "0"), `Section ${index}`),
            ...(index < sectionCount - 1 ? { sectionProperties: properties } : {}),
          } satisfies Paragraph;
          return {
            properties,
            content: [block],
            // Mounted views need not share object identity with the authoritative part.
            headers: new Map<HeaderFooterType, HeaderFooter>([
              ["default", structuredClone(source)],
              ["even", structuredClone(alternate)],
            ]),
          } satisfies Section;
        });
        original.package.document = {
          content: sections.flatMap(({ content }) => content),
          finalSectionProperties: sections.at(-1)?.properties,
          sections,
        };
        const snapshot = structuredClone(original);
        const op = operation({
          document: original,
          watermark: { kind: "text", text: "Updated" },
          serial: 0,
        });
        const applied = applyDocumentOp(original, op).unwrap();
        const mountedSections = applied.document.package.document.sections;
        expect(mountedSections?.length).toBe(sectionCount);
        const inherited = new Map<HeaderFooterType, string>();
        for (const [index, section] of (mountedSections ?? []).entries()) {
          for (const reference of section.properties.headerReferences ?? [])
            inherited.set(reference.type, reference.rId);
          expect([...(section.headers?.keys() ?? [])].sort()).toEqual([...inherited.keys()].sort());
          for (const [type, rId] of inherited) {
            expect(section.headers?.get(type)).toBe(applied.document.package.headers?.get(rId));
            expect(section.headers?.get(type)?.watermark).toMatchObject({
              kind: "text",
              text: "Updated",
            });
          }
          let expectedDefault = "rIdCoverage0_default";
          if (hasReferences)
            expectedDefault = index === sectionCount - 1 ? "rIdAlternate" : "rIdSource";
          expect(section.headers?.get("default")).toBe(
            applied.document.package.headers?.get(expectedDefault),
          );
          expect(section.headers?.has("first")).toBe(titlePage);
          expect(section.headers?.has("even")).toBe(evenPages);
        }
        expect(original).toStrictEqual(snapshot);
        expect(applyDocumentOps(applied.document, applied.inverse).unwrap().document).toStrictEqual(
          snapshot,
        );
        const removed = applyDocumentOp(applied.document, {
          type: DOCUMENT_OP_TYPES.SET_DOCUMENT_WATERMARK,
          change: { kind: "remove" },
          coverage: [],
          hosts: [],
        }).unwrap();
        for (const section of removed.document.package.document.sections ?? []) {
          for (const header of section.headers?.values() ?? []) {
            expect(header.watermark).toBeUndefined();
            expect([...(removed.document.package.headers?.values() ?? [])]).toContain(header);
          }
        }
        const legacy = ensureDocumentWatermarkHeaderCoverage({
          document: original,
          authority: "legacy",
          watermark: { kind: "text", text: "Updated" },
          coverage: op.coverage.map(({ type, rId, paraId }) => ({
            type,
            rId,
            content: [paragraph(paraId)],
          })),
        }).unwrap();
        for (const [index, section] of (legacy.package.document.sections ?? []).entries())
          expect(section.headers).toBe(sections.at(index)?.headers);
      },
    ),
    { seed: 20261019, numRuns: 40 },
  );
});

test("generated lowercase watermark colors normalize to the same canonical model", () => {
  assertProperty(
    fc.property(fc.integer({ min: 0, max: 0xffffff }), (value) => {
      const original = fixture(true, true);
      const color = value.toString(16).padStart(6, "0");
      const lower = applyDocumentOp(
        original,
        operation({
          document: original,
          watermark: { kind: "text", text: "Color", color },
          serial: 0,
        }),
      ).unwrap();
      const upper = applyDocumentOp(
        original,
        operation({
          document: original,
          watermark: { kind: "text", text: "Color", color: color.toUpperCase() },
          serial: 0,
        }),
      ).unwrap();
      expect(lower.document).toStrictEqual(upper.document);
      for (const header of lower.document.package.headers?.values() ?? [])
        expect(header.watermark).toMatchObject({ kind: "text", color: color.toUpperCase() });
      expect(applyDocumentOps(lower.document, lower.inverse).unwrap().document).toStrictEqual(
        original,
      );
    }),
    { seed: 20261020, numRuns: 40 },
  );
});
