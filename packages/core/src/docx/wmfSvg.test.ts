import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { renderEmfSvg } from "./metafileSvg";
import { parseDocx } from "./parser";
import { repackDocx } from "./rezip";
import { renderWmfSvg } from "./wmfSvg";

const FIXTURE = resolve(import.meta.dir, "__fixtures__/header-vml-emf.docx");

const META = {
  eof: 0x0000,
  saveDc: 0x001e,
  setBackgroundMode: 0x0102,
  setMapMode: 0x0103,
  setPolyFillMode: 0x0106,
  restoreDc: 0x0127,
  selectObject: 0x012d,
  deleteObject: 0x01f0,
  setWindowOrigin: 0x020b,
  setWindowExtent: 0x020c,
  setViewportExtent: 0x020e,
  lineTo: 0x0213,
  moveTo: 0x0214,
  createPenIndirect: 0x02fa,
  createBrushIndirect: 0x02fc,
  polygon: 0x0324,
  polyline: 0x0325,
  ellipse: 0x0418,
  rectangle: 0x041b,
  polyPolygon: 0x0538,
} as const;

const wmfRecord = (fn: number, parameters: readonly number[] = []): Uint8Array => {
  const bytes = new Uint8Array(6 + parameters.length * 2);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, bytes.byteLength / 2, true);
  view.setUint16(4, fn, true);
  for (const [index, parameter] of parameters.entries()) {
    view.setUint16(6 + index * 2, parameter & 0xffff, true);
  }
  return bytes;
};

const colorWords = (color: number): [number, number] => [color & 0xffff, color >>> 16];
const pointWords = ({ x, y }: { x: number; y: number }): [number, number] => [x, y];

type WmfOptions = {
  placeable?: boolean;
  bounds?: { left: number; top: number; right: number; bottom: number };
  records?: readonly Uint8Array[];
  objectCount?: number;
};

const vectorWmf = ({
  placeable = true,
  bounds = { left: 10, top: 20, right: 210, bottom: 120 },
  records = [],
  objectCount = 8,
}: WmfOptions = {}): Uint8Array => {
  const allRecords = [...records, wmfRecord(META.eof)];
  const recordBytes = allRecords.reduce((sum, item) => sum + item.byteLength, 0);
  const headerOffset = placeable ? 22 : 0;
  const bytes = new Uint8Array(headerOffset + 18 + recordBytes);
  const view = new DataView(bytes.buffer);
  if (placeable) {
    view.setUint32(0, 0x9ac6_cdd7, true);
    view.setInt16(6, bounds.left, true);
    view.setInt16(8, bounds.top, true);
    view.setInt16(10, bounds.right, true);
    view.setInt16(12, bounds.bottom, true);
    view.setUint16(14, 1_440, true);
    let checksum = 0;
    for (let offset = 0; offset < 20; offset += 2) checksum ^= view.getUint16(offset, true);
    view.setUint16(20, checksum, true);
  }
  view.setUint16(headerOffset, 1, true);
  view.setUint16(headerOffset + 2, 9, true);
  view.setUint16(headerOffset + 4, 0x0300, true);
  view.setUint32(headerOffset + 6, (18 + recordBytes) / 2, true);
  view.setUint16(headerOffset + 10, objectCount, true);
  view.setUint32(
    headerOffset + 12,
    Math.max(...allRecords.map((item) => item.byteLength / 2)),
    true,
  );
  let offset = headerOffset + 18;
  for (const item of allRecords) {
    bytes.set(item, offset);
    offset += item.byteLength;
  }
  return bytes;
};

const mappingRecords = (origin = { x: 10, y: 20 }, extent = { x: 200, y: 100 }) => [
  wmfRecord(META.setMapMode, [8]),
  wmfRecord(META.setWindowOrigin, [origin.y, origin.x]),
  wmfRecord(META.setWindowExtent, [extent.y, extent.x]),
];

const polygonRecord = (points: readonly { x: number; y: number }[]): Uint8Array =>
  wmfRecord(META.polygon, [points.length, ...points.flatMap(pointWords)]);

const minimalEmf = (): Uint8Array => {
  const emfRecord = (type: number, size: number, write?: (view: DataView) => void) => {
    const result = new Uint8Array(size);
    const recordView = new DataView(result.buffer);
    recordView.setUint32(0, type, true);
    recordView.setUint32(4, size, true);
    write?.(recordView);
    return result;
  };
  const records = [
    emfRecord(17, 12, (recordView) => recordView.setUint32(8, 8, true)),
    emfRecord(9, 16, (recordView) => {
      recordView.setInt32(8, 100, true);
      recordView.setInt32(12, 100, true);
    }),
    emfRecord(11, 16, (recordView) => {
      recordView.setInt32(8, 100, true);
      recordView.setInt32(12, 100, true);
    }),
    emfRecord(39, 24, (recordView) => {
      recordView.setUint32(8, 1, true);
      recordView.setUint32(12, 0, true);
      recordView.setUint32(16, 0x0033_2211, true);
    }),
    emfRecord(37, 12, (recordView) => recordView.setUint32(8, 1, true)),
    emfRecord(91, 48, (recordView) => {
      recordView.setUint32(24, 1, true);
      recordView.setUint32(28, 3, true);
      recordView.setUint32(32, 3, true);
      for (const [index, point] of [
        { x: 10, y: 10 },
        { x: 90, y: 10 },
        { x: 50, y: 90 },
      ].entries()) {
        recordView.setInt16(36 + index * 4, point.x, true);
        recordView.setInt16(38 + index * 4, point.y, true);
      }
    }),
    emfRecord(14, 20),
  ];
  const bytes = new Uint8Array(88 + records.reduce((sum, record) => sum + record.byteLength, 0));
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 1, true);
  view.setUint32(4, 88, true);
  view.setInt32(16, 100, true);
  view.setInt32(20, 100, true);
  view.setUint32(40, 0x464d_4520, true);
  view.setUint32(44, 0x0001_0000, true);
  view.setUint32(48, bytes.byteLength, true);
  view.setUint32(52, records.length + 1, true);
  view.setUint16(56, 2, true);
  let offset = 88;
  for (const record of records) {
    bytes.set(record, offset);
    offset += record.byteLength;
  }
  return bytes;
};

describe("renderWmfSvg", () => {
  test("uses the WMF stock pen, brush, and alternate polygon fill mode", () => {
    const rectangle = renderWmfSvg(
      vectorWmf({ records: [...mappingRecords(), wmfRecord(META.rectangle, [120, 210, 20, 10])] }),
    );
    const polygon = renderWmfSvg(
      vectorWmf({
        records: [
          ...mappingRecords(),
          polygonRecord([
            { x: 10, y: 20 },
            { x: 210, y: 120 },
            { x: 210, y: 20 },
            { x: 10, y: 120 },
          ]),
        ],
      }),
    );
    const brushOnly = renderWmfSvg(
      vectorWmf({
        records: [
          ...mappingRecords(),
          wmfRecord(META.createBrushIndirect, [0, ...colorWords(0x0000_00ff), 0]),
          wmfRecord(META.selectObject, [0]),
          polygonRecord([
            { x: 10, y: 20 },
            { x: 210, y: 20 },
            { x: 110, y: 120 },
          ]),
        ],
      }),
    );

    expect(rectangle).toContain('fill="#ffffff" stroke="#000000" stroke-width="1"');
    expect(polygon).toContain('fill-rule="evenodd"');
    expect(brushOnly).toContain('fill="#ff0000" stroke="#000000" stroke-width="1"');
  });

  test("keeps the output viewport fixed when the window extent changes", () => {
    const svg = renderWmfSvg(
      vectorWmf({
        placeable: false,
        records: [
          ...mappingRecords({ x: 0, y: 0 }, { x: 200, y: 100 }),
          wmfRecord(META.rectangle, [100, 200, 0, 0]),
          wmfRecord(META.setWindowExtent, [200, 400]),
          wmfRecord(META.rectangle, [200, 400, 0, 0]),
        ],
      }),
    );

    expect(svg).toContain('viewBox="0 0 200 100"');
    expect(svg?.match(/<rect x="0" y="0" width="200" height="100"/g)).toHaveLength(2);
  });

  test("applies text, anisotropic, and isotropic map modes", () => {
    const anisotropicThenText = renderWmfSvg(
      vectorWmf({
        placeable: false,
        records: [
          wmfRecord(META.setViewportExtent, [100, 200]),
          ...mappingRecords({ x: 0, y: 0 }, { x: 100, y: 100 }),
          wmfRecord(META.rectangle, [20, 20, 10, 10]),
          wmfRecord(META.setMapMode, [1]),
          wmfRecord(META.rectangle, [20, 20, 10, 10]),
        ],
      }),
    );
    const isotropic = renderWmfSvg(
      vectorWmf({
        placeable: false,
        records: [
          wmfRecord(META.setViewportExtent, [100, 200]),
          ...mappingRecords({ x: 0, y: 0 }, { x: 100, y: 100 }),
          wmfRecord(META.setMapMode, [7]),
          wmfRecord(META.rectangle, [20, 20, 10, 10]),
        ],
      }),
    );

    expect(anisotropicThenText).toContain('<rect x="20" y="10" width="20" height="10"');
    expect(anisotropicThenText).toContain('<rect x="10" y="10" width="10" height="10"');
    expect(isotropic).toContain('<rect x="10" y="10" width="10" height="10"');
    expect(
      renderWmfSvg(
        vectorWmf({
          records: [
            wmfRecord(META.setMapMode, [2]),
            polygonRecord([
              { x: 0, y: 0 },
              { x: 1, y: 0 },
              { x: 0, y: 1 },
            ]),
          ],
        }),
      ),
    ).toBeNull();
  });

  test("maps logical pen widths and preserves zero-width hairlines", () => {
    const renderWithPenWidth = (width: number) =>
      renderWmfSvg(
        vectorWmf({
          placeable: false,
          records: [
            wmfRecord(META.setViewportExtent, [100, 100]),
            ...mappingRecords({ x: 0, y: 0 }, { x: 1_000, y: 1_000 }),
            wmfRecord(META.createPenIndirect, [0, width, 0, ...colorWords(0)]),
            wmfRecord(META.selectObject, [0]),
            wmfRecord(META.rectangle, [1_000, 1_000, 0, 0]),
          ],
        }),
      );

    expect(renderWithPenWidth(20)).toContain('stroke-width="2"');
    expect(renderWithPenWidth(0)).toContain('stroke-width="1"');
  });

  test("renders common object, state, line, polygon, rectangle, and ellipse records", () => {
    const records = [
      ...mappingRecords(),
      wmfRecord(META.setBackgroundMode, [1]),
      wmfRecord(META.setPolyFillMode, [1]),
      wmfRecord(META.createPenIndirect, [0, 1, 0, ...colorWords(0x0033_2211)]),
      wmfRecord(META.createBrushIndirect, [0, ...colorWords(0x0066_5544), 0]),
      wmfRecord(META.selectObject, [0]),
      wmfRecord(META.selectObject, [1]),
      wmfRecord(META.saveDc),
      wmfRecord(META.moveTo, [30, 30]),
      wmfRecord(META.lineTo, [40, 80]),
      polygonRecord([
        { x: 30, y: 30 },
        { x: 60, y: 30 },
        { x: 45, y: 60 },
      ]),
      wmfRecord(META.polyline, [3, 30, 70, 40, 80, 50, 70]),
      wmfRecord(META.rectangle, [80, 100, 40, 60]),
      wmfRecord(META.ellipse, [90, 160, 50, 120]),
      wmfRecord(META.restoreDc, [0xffff]),
      wmfRecord(META.deleteObject, [0]),
      wmfRecord(META.deleteObject, [1]),
    ];
    const svg = renderWmfSvg(vectorWmf({ records }));

    expect(svg).toStartWith("<svg");
    expect(svg).toContain('viewBox="0 0 200 100"');
    expect(svg).toContain('fill="#445566"');
    expect(svg).toContain('fill-rule="evenodd"');
    expect(svg).toContain("<rect ");
    expect(svg).toContain("<ellipse ");
    expect(svg).toContain("stroke=");
  });

  test("renders POLYPOLYGON and applies the window origin with the correct sign", () => {
    const polyPolygon = wmfRecord(
      META.polyPolygon,
      [2, 3, 3, 10, 20, 30, 20, 20, 40, 110, 70, 130, 70, 120, 90],
    );
    const svg = renderWmfSvg(
      vectorWmf({
        records: [
          ...mappingRecords(),
          wmfRecord(META.createBrushIndirect, [0, ...colorWords(0x0000_00ff), 0]),
          wmfRecord(META.selectObject, [0]),
          polyPolygon,
        ],
      }),
    );

    expect(svg).toContain('d="M0 0 L20 0 L10 20 Z M100 50 L120 50 L110 70 Z"');
  });

  test("accepts a plain WMF header and derives its viewBox from the window extent", () => {
    const svg = renderWmfSvg(
      vectorWmf({
        placeable: false,
        records: [
          ...mappingRecords({ x: -20, y: -10 }, { x: 320, y: 180 }),
          wmfRecord(META.createBrushIndirect, [0, ...colorWords(0), 0]),
          wmfRecord(META.selectObject, [0]),
          polygonRecord([
            { x: 0, y: 0 },
            { x: 10, y: 0 },
            { x: 5, y: 10 },
          ]),
        ],
      }),
    );
    expect(svg).toContain('viewBox="0 0 320 180"');
  });

  test("fails closed for an unsupported record", () => {
    expect(renderWmfSvg(vectorWmf({ records: [wmfRecord(0x0626, [1])] }))).toBeNull();
  });

  test("enforces record, point, dimension, state-depth, and SVG-output caps", () => {
    expect(
      renderWmfSvg(
        vectorWmf({ records: Array.from({ length: 10_001 }, () => wmfRecord(META.saveDc)) }),
      ),
    ).toBeNull();
    expect(
      renderWmfSvg(
        vectorWmf({
          records: [polygonRecord(Array.from({ length: 20_001 }, () => ({ x: 1, y: 1 })))],
        }),
      ),
    ).toBeNull();
    expect(
      renderWmfSvg(
        vectorWmf({
          placeable: false,
          records: mappingRecords({ x: 0, y: 0 }, { x: 100_001, y: 1 }),
        }),
      ),
    ).toBeNull();
    expect(
      renderWmfSvg(
        vectorWmf({ records: Array.from({ length: 257 }, () => wmfRecord(META.saveDc)) }),
      ),
    ).toBeNull();
    const verboseRectangles = Array.from({ length: 9_990 }, () =>
      wmfRecord(META.rectangle, [32_000, 32_000, 0x8300, 0x8300]),
    );
    expect(
      renderWmfSvg(
        vectorWmf({
          bounds: { left: -32_000, top: -32_000, right: 32_000, bottom: 32_000 },
          records: [
            ...mappingRecords({ x: -32_000, y: -32_000 }, { x: 64_000, y: 64_000 }),
            ...verboseRectangles,
          ],
        }),
      ),
    ).toBeNull();
  });

  test("never throws for well-formed polygon WMFs and preserves their window viewBox", () => {
    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: 2_000 }),
        fc.integer({ min: 1, max: 2_000 }),
        fc.array(
          fc.record({
            x: fc.integer({ min: -2_000, max: 2_000 }),
            y: fc.integer({ min: -2_000, max: 2_000 }),
          }),
          { minLength: 3, maxLength: 40 },
        ),
        (width, height, points) => {
          const result = renderWmfSvg(
            vectorWmf({
              placeable: false,
              records: [
                ...mappingRecords({ x: 0, y: 0 }, { x: width, y: height }),
                polygonRecord(points),
              ],
            }),
          );
          expect(
            result === null ||
              (result.endsWith("</svg>") && result.includes(`viewBox="0 0 ${width} ${height}"`)),
          ).toBeTrue();
        },
      ),
      { numRuns: 200 },
    );
  });

  test("maps every full-window rectangle to the fixed viewBox", () => {
    fc.assert(
      fc.property(
        fc.integer({ min: 1, max: 2_000 }),
        fc.integer({ min: 1, max: 2_000 }),
        fc.integer({ min: 1, max: 2_000 }),
        fc.integer({ min: 1, max: 2_000 }),
        (viewWidth, viewHeight, windowWidth, windowHeight) => {
          const svg = renderWmfSvg(
            vectorWmf({
              placeable: false,
              records: [
                wmfRecord(META.setViewportExtent, [viewHeight, viewWidth]),
                ...mappingRecords({ x: 0, y: 0 }, { x: windowWidth, y: windowHeight }),
                wmfRecord(META.rectangle, [windowHeight, windowWidth, 0, 0]),
              ],
            }),
          );

          expect(svg).toContain(`viewBox="0 0 ${viewWidth} ${viewHeight}"`);
          expect(svg).toContain(`<rect x="0" y="0" width="${viewWidth}" height="${viewHeight}"`);
        },
      ),
      { numRuns: 200 },
    );
  });

  test("emits only finite mapped values while keeping the viewport fixed", () => {
    fc.assert(
      fc.property(
        fc.constantFrom(1, 7, 8),
        fc.integer({ min: 1, max: 2_000 }),
        fc.integer({ min: 1, max: 2_000 }),
        fc.integer({ min: 1, max: 2_000 }),
        fc.integer({ min: 1, max: 2_000 }),
        fc.integer({ min: 0, max: 2_000 }),
        (mode, viewWidth, viewHeight, windowWidth, windowHeight, penWidth) => {
          const svg = renderWmfSvg(
            vectorWmf({
              placeable: false,
              records: [
                wmfRecord(META.setViewportExtent, [viewHeight, viewWidth]),
                ...mappingRecords({ x: 0, y: 0 }, { x: windowWidth, y: windowHeight }),
                wmfRecord(META.setMapMode, [mode]),
                wmfRecord(META.createPenIndirect, [0, penWidth, 0, ...colorWords(0)]),
                wmfRecord(META.selectObject, [0]),
                wmfRecord(META.rectangle, [windowHeight, windowWidth, 0, 0]),
              ],
            }),
          );

          expect(svg).not.toBeNull();
          expect(svg).toContain(`viewBox="0 0 ${viewWidth} ${viewHeight}"`);
          expect(svg).not.toMatch(/(?:NaN|Infinity)/);
          const rectangle = svg?.match(/<rect [^>]+/)?.at(0);
          expect(rectangle).toBeDefined();
          for (const value of rectangle?.matchAll(/(?:x|y|width|height|stroke-width)="([^"]+)"/g) ??
            []) {
            expect(Number.isFinite(Number(value[1]))).toBeTrue();
          }
        },
      ),
      { numRuns: 200 },
    );
  });
});

describe("WMF package previews", () => {
  test("generates SVG previews for EMF and WMF while preserving their source bytes", async () => {
    const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
    const emf = minimalEmf();
    const wmf = vectorWmf({
      records: [
        ...mappingRecords(),
        wmfRecord(META.createBrushIndirect, [0, ...colorWords(0x00aa_5500), 0]),
        wmfRecord(META.selectObject, [0]),
        polygonRecord([
          { x: 20, y: 30 },
          { x: 100, y: 30 },
          { x: 60, y: 90 },
        ]),
      ],
    });
    zip.file("word/media/image1.emf", emf);
    zip.file("word/media/image2.wmf", wmf);
    const relationshipPath = "word/_rels/document.xml.rels";
    const relationships = await zip.file(relationshipPath)?.async("string");
    if (!relationships) throw new Error("Fixture is missing document relationships");
    zip.file(
      relationshipPath,
      relationships.replace(
        "</Relationships>",
        '<Relationship Id="rWmf" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.wmf"/></Relationships>',
      ),
    );

    const document = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
      preloadFonts: false,
    });
    const emfMedia = document.package.media?.get("word/media/image1.emf");
    const wmfMedia = document.package.media?.get("word/media/image2.wmf");
    expect(emfMedia?.dataUrl).toStartWith("data:image/svg+xml");
    expect(wmfMedia?.dataUrl).toStartWith("data:image/svg+xml");
    expect(new Uint8Array(emfMedia?.data ?? new ArrayBuffer(0))).toEqual(emf);
    expect(new Uint8Array(wmfMedia?.data ?? new ArrayBuffer(0))).toEqual(wmf);

    const output = await JSZip.loadAsync(await repackDocx(document));
    expect(await output.file("word/media/image1.emf")?.async("uint8array")).toEqual(emf);
    expect(await output.file("word/media/image2.wmf")?.async("uint8array")).toEqual(wmf);
    expect(renderEmfSvg(emf)).not.toBeNull();
  });
});
