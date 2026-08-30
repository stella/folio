import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "./parser";
import { repackDocx } from "./rezip";
import { renderEmfSvg } from "./metafileSvg";

const FIXTURE = resolve(import.meta.dir, "__fixtures__/header-vml-emf.docx");

const record = (type: number, size: number, write?: (view: DataView) => void): Uint8Array => {
  const bytes = new Uint8Array(size);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, type, true);
  view.setUint32(4, size, true);
  write?.(view);
  return bytes;
};

const uint32Record = (type: number, value: number): Uint8Array =>
  record(type, 12, (view) => view.setUint32(8, value, true));

const pointRecord = (type: number, x: number, y: number): Uint8Array =>
  record(type, 16, (view) => {
    view.setInt32(8, x, true);
    view.setInt32(12, y, true);
  });

const polyBezierTo16 = (points: readonly { x: number; y: number }[]): Uint8Array =>
  record(88, 28 + points.length * 4, (view) => {
    view.setUint32(24, points.length, true);
    for (const [index, point] of points.entries()) {
      view.setInt16(28 + index * 4, point.x, true);
      view.setInt16(30 + index * 4, point.y, true);
    }
  });

const polyPolygon16 = (polygons: readonly (readonly { x: number; y: number }[])[]): Uint8Array => {
  const points = polygons.flat();
  return record(91, 32 + polygons.length * 4 + points.length * 4, (view) => {
    view.setUint32(24, polygons.length, true);
    view.setUint32(28, points.length, true);
    for (const [index, polygon] of polygons.entries()) {
      view.setUint32(32 + index * 4, polygon.length, true);
    }
    const pointsOffset = 32 + polygons.length * 4;
    for (const [index, point] of points.entries()) {
      view.setInt16(pointsOffset + index * 4, point.x, true);
      view.setInt16(pointsOffset + index * 4 + 2, point.y, true);
    }
  });
};

type VectorEmfOptions = {
  windowExtent?: { x: number; y: number };
  windowOrigin?: { x: number; y: number };
  viewportOrigin?: { x: number; y: number };
  extraRecord?: Uint8Array;
  brushHandle?: number;
  deleteBrushBeforePolygon?: boolean;
  duplicateBrushHandle?: boolean;
  mapMode?: 1 | 8;
  includeMoveTo?: boolean;
  polygonPointCount?: number;
  polygonBeforeFillPath?: boolean;
};

const vectorEmf = ({
  windowExtent = { x: 1_000, y: 1_000 },
  windowOrigin = { x: 0, y: 0 },
  viewportOrigin = { x: 0, y: 0 },
  extraRecord,
  brushHandle = 1,
  deleteBrushBeforePolygon = false,
  duplicateBrushHandle = false,
  mapMode = 8,
  includeMoveTo = true,
  polygonPointCount = 3,
  polygonBeforeFillPath = false,
}: VectorEmfOptions = {}): Uint8Array => {
  const polygon =
    polygonPointCount === 3
      ? [
          { x: 500, y: 500 },
          { x: 700, y: 500 },
          { x: 600, y: 700 },
        ]
      : Array.from({ length: polygonPointCount }, (_, index) => ({
          x: -32_000 + (index % 64_000),
          y: -31_000 + ((index * 17) % 62_000),
        }));
  const records = [
    uint32Record(17, mapMode),
    pointRecord(10, windowOrigin.x, windowOrigin.y),
    pointRecord(12, viewportOrigin.x, viewportOrigin.y),
    pointRecord(9, windowExtent.x, windowExtent.y),
    pointRecord(11, 100, 100),
    record(33, 8),
    uint32Record(19, 1),
    record(39, 24, (view) => {
      view.setUint32(8, 1, true);
      view.setUint32(12, 0, true);
      view.setUint32(16, 0x0033_2211, true);
    }),
    ...(duplicateBrushHandle
      ? [
          record(39, 24, (view) => {
            view.setUint32(8, 1, true);
            view.setUint32(12, 0, true);
            view.setUint32(16, 0x0066_5544, true);
          }),
        ]
      : []),
    uint32Record(37, brushHandle),
    record(95, 52, (view) => {
      view.setUint32(8, 2, true);
      view.setUint32(28, 5, true);
    }),
    uint32Record(37, 2),
    record(59, 8),
    ...(includeMoveTo ? [pointRecord(27, 100, 100)] : []),
    polyBezierTo16([
      { x: 200, y: 100 },
      { x: 200, y: 200 },
      { x: 100, y: 200 },
    ]),
    record(61, 8),
    record(60, 8),
    ...(polygonBeforeFillPath ? [polyPolygon16([polygon])] : []),
    record(62, 24),
    ...(deleteBrushBeforePolygon ? [uint32Record(40, 1)] : []),
    uint32Record(19, 2),
    polyPolygon16([polygon]),
    ...(extraRecord ? [extraRecord] : []),
    uint32Record(34, 0xffff_ffff),
    ...(!deleteBrushBeforePolygon ? [uint32Record(40, 1)] : []),
    uint32Record(40, 2),
    record(14, 20),
  ];
  const byteLength = 88 + records.reduce((sum, item) => sum + item.byteLength, 0);
  const header = record(1, 88, (view) => {
    view.setInt32(8, 0, true);
    view.setInt32(12, 0, true);
    view.setInt32(16, 100, true);
    view.setInt32(20, 100, true);
    view.setInt32(24, 0, true);
    view.setInt32(28, 0, true);
    view.setInt32(32, 2_646, true);
    view.setInt32(36, 2_646, true);
    view.setUint32(40, 0x464d_4520, true);
    view.setUint32(44, 0x0001_0000, true);
    view.setUint32(48, byteLength, true);
    view.setUint32(52, records.length + 1, true);
    view.setUint16(56, 3, true);
    view.setInt32(72, 100, true);
    view.setInt32(76, 100, true);
    view.setInt32(80, 26, true);
    view.setInt32(84, 26, true);
  });
  const bytes = new Uint8Array(byteLength);
  bytes.set(header);
  let offset = header.byteLength;
  for (const item of records) {
    bytes.set(item, offset);
    offset += item.byteLength;
  }
  return bytes;
};

describe("renderEmfSvg", () => {
  test("maps filled Bezier and polygon paths into the EMF bounds", () => {
    const svg = renderEmfSvg(vectorEmf());

    expect(svg).toContain('viewBox="0 0 100 100"');
    expect(svg).toContain('d="M10 10 C20 10 20 20 10 20 Z"');
    expect(svg).toContain('fill="#112233" fill-rule="evenodd"');
    expect(svg).toContain('d="M50 50 L70 50 L60 70 Z"');
    expect(svg).toContain('fill="#112233" fill-rule="nonzero"');
    expect(svg).not.toContain("stroke=");
  });

  test("rejects invalid signatures, truncation, and unsupported records", () => {
    const invalidSignature = vectorEmf();
    new DataView(invalidSignature.buffer).setUint32(40, 0, true);
    expect(renderEmfSvg(invalidSignature)).toBeNull();

    const valid = vectorEmf();
    expect(renderEmfSvg(valid.subarray(0, valid.length - 4))).toBeNull();
    expect(renderEmfSvg(vectorEmf({ extraRecord: record(42, 24) }))).toBeNull();
  });

  test("rejects zero mapping extents before evaluating any points", () => {
    expect(renderEmfSvg(vectorEmf({ windowExtent: { x: 0, y: 1_000 } }))).toBeNull();
  });

  test("selects stock brushes using unsigned EMF handles", () => {
    const svg = renderEmfSvg(vectorEmf({ brushHandle: 0x8000_0004 }));
    expect(svg).toContain('fill="#000000"');
  });

  test("uses device coordinates for MM_TEXT instead of anisotropic extents", () => {
    const svg = renderEmfSvg(
      vectorEmf({
        mapMode: 1,
        windowOrigin: { x: 10, y: 20 },
        viewportOrigin: { x: 5, y: 7 },
      }),
    );
    expect(svg).toContain('d="M95 87 C195 87 195 187 95 187 Z"');
  });

  test("rejects MM_TEXT coordinates far outside the declared output", () => {
    const bytes = vectorEmf({ mapMode: 1 });
    const view = new DataView(bytes.buffer);
    view.setInt32(8, -2_000_000_000, true);
    view.setInt32(16, -1_999_999_900, true);
    expect(renderEmfSvg(bytes)).toBeNull();
  });

  test("starts a Bezier path from the current GDI point", () => {
    const svg = renderEmfSvg(vectorEmf({ includeMoveTo: false }));
    expect(svg).toContain('d="M0 0 C20 10 20 20 10 20 Z"');
  });

  test("does not reuse a selected brush after its object is deleted", () => {
    expect(renderEmfSvg(vectorEmf({ deleteBrushBeforePolygon: true }))).toBeNull();
  });

  test("preserves a completed path while rendering standalone geometry", () => {
    const svg = renderEmfSvg(vectorEmf({ polygonBeforeFillPath: true }));

    expect(svg).toContain('d="M10 10 C20 10 20 20 10 20 Z"');
    expect(svg?.match(/<path /gu)).toHaveLength(3);
  });

  test("rejects object handles that are recreated before deletion", () => {
    expect(renderEmfSvg(vectorEmf({ duplicateBrushHandle: true }))).toBeNull();
  });

  test("rejects point counts that exceed their record payload", () => {
    const bytes = vectorEmf();
    const view = new DataView(bytes.buffer);
    let offset = 88;
    while (offset + 8 <= bytes.length) {
      const type = view.getUint32(offset, true);
      const size = view.getUint32(offset + 4, true);
      if (type === 88) {
        view.setUint32(offset + 24, 1_000_001, true);
        break;
      }
      offset += size;
    }
    expect(renderEmfSvg(bytes)).toBeNull();
  });
});

describe("vector-only EMF package previews", () => {
  test("uses generated SVG while preserving original EMF bytes and MIME", async () => {
    const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
    const emf = vectorEmf();
    zip.file("word/media/image1.emf", emf);
    const input = await zip.generateAsync({ type: "arraybuffer" });

    const document = await parseDocx(input, { preloadFonts: false });
    const media = document.package.media?.get("word/media/image1.emf");
    expect(media?.dataUrl).toStartWith("data:image/svg+xml");
    expect(media?.mimeType).toBe("image/x-emf");
    expect(new Uint8Array(media?.data ?? new ArrayBuffer(0))).toEqual(emf);

    const output = await repackDocx(document);
    const outputZip = await JSZip.loadAsync(output);
    const roundTripped = await outputZip.file("word/media/image1.emf")?.async("uint8array");
    expect(roundTripped).toEqual(emf);
  });

  test("bounds generated previews across the complete package", async () => {
    const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
    const emf = vectorEmf({
      polygonPointCount: 19_990,
      windowExtent: { x: 1, y: 1 },
    });
    const relationshipPath = "word/_rels/document.xml.rels";
    const relationships = await zip.file(relationshipPath)?.async("string");
    if (!relationships) {
      throw new Error("Fixture is missing document relationships");
    }
    const mediaCount = 24;
    const addedRelationships = Array.from(
      { length: mediaCount },
      (_, index) =>
        `<Relationship Id="rEmfBudget${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/budget-${index}.emf"/>`,
    ).join("");
    zip.file(
      relationshipPath,
      relationships.replace("</Relationships>", `${addedRelationships}</Relationships>`),
    );
    for (let index = 0; index < mediaCount; index += 1) {
      zip.file(`word/media/budget-${index}.emf`, emf);
    }

    const input = await zip.generateAsync({ type: "arraybuffer" });
    const document = await parseDocx(input, { preloadFonts: false });
    const previews = Array.from({ length: mediaCount }, (_, index) =>
      document.package.media?.get(`word/media/budget-${index}.emf`),
    );
    const generatedCount = previews.filter((media) =>
      media?.dataUrl?.startsWith("data:image/svg+xml"),
    ).length;

    expect(generatedCount).toBeGreaterThan(0);
    expect(generatedCount).toBeLessThan(mediaCount);
    expect(previews.at(-1)?.dataUrl).toStartWith("data:image/x-emf");
  });
});
