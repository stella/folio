import { expect } from "bun:test";
import { panic } from "better-result";
import type { Document } from "../packages/docx-core/src/model/document";
import { reviewDifferences } from "./reviewDifferences";

/** Watermark save equality includes secondary-story identity and authored decoration. */
export const expectCanonicalWatermarkRoundTrip = (before: Document, after: Document) => {
  expect(reviewDifferences(before, after)).toEqual({ messages: [], omitted: 0 });
  expect(after.package.document.finalSectionProperties).toEqual(
    before.package.document.finalSectionProperties,
  );
  expect([...(after.package.headers?.keys() ?? [])].sort()).toEqual(
    [...(before.package.headers?.keys() ?? [])].sort(),
  );
  for (const [rId, header] of before.package.headers ?? []) {
    const reopened = after.package.headers?.get(rId) ?? panic("Expected reopened watermark header");
    expect(reopened.type).toBe(header.type);
    expect(reopened.hdrFtrType).toBe(header.hdrFtrType);
    expect(reopened.watermark).toEqual(header.watermark);
    expect(reopened.watermarkBlockIndex).toBe(header.watermarkBlockIndex);
    expect(
      reviewDifferences(
        { package: { document: { content: header.content } } },
        { package: { document: { content: reopened.content } } },
      ),
    ).toEqual({ messages: [], omitted: 0 });
  }
};
