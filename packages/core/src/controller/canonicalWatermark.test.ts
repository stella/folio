import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document } from "../types/document";
import { createCanonicalWatermarkOperation } from "./canonicalWatermark";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(30_000));

test("generated watermark histories allocate distinct coverage and host identities and preserve exact history", () => {
  assertProperty(
    fc.property(
      fc.integer({ min: 0, max: 2 }),
      fc.boolean(),
      fc.array(fc.stringMatching(/^[A-Z]{1,12}$/u), { minLength: 1, maxLength: 5 }),
      (headerCount, evenPages, texts) => {
        const document = {
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "00000001",
                  content: [{ type: "run", content: [{ type: "text", text: "Body" }] }],
                },
              ],
              finalSectionProperties: { titlePg: true },
            },
            settings: { defaultTabStop: 720, evenAndOddHeaders: evenPages },
            headers: new Map(
              Array.from({ length: headerCount }, (_, index) => [
                `rId_wm_${index === 0 ? "default" : "first"}`,
                {
                  type: "header",
                  hdrFtrType: "default",
                  content: [
                    {
                      type: "paragraph",
                      paraId: (32 + index).toString(16).padStart(8, "0"),
                      content: [{ type: "run", content: [{ type: "text", text: "Header" }] }],
                    },
                  ],
                },
              ]),
            ),
          },
        } satisfies Document;
        const session = createCanonicalSession(document).unwrap();
        let state = EditorState.create({ doc: session.projection.doc });
        const before = session.document;
        let steps = 0;
        for (const text of texts) {
          const operation = createCanonicalWatermarkOperation(session.document, {
            kind: "set",
            watermark: { kind: "text", text },
          }).unwrap();
          const allocated = [...operation.coverage, ...operation.hosts].map(({ paraId }) => paraId);
          expect(new Set(allocated).size).toBe(allocated.length);
          for (const id of allocated) expect(id).toMatch(/^[0-9A-F]{8}$/u);
          for (const coverage of operation.coverage)
            expect(session.document.package.headers?.has(coverage.rId)).toBe(false);
          const prepared = session.prepareOperations(state, [operation]).unwrap();
          state = publishCanonicalProjection({ state, commit: prepared, session }).unwrap().state;
          steps += 1;
          for (const header of session.document.package.headers?.values() ?? []) {
            expect(header.watermark).toMatchObject({
              kind: "text",
              text,
              font: "Calibri",
              color: "C0C0C0",
              diagonal: true,
            });
            expect(header.content.at(header.watermarkBlockIndex ?? -1)?.type).toBe("paragraph");
          }
          const removal = createCanonicalWatermarkOperation(session.document, {
            kind: "remove",
          }).unwrap();
          const removed = session.prepareOperations(state, [removal]).unwrap();
          state = publishCanonicalProjection({ state, commit: removed, session }).unwrap().state;
          steps += 1;
          for (const header of session.document.package.headers?.values() ?? [])
            expect(header.watermark).toBeUndefined();
        }
        const after = session.document;
        for (let index = 0; index < steps; index += 1) {
          const commit = session.prepareUndo(state).unwrap();
          state = publishCanonicalProjection({ state, commit, session }).unwrap().state;
        }
        expect(session.document).toStrictEqual(before);
        for (let index = 0; index < steps; index += 1) {
          const commit = session.prepareRedo(state).unwrap();
          state = publishCanonicalProjection({ state, commit, session }).unwrap().state;
        }
        expect(session.document).toStrictEqual(after);
        expect(session.projection.doc.textContent).toBe("Body");
      },
    ),
    { seed: 20261015, numRuns: 30 },
  );
});
