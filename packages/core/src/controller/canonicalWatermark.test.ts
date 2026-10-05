import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { validateDocxPackage } from "@stll/docx-core";
import { DOCUMENT_OP_TYPES } from "@stll/docx-core/ops";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { FOLIO_DOCX_SERIALIZATION_MODE } from "../types/docxSerialization";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { expectCanonicalWatermarkRoundTrip } from "../../../../test/canonicalWatermarkRoundTrip";
import { EditorState } from "prosemirror-state";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document } from "../types/document";
import { createCanonicalWatermarkOperation } from "./canonicalWatermark";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(30_000));

test("generated watermark histories allocate distinct coverage and host identities and preserve exact history", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.integer({ min: 0, max: 2 }),
      fc.boolean(),
      fc.array(fc.stringMatching(/^[A-Z]{1,12}$/u), { minLength: 1, maxLength: 5 }),
      fc.constantFrom("fresh", "parsed"),
      fc.boolean(),
      async (headerCount, evenPages, texts, sourceKind, hostFormatting) => {
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
              finalSectionProperties: {
                titlePg: true,
                ...(headerCount === 0
                  ? {}
                  : {
                      headerReferences: Array.from({ length: headerCount }, (_, index) => ({
                        type: index === 0 ? ("default" as const) : ("first" as const),
                        rId: `rId_wm_${index === 0 ? "default" : "first"}`,
                      })),
                    }),
              },
            },
            settings: { defaultTabStop: 720, evenAndOddHeaders: evenPages },
            headers: new Map(
              Array.from({ length: headerCount }, (_, index) => [
                `rId_wm_${index === 0 ? "default" : "first"}`,
                {
                  type: "header",
                  hdrFtrType: index === 0 ? "default" : "first",
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
        const source =
          sourceKind === "fresh"
            ? document
            : await parseDocx(await createDocx(document), { preloadFonts: false });
        const session = createCanonicalSession(source).unwrap();
        expect(session.document.originalBuffer === undefined).toBe(sourceKind === "fresh");
        const assertSave = async () => {
          const snapshot = session.captureSaveSnapshot();
          for (const mode of Object.values(FOLIO_DOCX_SERIALIZATION_MODE)) {
            const saved = await serializeCanonicalSave({
              snapshot,
              featureFlags: { selectiveSave: true },
              options: { mode },
            });
            expect(await validateDocxPackage(new Uint8Array(saved.buffer))).toEqual({
              valid: true,
            });
            const reopened = await parseDocx(saved.buffer, { preloadFonts: false });
            expectCanonicalWatermarkRoundTrip(snapshot.document, reopened);
          }
        };
        let state = EditorState.create({ doc: session.projection.doc });
        const before = session.document;
        let steps = 0;
        const coverageIdentities = new Map<string, string>();
        for (const text of texts) {
          const operation = createCanonicalWatermarkOperation(session.document, {
            kind: "set",
            watermark: { kind: "text", text },
          }).unwrap();
          const allocated = [...operation.coverage, ...operation.hosts].map(({ paraId }) => paraId);
          expect(new Set(allocated).size).toBe(allocated.length);
          for (const id of allocated) expect(id).toMatch(/^[0-9A-F]{8}$/u);
          for (const coverage of operation.coverage) {
            expect(session.document.package.headers?.has(coverage.rId)).toBe(false);
            coverageIdentities.set(coverage.rId, coverage.paraId);
          }
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
          if (hostFormatting && operation.coverage.length > 0) {
            const formatting = session
              .prepareOperations(
                state,
                operation.coverage.map(({ rId, paraId }) => ({
                  type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
                  story: { kind: "header", rId },
                  blockId: paraId,
                  patch: { styleId: "HeaderHost" },
                })),
              )
              .unwrap();
            state = publishCanonicalProjection({ state, commit: formatting, session }).unwrap()
              .state;
            steps += 1;
            for (const { rId } of operation.coverage) {
              const host = session.document.package.headers?.get(rId)?.content.at(0);
              expect(host?.type === "paragraph" ? host.formatting?.styleId : undefined).toBe(
                "HeaderHost",
              );
            }
          }
          await assertSave();
          const removal = createCanonicalWatermarkOperation(session.document, {
            kind: "remove",
          }).unwrap();
          const removed = session.prepareOperations(state, [removal]).unwrap();
          state = publishCanonicalProjection({ state, commit: removed, session }).unwrap().state;
          steps += 1;
          for (const header of session.document.package.headers?.values() ?? []) {
            expect(header.watermark).toBeUndefined();
            expect(header.watermarkBlockIndex).toBeUndefined();
          }
          for (const [rId, paraId] of coverageIdentities) {
            // Sole-host retention preserves only the allocated identity: no
            // pPr, run, drawing, or other decoration survives removal.
            expect(session.document.package.headers?.get(rId)?.content).toStrictEqual([
              { type: "paragraph", paraId, content: [] },
            ]);
          }
          await assertSave();
        }
        const after = session.document;
        for (let index = 0; index < steps; index += 1) {
          const commit = session.prepareUndo(state).unwrap();
          state = publishCanonicalProjection({ state, commit, session }).unwrap().state;
        }
        expect(session.document).toStrictEqual(before);
        await assertSave();
        for (let index = 0; index < steps; index += 1) {
          const commit = session.prepareRedo(state).unwrap();
          state = publishCanonicalProjection({ state, commit, session }).unwrap().state;
        }
        expect(session.document).toStrictEqual(after);
        await assertSave();
        expect(session.projection.doc.textContent).toBe("Body");
      },
    ),
    {
      seed: 20261015,
      numRuns: 30,
      examples: [
        [0, false, ["A"], "fresh", true],
        [2, true, ["A"], "parsed", true],
      ],
    },
  );
});
