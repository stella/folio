/**
 * A run that holds a drawing can carry tracked formatting (`w:rPrChange`)
 * like any run. The editor projection keeps it: the change list shows it, a
 * save from the editor keeps it, and accepting or rejecting it through the
 * reviewer resolves it.
 */
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { listedChangeCoverageProblem } from "../../../../test/listedChangeCoverage";
import { DOCUMENT_SHAPES } from "../__tests__/documentShapes";
import {
  createHarnessState,
  parseShapeDocument,
  saveHarnessState,
} from "../__tests__/editorHarness";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Run } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import { getTrackedChangesFromDoc } from "./read";

const REVISION = 77;
const IMAGE_SHAPE =
  DOCUMENT_SHAPES.find(({ id }) => id === "image") ??
  ((): never => {
    throw new TypeError("The image shape is missing.");
  })();

/** The image shape with its drawing's run struck through as a tracked change. */
const trackedDrawingPackage = async () => {
  const zip = await ZipArchive.read((await IMAGE_SHAPE.build()).slice().buffer);
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  const tracked = xml.replace(
    "<w:r><w:drawing>",
    `<w:r><w:rPr><w:strike/><w:rPrChange w:id="${String(REVISION)}" w:author="Reviewer" w:date="2026-01-02T03:04:05Z"><w:rPr/></w:rPrChange></w:rPr><w:drawing>`,
  );
  expect(tracked).not.toBe(xml);
  zip.set("word/document.xml", tracked);
  return zip.generate();
};

const drawingRuns = (document: Document): Run[] =>
  document.package.document.content.flatMap((block) =>
    block.type === "paragraph"
      ? block.content.flatMap((item) =>
          item.type === "run" &&
          item.content.some((content) => content.type === "drawing")
            ? [item]
            : [],
        )
      : [],
  );

const drawingRun = (document: Document): Run => {
  const [run, ...rest] = drawingRuns(document);
  if (run === undefined || rest.length > 0)
    throw new TypeError("Expected exactly one drawing run.");
  return run;
};

describe("tracked formatting on a drawing run", () => {
  test("the change list shows it and reaches every revision", async () => {
    const document = await parseShapeDocument(await trackedDrawingPackage());
    expect(
      drawingRun(document).propertyChanges?.map(({ info }) => info.id),
    ).toEqual([REVISION]);
    const doc = toProseDoc(document);
    expect(
      getTrackedChangesFromDoc(doc).map(({ id, type }) => ({ id, type })),
    ).toEqual([{ id: REVISION, type: "formatting" }]);
    expect(listedChangeCoverageProblem(document, doc)).toBeNull();
  });

  test("a save from the editor keeps it", async () => {
    const document = await parseShapeDocument(await trackedDrawingPackage());
    const saved = await saveHarnessState(
      createHarnessState(document, "editing"),
      document,
    );
    const reopened = drawingRun(await parseShapeDocument(saved.bytes));
    expect(reopened.formatting?.strike).toBe(true);
    expect(reopened.propertyChanges?.map(({ info }) => info.id)).toEqual([
      REVISION,
    ]);
  });

  for (const decision of ["accept", "reject"] as const)
    test(`the reviewer can ${decision} it`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        (await trackedDrawingPackage()).slice().buffer,
      );
      expect(reviewer.getChanges().map(({ id }) => id)).toEqual([REVISION]);
      expect(
        decision === "accept"
          ? reviewer.acceptChange(REVISION)
          : reviewer.rejectChange(REVISION),
      ).toBe(true);
      expect(reviewer.getChanges()).toEqual([]);
      const reopened = drawingRun(
        await parseShapeDocument(new Uint8Array(await reviewer.toBuffer())),
      );
      expect(reopened.propertyChanges ?? []).toEqual([]);
      expect(reopened.formatting?.strike === true).toBe(decision === "accept");
    });
});
