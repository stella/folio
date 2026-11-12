/**
 * A run that holds a picture, a shape or a text box anchor can carry tracked
 * formatting (`w:rPrChange`) like any run. The editor projection keeps it: the
 * change list shows it, a save from the editor keeps it with the formatting it
 * replaced, and accepting or rejecting it, in the editor (on its own or with
 * every change) or through the reviewer, resolves it.
 */
import { describe, expect, test } from "bun:test";
import type { EditorState } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { listedChangeCoverageProblem } from "../../../../test/listedChangeCoverage";
import {
  type DocumentShape,
  documentShape,
  TEXT_BOX_SHAPE,
} from "../__tests__/documentShapes";
import {
  createHarnessState,
  parseShapeDocument,
  resolveAllChanges,
  saveHarnessState,
} from "../__tests__/editorHarness";
import { acceptChange, rejectChange } from "../prosemirror/commands/comments";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Run, TextFormatting } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import { getTrackedChangesFromDoc } from "./read";

const REVISION = 77;

/** Each object a run can hold, by the shape that holds it once. */
const CASES: readonly {
  object: string;
  shape: DocumentShape;
  direct: TextFormatting;
}[] = [
  { object: "picture", shape: documentShape("image"), direct: {} },
  {
    object: "shape",
    shape: documentShape("drawing-shape"),
    direct: { bold: true },
  },
  {
    object: "text box anchor",
    shape: TEXT_BOX_SHAPE,
    direct: { italic: true },
  },
];

/** The drawing's run, its own `w:rPr` (if any) and the drawing it holds. */
const DRAWING_RUN =
  /<w:r>(?:<w:rPr>((?:(?!<\/w:rPr>).)*)<\/w:rPr>)?<w:drawing>/u;

/**
 * The shape with its drawing's run struck through as a tracked change, the
 * run's direct formatting kept as the formatting the change replaced.
 */
const trackedPackage = async (shape: DocumentShape) => {
  const zip = await ZipArchive.read((await shape.build()).slice().buffer);
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  expect(xml.match(new RegExp(DRAWING_RUN, "gu"))).toHaveLength(1);
  const tracked = xml.replace(
    DRAWING_RUN,
    (_run, direct: string | undefined = "") =>
      `<w:r><w:rPr>${direct}<w:strike/><w:rPrChange w:id="${String(REVISION)}" w:author="Reviewer" w:date="2026-01-02T03:04:05Z"><w:rPr>${direct}</w:rPr></w:rPrChange></w:rPr><w:drawing>`,
  );
  zip.set("word/document.xml", tracked);
  return zip.generate();
};

const drawingRuns = (document: Document): Run[] =>
  document.package.document.content.flatMap((block) =>
    block.type === "paragraph"
      ? block.content.flatMap((item) =>
          item.type === "run" &&
          item.content.some(
            (content) => content.type === "drawing" || content.type === "shape",
          )
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

/** Resolve the change on the one node carrying it, by that node's range. */
const resolveAtom = (
  state: EditorState,
  decision: "accept" | "reject",
): EditorState => {
  let carrier: { from: number; to: number } | undefined;
  state.doc.descendants((node, pos) => {
    if (node.marks.some(({ type }) => type.name === "runPropertyChange")) {
      if (carrier !== undefined)
        throw new TypeError("Expected one node carrying the change.");
      carrier = { from: pos, to: pos + node.nodeSize };
    }
  });
  if (carrier === undefined) throw new TypeError("No node carries the change.");
  let resolved = state;
  const command =
    decision === "accept"
      ? acceptChange(carrier.from, carrier.to)
      : rejectChange(carrier.from, carrier.to);
  expect(
    command(state, (tr) => {
      resolved = resolved.apply(tr);
    }),
  ).toBe(true);
  return resolved;
};

/** The run's formatting as the change left it, accepted or rejected. */
const resolvedFormatting = (
  direct: TextFormatting,
  decision: "accept" | "reject",
): TextFormatting =>
  decision === "accept" ? { ...direct, strike: true } : direct;

type ExpectResolvedOptions = {
  run: Run;
  direct: TextFormatting;
  decision: "accept" | "reject";
};

const expectResolved = ({ run, direct, decision }: ExpectResolvedOptions) => {
  expect(run.propertyChanges ?? []).toEqual([]);
  expect(run.formatting ?? {}).toEqual(resolvedFormatting(direct, decision));
};

describe.each(CASES)(
  "tracked formatting on a run holding a $object",
  ({ shape, direct }) => {
    test("the change list shows it and reaches every revision", async () => {
      const document = await parseShapeDocument(await trackedPackage(shape));
      expect(
        drawingRun(document).propertyChanges?.map(({ info }) => info.id),
      ).toEqual([REVISION]);
      const doc = toProseDoc(document);
      expect(
        getTrackedChangesFromDoc(doc).map(({ id, type }) => ({ id, type })),
      ).toEqual([{ id: REVISION, type: "formatting" }]);
      expect(listedChangeCoverageProblem(document, doc)).toBeNull();
    });

    test("a save from the editor keeps it and what it replaced", async () => {
      const document = await parseShapeDocument(await trackedPackage(shape));
      const saved = await saveHarnessState(
        createHarnessState(document, "editing"),
        document,
      );
      const xml =
        (await (
          await ZipArchive.read(saved.bytes.slice().buffer)
        )
          .entry("word/document.xml")
          ?.text()) ?? "";
      expect(xml).toContain(`<w:rPrChange w:id="${String(REVISION)}"`);
      const reopened = drawingRun(await parseShapeDocument(saved.bytes));
      expect(reopened.formatting ?? {}).toEqual({ ...direct, strike: true });
      expect(
        reopened.propertyChanges?.map(({ info, previousFormatting }) => ({
          id: info.id,
          previousFormatting: previousFormatting ?? {},
        })),
      ).toEqual([{ id: REVISION, previousFormatting: direct }]);
    });

    for (const decision of ["accept", "reject"] as const) {
      test.each(["alone", "with every change"] as const)(
        `the editor can ${decision} it %s`,
        async (scope) => {
          const document = await parseShapeDocument(
            await trackedPackage(shape),
          );
          const state = createHarnessState(document, "editing");
          const resolved =
            scope === "alone"
              ? resolveAtom(state, decision)
              : resolveAllChanges(state, decision);
          expect(getTrackedChangesFromDoc(resolved.doc)).toEqual([]);
          const saved = await saveHarnessState(resolved, document);
          expectResolved({
            run: drawingRun(await parseShapeDocument(saved.bytes)),
            direct,
            decision,
          });
        },
      );

      test(`the reviewer can ${decision} it`, async () => {
        const reviewer = await FolioDocxReviewer.fromBuffer(
          (await trackedPackage(shape)).slice().buffer,
        );
        expect(reviewer.getChanges().map(({ id }) => id)).toEqual([REVISION]);
        expect(
          decision === "accept"
            ? reviewer.acceptChange(REVISION)
            : reviewer.rejectChange(REVISION),
        ).toBe(true);
        expect(reviewer.getChanges()).toEqual([]);
        expectResolved({
          run: drawingRun(
            await parseShapeDocument(new Uint8Array(await reviewer.toBuffer())),
          ),
          direct,
          decision,
        });
      });
    }
  },
);
