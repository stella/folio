/**
 * A list a suggestion started leaves no definition behind once rejected. The
 * session defines a list when a command starts one; the save writes only the
 * lists something still uses, plus every list the source package had, so a
 * rejected bullet saves as the source and an accepted one keeps its list.
 */
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { schema } from "../prosemirror/schema";
import type { Document } from "../types/document";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "./canonicalSession";
import { prepareCanonicalCommands } from "./canonicalStructure";

const TARGET = "Dear reader,";

/** The numbering instances a package defines, by id. */
const definedNumIds = async (buffer: ArrayBuffer): Promise<number[]> => {
  const xml =
    (await (
      await ZipArchive.read(buffer)
    )
      .entry("word/numbering.xml")
      ?.text()) ?? "";
  return [...xml.matchAll(/<w:num w:numId="(\d+)"/gu)].map(([, id]) =>
    Number(id),
  );
};

const sourceOf = async (markdown: string, unusedList: boolean) => {
  const document: Document = fromMarkdown(markdown);
  if (unusedList) {
    // A list the source defines and nothing uses, as authored files often have.
    const listed = fromMarkdown("1. Item").package.numbering;
    document.package.numbering = listed;
  }
  return createDocx(document);
};

const resolved = async (
  source: ArrayBuffer,
  resolution: "accept" | "reject",
) => {
  const session = createCanonicalSession(
    await parseDocx((await prepareCanonicalDocxInput(source)).unwrap(), {
      preloadFonts: false,
      detectVariables: false,
    }),
  ).unwrap();
  session.setMode({ type: "suggesting", author: "Reviewer" });
  let state = EditorState.create({ schema, doc: session.projection.doc });
  let position: number | undefined;
  state.doc.descendants((node, at) => {
    if (node.type.name === "paragraph" && node.textContent === TARGET)
      position = at;
  });
  if (position === undefined) throw new Error("no target paragraph");
  state = state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, position + 2)),
  );
  state = publishCanonicalProjection({
    session,
    state,
    commit: prepareCanonicalCommands(session, state, [
      { type: "toggleList", kind: "bullet" },
    ]).unwrap(),
  }).unwrap().state;
  const ids = [
    ...new Set(
      [
        ...JSON.stringify(session.document).matchAll(/"id":(\d+),"author"/gu),
      ].map(([, id]) => Number(id)),
    ),
  ];
  expect(ids.length).toBeGreaterThan(0);
  const commit = session
    .prepareResolve(state, { revisionIds: ids, resolution })
    .unwrap();
  state.apply(commit.transaction);
  expect(commit.publish().isOk()).toBe(true);
  return (
    await serializeCanonicalSave({ snapshot: session.captureSaveSnapshot() })
  ).buffer;
};

describe("a bullet list started while suggesting", () => {
  for (const unusedList of [false, true])
    describe(
      unusedList
        ? "beside a list the source defines and does not use"
        : "in a document without lists",
      () => {
        test("saves the source's numbering once rejected", async () => {
          const source = await sourceOf(`${TARGET}\n\nSecond`, unusedList);
          expect(await definedNumIds(await resolved(source, "reject"))).toEqual(
            await definedNumIds(source),
          );
        });
        test("keeps its list once accepted", async () => {
          const source = await sourceOf(`${TARGET}\n\nSecond`, unusedList);
          const accepted = await definedNumIds(
            await resolved(source, "accept"),
          );
          expect(accepted.length).toBe(
            (await definedNumIds(source)).length + 1,
          );
          const reopened = await parseDocx(await resolved(source, "accept"), {
            preloadFonts: false,
          });
          const target = reopened.package.document.content.find(
            (block) =>
              block.type === "paragraph" &&
              JSON.stringify(block.content).includes(TARGET),
          );
          const numPr =
            target?.type === "paragraph" ? target.formatting?.numPr : undefined;
          expect(numPr?.kind === "reference" ? accepted : []).toContain(
            numPr?.kind === "reference" ? numPr.numId : -1,
          );
        });
      },
    );
});
