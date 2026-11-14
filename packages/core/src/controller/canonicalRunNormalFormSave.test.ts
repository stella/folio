import { panic } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import { EditorState } from "prosemirror-state";

import {
  DOCUMENT_OP_TYPES,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  paragraphLength,
  paragraphLogicalText,
  runSeams,
  storyParagraphs,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import { scanTestTimeout } from "../../../../test/scan-timeout";
import { createCanonicalSession } from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import type { SaveDiagnostic } from "../docx/saveDiagnostics";
import { schema } from "../prosemirror/schema";
import { FOLIO_DOCX_SERIALIZATION_MODE } from "../types/docxSerialization";

/*
 * Every edit leaves the paragraphs it touched in their run normal form, and
 * only those: in each corpus document, a character typed at the start of one
 * paragraph and deleted again leaves that paragraph in normal form, and the
 * selective save writes every other paragraph byte for byte as the source
 * held it.
 */

setDefaultTimeout(scanTestTimeout(120_000));

const CORPUS = path.join(
  import.meta.dir,
  "../docx/__tests__/__fixtures__/corpus",
);

/** Each paragraph's XML in a package's main part, by its paragraph id. */
const paragraphXml = async (
  bytes: ArrayBuffer | Uint8Array,
): Promise<Map<string, string>> => {
  const zip = await ZipArchive.read(bytes);
  const part =
    zip.entry("word/document.xml") ?? panic("Expected the main part.");
  const xml = await part.text();
  const out = new Map<string, string>();
  // Paragraphs that hold no nested paragraph: their XML is their own.
  for (const match of xml.matchAll(/<w:p\b(?:(?!<w:p\b)[\s\S])*?<\/w:p>/gu)) {
    const id = /w14:paraId="([0-9A-Fa-f]{8})"/u.exec(match[0])?.[1];
    if (id !== undefined) out.set(id.toUpperCase(), match[0]);
  }
  return out;
};

test("an edit leaves only the paragraph it touched rewritten, in run normal form", async () => {
  let compared = 0;
  for (const name of readdirSync(CORPUS).filter((file) =>
    file.endsWith(".docx"),
  )) {
    const prepared = await prepareCanonicalDocxInput(
      readFileSync(path.join(CORPUS, name)),
    );
    if (prepared.isErr()) continue;
    const source = prepared.value;
    const opened = createCanonicalSession(
      await parseDocx(source, { preloadFonts: false, detectVariables: false }),
    );
    if (opened.isErr()) continue;
    const session = opened.value;
    const target = storyParagraphs(session.document.package.document)
      .map(({ paragraph }) => paragraph)
      .find(
        (paragraph) =>
          paragraph.paraId !== undefined &&
          paragraphLength(paragraph) > 0 &&
          paragraph.content.at(0)?.type === "run",
      );
    const blockId = target?.paraId;
    if (target === undefined || blockId === undefined) continue;
    const text = paragraphLogicalText(target);
    let state = EditorState.create({ schema, doc: session.projection.doc });
    const at = { story: OP_STORIES.MAIN, blockId, offset: 0 };
    for (const ops of [
      [
        {
          type: DOCUMENT_OP_TYPES.INSERT_TEXT,
          at,
          text: "x",
          runProps: INHERIT_RUN_PROPS,
        },
      ],
      [
        {
          type: DOCUMENT_OP_TYPES.DELETE_RANGE,
          from: at,
          to: { ...at, offset: 1 },
        },
      ],
    ] as const) {
      const commit = session.prepareOperations(state, ops);
      if (commit.isErr()) break;
      state = state.apply(commit.value.transaction);
      commit.value.publish().unwrap();
    }
    const edited = storyParagraphs(session.document.package.document).find(
      ({ paragraph }) => paragraph.paraId === blockId,
    )?.paragraph;
    if (edited === undefined || paragraphLogicalText(edited) !== text) continue;
    expect({ name, seams: runSeams(edited) }).toEqual({ name, seams: [] });

    const diagnostics: SaveDiagnostic[] = [];
    const snapshot = session.captureSaveSnapshot();
    const saved = await serializeCanonicalSave({
      snapshot,
      featureFlags: { selectiveSave: true },
      options: {
        mode: FOLIO_DOCX_SERIALIZATION_MODE.preferSelective,
        onDiagnostic: (diagnostic) => diagnostics.push(diagnostic),
      },
    });
    if (diagnostics.some(({ type }) => type === "selectiveSaveRefused"))
      continue;
    const before = await paragraphXml(source);
    const after = await paragraphXml(saved.buffer);
    const changed = new Set(
      snapshot.changedBlockIds.map((id) => id.toUpperCase()),
    );
    expect({ name, changed: [...changed] }).toEqual({
      name,
      changed: [blockId.toUpperCase()],
    });
    for (const [id, xml] of before) {
      if (changed.has(id)) continue;
      expect({ name, id, xml: after.get(id) }).toEqual({ name, id, xml });
    }
    compared += 1;
  }
  // Most of the corpus takes the edit and saves selectively.
  expect(compared).toBeGreaterThan(10);
});
