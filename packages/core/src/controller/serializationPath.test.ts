import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { CANONICAL_SAVE_FALLBACK_DIAGNOSTIC } from "../../../../test/canonicalSaveDiagnostics";
import {
  SERIALIZATION_PATH_CASES,
  serializationPathSource,
} from "../../../../test/serializationPaths";
import { createCanonicalSession } from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { SaveDiagnostic } from "../docx/saveDiagnostics";
import type { FolioDocxSerializationPath } from "../types/docxSerialization";
import type { SelectiveSaveDecline } from "../types/selectiveSaveDecline";

test.each(SERIALIZATION_PATH_CASES)(
  "canonical $kind reports the path of the returned bytes once",
  async (scenario) => {
    const model = serializationPathSource();
    const source =
      scenario.source === "model"
        ? model
        : await parseDocx(await createDocx(model), { preloadFonts: false });
    const session = createCanonicalSession(source).unwrap();
    const state = EditorState.create({ doc: session.projection.doc });
    const edit = session
      .prepareReplace(state, { from: 2, to: 2, text: "x", semantic: "typing" })
      .unwrap();
    state.apply(edit.transaction);
    edit.publish().unwrap();
    const paths: FolioDocxSerializationPath[] = [];
    const diagnostics: SaveDiagnostic[] = [];
    const declines: SelectiveSaveDecline[] = [];
    const saved = await serializeCanonicalSave({
      snapshot: session.captureSaveSnapshot(),
      featureFlags: scenario.flags,
      options: {
        mode: scenario.mode,
        onDiagnostic: (diagnostic) => diagnostics.push(diagnostic),
        onSerializationPath: (path) => paths.push(path),
        onSelectiveDecline: (decline) => declines.push(decline),
      },
    });
    expect(paths).toEqual([scenario.expectedPath]);
    expect(saved.path).toBe(scenario.expectedPath);
    expect(diagnostics).toEqual(saved.diagnostics);
    expect(diagnostics).toEqual(
      scenario.kind === "declined" ? [CANONICAL_SAVE_FALLBACK_DIAGNOSTIC] : [],
    );
    expect(declines).toEqual(scenario.expectedDeclines);
    expect(saved.version).toBe(session.version);
    const reopened = await parseDocx(saved.buffer, { preloadFonts: false });
    expect(reopened.package.document.content.at(0)).toMatchObject({
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text: "Bxody" }] }],
    });
    expect(session.canUndo).toBe(true);
  },
);
