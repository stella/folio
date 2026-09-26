/**
 * The headless editor path an integrator drives: parse, `toProseDoc`, run a
 * command from `@stll/folio-core/prosemirror` with the plugin the editor
 * installs, `fromProseDoc`, `repackDocx`.
 */

import { repackDocx } from "@stll/folio-core/docx/rezip";
import { fromProseDoc, type toggleNumberedList, toProseDoc } from "@stll/folio-core/prosemirror";
import { createDocumentNumberingPlugin } from "@stll/folio-core/prosemirror/plugins/documentNumbering";
import { parseDocx } from "@stll/folio-core/server";
import { EditorState, TextSelection } from "prosemirror-state";

import { openReviewer, toArrayBuffer } from "./documents.ts";

export type Command = typeof toggleNumberedList;

/** Where the paragraph reading `text` starts. */
const paragraphPosition = (state: EditorState, text: string): number => {
  let position: number | null = null;
  state.doc.descendants((node, pos) => {
    if (position === null && node.type.name === "paragraph" && node.textContent === text) {
      position = pos;
    }
    return position === null;
  });
  if (position === null) {
    throw new Error(`no paragraph "${text}"`);
  }
  return position;
};

/** Run `command` with the caret inside the paragraph at `position`. */
const runAt = (state: EditorState, position: number, command: Command): EditorState => {
  let next = state.apply(state.tr.setSelection(TextSelection.create(state.doc, position + 1)));
  const before = next;
  command(before, (transaction) => {
    next = before.apply(transaction);
  });
  return next;
};

/** Run an editor list command with the caret in each named paragraph, then save. */
export const toggleAndSave = async (
  bytes: Uint8Array,
  texts: readonly string[],
  command: Command,
): Promise<Uint8Array> => {
  const document = await parseDocx(toArrayBuffer(bytes));
  let state = EditorState.create({
    doc: toProseDoc(document, {
      styles: document.package.styles,
      theme: document.package.theme,
    }),
    plugins: [createDocumentNumberingPlugin(document.package.numbering)],
  });
  for (const text of texts) {
    state = runAt(state, paragraphPosition(state, text), command);
  }
  return new Uint8Array(await repackDocx(fromProseDoc(state.doc, document)));
};

export const labelsOf = async (bytes: Uint8Array): Promise<string[]> =>
  (await openReviewer(bytes))
    .getContent()
    .map((block) => `${block.displayLabel ?? "·"} ${block.text}`);
