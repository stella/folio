import { expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, type Command, type Transaction } from "prosemirror-state";
import { Step } from "prosemirror-transform";

import { buildDocumentPackage, DOCUMENT_CLASSES } from "../../../../benchmarks/compare/documents";
import { zipPackage } from "../../../../benchmarks/compare/package-xml";
import { applyVariant, EDIT_VARIANTS } from "../../../../benchmarks/compare/variants";
import { FolioDocxReviewer, type FolioDocumentStoryHandle } from "../ai-edits/headless";
import { parseDocx } from "../docx/parser";
import type { Document } from "../types/document";
import {
  acceptAllChanges,
  acceptChange,
  rejectAllChanges,
  rejectChange,
} from "../prosemirror/commands/comments";
import {
  footnoteToProseDoc,
  headerFooterToProseDoc,
  toProseDoc,
} from "../prosemirror/conversion/toProseDoc";
import { createDocumentNumberingPlugin } from "../prosemirror/plugins/documentNumbering";
import { createDocumentStylesPlugin } from "../prosemirror/plugins/documentStyles";
import { schema } from "../prosemirror/schema";
import { compareDocx } from "./compare";

const OPTIONS = {
  author: "folio compare benchmark",
  timestamp: "2000-01-01T00:00:00.000Z",
  onUnverified: "emit",
} as const;

const storyDoc = (document: Document, handle: FolioDocumentStoryHandle): PMNode => {
  const options = {
    ...(document.package.styles && { styles: document.package.styles }),
    ...(document.package.theme && { theme: document.package.theme }),
  };
  switch (handle.type) {
    case "main":
      return toProseDoc(document, options);
    case "header":
    case "footer": {
      const part = (
        handle.type === "header" ? document.package.headers : document.package.footers
      )?.get(handle.relationshipId);
      if (!part) throw new Error(`Missing ${handle.type} ${handle.relationshipId}`);
      return headerFooterToProseDoc(part.content, options);
    }
    case "footnote":
    case "endnote": {
      const note = (
        handle.type === "footnote" ? document.package.footnotes : document.package.endnotes
      )?.find(({ id }) => id === handle.noteId);
      if (!note) throw new Error(`Missing ${handle.type} ${String(handle.noteId)}`);
      return footnoteToProseDoc(note.content, options);
    }
    default: {
      handle satisfies never;
      throw new Error("Unhandled document story");
    }
  }
};

const run = (state: EditorState, command: Command) => {
  let transaction: Transaction | null = null;
  const handled = command(state, (dispatched) => {
    transaction = dispatched;
  });
  return { handled, transaction, doc: transaction?.doc ?? state.doc };
};

test("generated compare corpus resolves every parsed story like the full-range commands", async () => {
  let configurations = 0;
  let stories = 0;
  let changedStories = 0;
  for (const documentClass of DOCUMENT_CLASSES) {
    for (const variant of EDIT_VARIANTS) {
      const parts = buildDocumentPackage({ documentClass, size: "s" });
      const targetParts = applyVariant({ parts, variant });
      if (!targetParts) continue;
      const id = `${documentClass}/s/${variant}`;
      let compared;
      try {
        compared = await compareDocx(
          await zipPackage(parts),
          await zipPackage(targetParts),
          OPTIONS,
        );
      } catch (error) {
        throw new Error(`${id}: comparison threw`, { cause: error });
      }
      if (compared.isErr()) throw new Error(`${id}: ${compared.error.message}`);
      const parsed = await parseDocx(compared.value.buffer, {
        detectVariables: false,
        preloadFonts: false,
      });
      const reviewer = await FolioDocxReviewer.fromBuffer(compared.value.buffer);
      configurations += 1;
      for (const { handle } of reviewer.listStories()) {
        const doc = storyDoc(parsed, handle);
        const state = EditorState.create({
          schema,
          doc,
          plugins: [
            createDocumentStylesPlugin(parsed.package.styles),
            createDocumentNumberingPlugin(parsed.package.numbering),
          ],
        });
        stories += 1;
        for (const mode of ["accept", "reject"] as const) {
          const bulk = run(state, mode === "accept" ? acceptAllChanges() : rejectAllChanges());
          const full = run(
            state,
            mode === "accept"
              ? acceptChange(0, doc.content.size)
              : rejectChange(0, doc.content.size),
          );
          const label = `${id} ${JSON.stringify(handle)} ${mode}`;
          if (
            bulk.handled !== full.handled ||
            JSON.stringify(bulk.doc.toJSON()) !== JSON.stringify(full.doc.toJSON())
          ) {
            throw new Error(`${label}: bulk resolution differs from full-range resolution`);
          }
          if (!bulk.transaction) continue;
          changedStories += 1;
          if (bulk.transaction.steps.length !== 1) {
            throw new Error(
              `${label}: bulk resolution used ${String(bulk.transaction.steps.length)} steps`,
            );
          }
          const step = bulk.transaction.steps.at(0);
          if (!step) throw new Error(`${label}: missing resolution step`);
          const replay = Step.fromJSON(schema, step.toJSON()).apply(doc);
          if (
            !replay.doc ||
            JSON.stringify(replay.doc.toJSON()) !== JSON.stringify(bulk.doc.toJSON())
          ) {
            throw new Error(
              `${label}: serialized resolution step differs from cached result: ${replay.failed ?? "document mismatch"}`,
            );
          }
        }
      }
    }
  }
  expect(configurations).toBeGreaterThan(0);
  expect(stories).toBeGreaterThan(configurations);
  expect(changedStories).toBeGreaterThan(0);
}, 120_000);
