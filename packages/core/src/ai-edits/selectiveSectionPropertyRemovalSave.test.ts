import { expect, test } from "bun:test";

import { createDocx, repackDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import type { Document, Paragraph } from "../types/document";
import { EditorState, type Transaction } from "prosemirror-state";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { rejectAIEditRevision } from "../prosemirror/commands/comments";
import {
  ParagraphChangeTrackerExtension,
  repackWithEditorSectionRemovals,
} from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { schema } from "../prosemirror/schema";

const REVISION = { id: 42, author: "Reviewer", date: "2026-05-15T12:00:00Z" };
const headerId = "rIdHeader";
const footerId = "rIdFooter";

const sectionParagraph = (pPrMark: Paragraph["pPrMark"]): Paragraph => ({
  type: "paragraph",
  paraId: "70000001",
  content: [{ type: "run", content: [{ type: "text", text: "Section end" }] }],
  ...(pPrMark ? { pPrMark } : {}),
  sectionProperties: {
    pageWidth: 12_240,
    headerReferences: [{ type: "default", rId: headerId }],
    footerReferences: [{ type: "even", rId: footerId }],
    propertyChanges: [
      {
        type: "sectionPropertyChange",
        info: REVISION,
        previousProperties: {},
        previousReferences: {},
      },
    ],
  },
});

const sourceDocument = (pPrMark?: Paragraph["pPrMark"]): Document => ({
  package: {
    document: {
      content: [
        sectionParagraph(pPrMark),
        {
          type: "paragraph",
          paraId: "70000002",
          content: [{ type: "run", content: [{ type: "text", text: "Following section" }] }],
        },
      ],
      finalSectionProperties: { pageWidth: 12_240 },
    },
    headers: new Map([
      [
        headerId,
        {
          type: "header",
          hdrFtrType: "default",
          content: [
            {
              type: "paragraph",
              content: [{ type: "run", content: [{ type: "text", text: "Header" }] }],
            },
          ],
        },
      ],
    ]),
    footers: new Map([
      [
        footerId,
        {
          type: "footer",
          hdrFtrType: "even",
          content: [
            {
              type: "paragraph",
              content: [{ type: "run", content: [{ type: "text", text: "Footer" }] }],
            },
          ],
        },
      ],
    ]),
  },
});

const tracker = ParagraphChangeTrackerExtension().onSchemaReady({ schema }).plugins?.at(0);
if (!tracker) throw new Error("Expected paragraph change tracker plugin");

const rejectAndReopen = async (document: Document): Promise<Document> => {
  const source = await createDocx(document, { updateModifiedDate: false });
  const parsed = await parseDocx(source, { preloadFonts: false });
  const view = {
    state: EditorState.create({
      schema,
      doc: toProseDoc(parsed),
      plugins: [tracker],
    }),
    dispatch(tr: Transaction) {
      view.state = view.state.apply(tr);
    },
  };
  expect(rejectAIEditRevision(REVISION.id)(view.state, view.dispatch)).toBe(true);
  const edited = fromProseDoc(view.state.doc, parsed, {
    stylesheetSource: { type: "package" },
  });
  const saved = await repackWithEditorSectionRemovals({
    state: view.state,
    document: edited,
    repack: () => repackDocx(edited, { updateModifiedDate: false }),
  });
  return await parseDocx(saved, { preloadFonts: false });
};

test("selective rejection can repack after clearing a referenced section endpoint", async () => {
  const reopened = await rejectAndReopen(sourceDocument());

  expect(reopened.package.document.content).toHaveLength(2);
  expect((reopened.package.document.content.at(0) as Paragraph).sectionProperties).toBeUndefined();
});

test("rejecting an empty inserted final paragraph moves its section to the preceding paragraph", async () => {
  const base = sourceDocument();
  const { propertyChanges: _propertyChanges, ...sectionProperties } =
    sectionParagraph(undefined).sectionProperties ?? {};
  const document: Document = {
    package: {
      ...base.package,
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "70000003",
            content: [{ type: "run", content: [{ type: "text", text: "Before" }] }],
          },
          {
            type: "paragraph",
            paraId: "70000004",
            content: [],
            pPrMark: { kind: "ins", info: REVISION },
            sectionProperties,
          },
        ],
        finalSectionProperties: { pageWidth: 12_240 },
      },
    },
  };

  const reopened = await rejectAndReopen(document);

  expect(reopened.package.document.content).toHaveLength(1);
  const survivor = reopened.package.document.content.at(0) as Paragraph;
  expect(survivor.paraId).toBe("70000003");
  expect(survivor.sectionProperties?.headerReferences).toEqual([
    { type: "default", rId: headerId },
  ]);
  expect(survivor.sectionProperties?.footerReferences).toEqual([{ type: "even", rId: footerId }]);
});

test("clearing properties and removing the same paragraph mark counts one endpoint", async () => {
  const reopened = await rejectAndReopen(sourceDocument({ kind: "ins", info: REVISION }));

  expect(reopened.package.document.content).toHaveLength(1);
  expect(reopened.package.document.content.at(0)?.type).toBe("paragraph");
});
