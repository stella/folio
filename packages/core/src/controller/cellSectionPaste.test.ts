/**
 * A paste or a drop puts no section break inside a table cell: the text comes,
 * the break it carried does not. Pasted between paragraphs outside a table it
 * still ends a section.
 */

import { panic } from "better-result";
import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { DOMParser, type Node as PMNode } from "prosemirror-model";
import { NodeSelection, TextSelection } from "prosemirror-state";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import {
  createHarnessState,
  EDITOR_MODES,
  findTextblock,
  HeadlessEditorView,
  parseShapeDocument,
  saveHarnessState,
} from "../__tests__/editorHarness";
import { createDocx } from "../docx/rezip";
import { singletonManager } from "../prosemirror/schema";
import { sectionPropertiesOf } from "../prosemirror/sectionCarrier";
import type { Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import {
  createHiddenEditorManager,
  type HiddenEditorManagerDeps,
} from "./hiddenEditorManager";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const source = (): Document => {
  const document = createEmptyDocument();
  document.package.document.content = [
    paragraph("74000001", "Outside"),
    {
      ...paragraph("74000002", "Ending"),
      sectionProperties: { pageWidth: 12_240 },
    },
    {
      type: "table",
      columnWidths: [2400],
      rows: [
        {
          type: "tableRow",
          cells: [
            { type: "tableCell", content: [paragraph("74000003", "Inside")] },
          ],
        },
      ],
    },
    paragraph("74000004", "Last"),
  ];
  return document;
};

// The paragraph with the break is the middle one: the outer two run on into
// the paragraph the paste lands in.
const BROKEN =
  '<p>First</p><p data-section-break="nextPage">Pasted</p><p>Third</p>';
const TABLE =
  '<table><tbody><tr><td><p data-section-break="nextPage">Pasted</p></td></tr></tbody></table>';

/** The texts of the paragraphs that end a section, in cells and outside them. */
const sections = (doc: PMNode) => {
  const inside: string[] = [];
  const outside: string[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") return true;
    if (sectionPropertiesOf(node) !== null) {
      const $pos = doc.resolve(pos);
      const inCell = Array.from({ length: $pos.depth }, (_, depth) =>
        $pos.node(depth + 1),
      ).some((holder) => holder.type.spec["tableRole"] === "cell");
      (inCell ? inside : outside).push(node.textContent);
    }
    return false;
  });
  return { inside, outside };
};

const caretIn = (doc: PMNode, text: string): TextSelection => {
  const found = findTextblock(doc, text) ?? panic(`No ${text}.`);
  return TextSelection.create(doc, found.pos + found.node.nodeSize - 1);
};

let releaseDom: (() => void) | undefined;
beforeAll(() => {
  releaseDom = acquireHarnessDom();
});
afterAll(() => releaseDom?.());

for (const session of ["prosemirror", "canonical"] as const)
  describe(`${session}: pasted HTML`, () => {
    const paste = (text: string, html: string): PMNode => {
      const host = document.createElement("div");
      document.body.append(host);
      const opened = source();
      const refusals: string[] = [];
      const deps: HiddenEditorManagerDeps = {
        getHost: () => host,
        getDocument: () => opened,
        getStyles: () => null,
        getExtensionManager: () => singletonManager,
        getExternalPlugins: () => [],
        getCollaboration: () => undefined,
        getCollaborationModules: () => null,
        getPrecomputedInitialState: () => null,
        getReadOnly: () => false,
        getDocumentIdentity: () => "0",
        getDocumentContext: () => opened,
        getEmbeddedFontFamilyMap: () => undefined,
        getExperimentalSession: () => session,
        onSessionRefusal: (reason) => refusals.push(reason),
        onTransaction: () => {},
        onCanonicalCompositionLayout: () => {},
        onSelectionChange: () => {},
        onKeyDown: () => false,
        onCopy: () => {},
        onCut: () => {},
        onPaste: () => {},
        onReadOnlyEditAttempt: () => {},
        onEditorViewReady: () => {},
        onEditorViewDestroy: () => {},
        onRemoteSelectionsChange: () => {},
      };
      const manager = createHiddenEditorManager(deps);
      try {
        manager.ensureView();
        const view = manager.getView() ?? panic("The editor mounts.");
        view.dispatch(
          view.state.tr.setSelection(caretIn(view.state.doc, text)),
        );
        view.pasteHTML(html);
        // A pasted table is no canonical edit yet: the paste is refused whole.
        const refused = session === "canonical" && html === TABLE;
        expect(refusals).toHaveLength(refused ? 1 : 0);
        expect(view.state.doc.textContent.includes("Pasted")).toBe(!refused);
        return view.state.doc;
      } finally {
        manager.destroyView();
        host.remove();
      }
    };

    test("into a cell comes without its section break", () => {
      expect(sections(paste("Inside", BROKEN))).toEqual({
        inside: [],
        outside: ["Ending"],
      });
    });

    test("holding a table comes without its cells' section breaks", () => {
      expect(sections(paste("Last", TABLE))).toEqual({
        inside: [],
        outside: ["Ending"],
      });
    });

    test("outside every table keeps its section break", () => {
      expect(sections(paste("Last", BROKEN)).outside).toHaveLength(2);
    });
  });

for (const mode of EDITOR_MODES)
  describe(`the default editor, ${mode}`, () => {
    test("a paste into a cell comes without its section break", () => {
      const view = new HeadlessEditorView(createHarnessState(source(), mode));
      view.dispatch(
        view.state.tr.setSelection(caretIn(view.state.doc, "Inside")),
      );
      const host = document.createElement("div");
      host.innerHTML = BROKEN;
      view.paste(DOMParser.fromSchema(view.state.schema).parseSlice(host));
      expect(view.state.doc.textContent).toContain("Pasted");
      expect(sections(view.state.doc)).toEqual({
        inside: [],
        outside: ["Ending"],
      });
    });

    test.each([
      ["text copied", false, false],
      ["text moved", true, false],
      ["a paragraph copied", false, true],
      ["a paragraph moved", true, true],
    ])(
      "%s into a cell comes without its section break, and saves",
      async (_title, move, whole) => {
        const opened = await parseShapeDocument(
          new Uint8Array(await createDocx(source())),
        );
        const view = new HeadlessEditorView(createHarnessState(opened, mode));
        const ending =
          findTextblock(view.state.doc, "Ending") ?? panic("Ending.");
        const first =
          findTextblock(view.state.doc, "Outside") ?? panic("Outside.");
        view.dispatch(
          view.state.tr.setSelection(
            whole
              ? NodeSelection.create(view.state.doc, ending.pos)
              : TextSelection.create(
                  view.state.doc,
                  first.pos + 1,
                  ending.pos + ending.node.nodeSize - 1,
                ),
          ),
        );
        const target = caretIn(view.state.doc, "Inside").from;
        view.drag(target, { move });
        expect(sections(view.state.doc).inside).toEqual([]);
        const saved = await saveHarnessState(view.state, opened);
        expect(saved).toBeDefined();
      },
    );
  });
