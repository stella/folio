import { panic } from "better-result";
import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import {
  type BlockContent,
  type Document,
  type Paragraph,
  PARAGRAPH_MARK_CHANGE_KINDS,
  type ParagraphMarkChangeKind,
} from "@stll/docx-core/model";
import {
  type AppliedDocumentOp,
  applyDocumentOp,
  applyDocumentOpEnvelope,
  applyDocumentOps,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  DOCUMENT_OP_REFUSAL_REASONS,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type EditorIntent,
  normalizeForOps,
  OP_STORIES,
  storyParagraphs,
  toOpEnvelope,
  validateOpsDocument,
} from "@stll/docx-core/ops";

import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { createCanonicalSession } from "./canonicalSession";

const CONTAINERS = ["body", "cell", "sdt", "customXml"] as const;
type Container = (typeof CONTAINERS)[number];
type Placement = "terminal" | "interior" | "outer-following";
const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});
const wrap = (
  content: BlockContent[],
  container: Container,
): BlockContent[] => {
  switch (container) {
    case "body":
      return content;
    case "cell":
      return [
        {
          type: "table",
          rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }],
        },
        paragraph("10000004", "Outside"),
      ];
    case "sdt":
      return [
        {
          type: "blockSdt",
          properties: { id: 77, sdtType: "richText" },
          content,
        },
      ];
    case "customXml":
      return [
        {
          type: "blockCustomXml",
          openingXml: '<w:customXml w:element="clause">',
          closingXml: "</w:customXml>",
          content,
        },
      ];
    default: {
      const unreachable: never = container;
      return unreachable;
    }
  }
};
type ImportedSourceOptions = {
  container: Container;
  placement: Placement;
  mark?: ParagraphMarkChangeKind;
};
const importedSource = async ({
  container,
  placement,
  mark,
}: ImportedSourceOptions) => {
  const content = [
    paragraph("10000001", "Kept"),
    {
      ...paragraph("10000002", ""),
      ...(mark === undefined
        ? {}
        : {
            pPrMark: {
              kind: mark,
              info: { id: 7, author: "Reviewer", date: "2026-01-02T03:04:05Z" },
            },
          }),
    },
  ];
  if (placement === "interior")
    content.push(paragraph("10000003", "Following"));
  const input = fromMarkdown("Seed");
  input.package.document.content = wrap(content, container);
  if (placement === "outer-following")
    input.package.document.content.push(
      paragraph("10000003", "Following outside"),
    );
  const parsed = normalizeForOps(await parseDocx(await createDocx(input)));
  expect(validateOpsDocument(parsed).isOk()).toBe(true);
  // The operation wire carries portable fields, not core's private source
  // capability Symbols. Keep their ownership at the canonical-session boundary.
  const source = normalizeForOps(structuredClone(parsed));
  const captured =
    storyParagraphs(source.package.document).find(
      ({ paragraph: item }) => item.paraId === "10000002",
    )?.paragraph ??
    panic("The actual parsed source must preserve its target paragraph.");
  expect(captured.pPrMark?.kind).toBe(mark);
  return { source, captured, parsed };
};
const removal = (captured: Paragraph) =>
  ({
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: OP_STORIES.MAIN,
    expected: [captured],
    blocks: [],
  }) as const;
const exactUndoRedo = (source: Document, edited: AppliedDocumentOp) => {
  let restored = edited.document;
  const redo: DocumentOp[] = [];
  for (const envelope of JSON.parse(
    JSON.stringify(edited.inverse.map(toOpEnvelope)),
  )) {
    const undo = applyDocumentOpEnvelope(restored, envelope).unwrap();
    restored = undo.document;
    redo.unshift(...undo.inverse);
  }
  expect(restored).toStrictEqual(source);
  expect(applyDocumentOps(restored, redo).unwrap().document).toStrictEqual(
    edited.document,
  );
};

const exactRoundTrip = ({
  source,
  captured,
}: Awaited<ReturnType<typeof importedSource>>) =>
  exactUndoRedo(source, applyDocumentOp(source, removal(captured)).unwrap());

for (const container of CONTAINERS)
  for (const mark of PARAGRAPH_MARK_CHANGE_KINDS) {
    test(`${container}/${mark}: empty removal refuses an imported container-final mark atomically`, async () => {
      const { source, captured } = await importedSource({
        container,
        placement: "terminal",
        mark,
      });
      const original = structuredClone(source);
      const refused = applyDocumentOp(source, removal(captured));
      expect(refused.isErr()).toBe(true);
      if (refused.isOk())
        panic(
          "The source-final mark cannot be restored by the ordinary insertion inverse.",
        );
      expect(refused.error.reason).toBe(
        DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
      );
      expect(source).toStrictEqual(original);
      const changed = applyDocumentOp(source, {
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
        story: OP_STORIES.MAIN,
        blockId: "10000002",
        patch: { alignment: "right" },
      }).unwrap();
      const stale = applyDocumentOp(changed.document, removal(captured));
      expect(stale.isErr()).toBe(true);
      if (stale.isOk())
        panic("A changed captured paragraph must refuse as stale first.");
      expect(stale.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
    });
    test(`${container}/${mark}: an imported interior mark keeps its exact JSON inverse and redo`, async () => {
      exactRoundTrip(
        await importedSource({ container, placement: "interior", mark }),
      );
    });
  }
for (const container of CONTAINERS)
  test(`${container}: an imported unmarked final paragraph keeps exact JSON inverse and redo`, async () => {
    exactRoundTrip(await importedSource({ container, placement: "terminal" }));
  });
for (const container of ["sdt", "customXml"] as const)
  for (const mark of PARAGRAPH_MARK_CHANGE_KINDS)
    test(`${container}/${mark}: a wrapper-final mark followed in the outer list is not container-final`, async () => {
      exactRoundTrip(
        await importedSource({ container, placement: "outer-following", mark }),
      );
    });

test("a refused imported-final-mark removal leaves the canonical session unpublished", async () => {
  const { parsed, captured } = await importedSource({
    container: "body",
    placement: "terminal",
    mark: "ins",
  });
  const session = createCanonicalSession(parsed).unwrap();
  const state = EditorState.create({ doc: session.projection.doc });
  const before = session.document;
  const snapshot = session.captureSaveSnapshot();
  let events = 0;
  const unsubscribe = session.onCommitted(() => events++);
  const position = {
    story: OP_STORIES.MAIN,
    blockId: "10000001",
    offset: 0,
  } as const;
  const refused = session.prepareOps(state, [removal(captured)], {
    type: "text",
    anchor: position,
    head: position,
  });
  expect(refused.isErr()).toBe(true);
  expect(session.document).toBe(before);
  expect(session.version).toBe(0);
  expect(session.captureSaveSnapshot()).toStrictEqual(snapshot);
  expect(events).toBe(0);
  expect(session.prepareUndo(state, OP_STORIES.MAIN).isErr()).toBe(true);
  unsubscribe();
});

const directContexts = CONTAINERS.flatMap((container) =>
  (container === "cell"
    ? (["interior"] as const)
    : (["terminal", "interior"] as const)
  ).map((placement) => ({ container, placement })),
);
for (const { container, placement } of directContexts)
  for (const mark of PARAGRAPH_MARK_CHANGE_KINDS)
    test(`${container}/${mark}/${placement}: new direct table-side insertion uses the ordinary final-mark admission owner`, async () => {
      const input = fromMarkdown("Seed");
      const content: BlockContent[] = [
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: [
                { type: "tableCell", content: [paragraph("30000011", "Cell")] },
              ],
            },
          ],
        },
      ];
      if (placement === "interior")
        content.push(paragraph("30000002", "Following"));
      input.package.document.content = wrap(content, container);
      const source = normalizeForOps(
        structuredClone(await parseDocx(await createDocx(input))),
      );
      expect(validateOpsDocument(source).isOk()).toBe(true);
      const { captured } = await importedSource({
        container: "body",
        placement: "terminal",
        mark,
      });
      const intent = {
        type: "insertParagraphs",
        story: OP_STORIES.MAIN,
        at: { type: "afterTable", blockId: "30000011" },
        paragraphs: [captured],
      } as const satisfies EditorIntent;
      const ids = createEditorIntentIdAllocator()(source, intent);
      const compiled = compileEditorIntent(source, {
        intent,
        mode: { type: "editing", newIds: ids.newIds },
      }).unwrap();
      const original = structuredClone(source);
      const edited = applyDocumentOps(source, compiled.ops);
      if (placement === "terminal") {
        expect(edited.isErr()).toBe(true);
        if (edited.isOk())
          panic(
            "The new table-side path must preserve ordinary final-mark insertion admission.",
          );
        expect(edited.error.reason).toBe(
          DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
        );
      } else exactUndoRedo(source, edited.unwrap());
      expect(source).toStrictEqual(original);
    });
