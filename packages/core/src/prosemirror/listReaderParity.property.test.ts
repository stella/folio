import { expect, test } from "bun:test";
import { Fragment } from "prosemirror-model";
import { schema } from "./schema";
import { resolveNumberedRefFields } from "./numberedRefFields";
import fc from "fast-check";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";
import { paragraphNumberingReference } from "@stll/docx-core/model";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document, Paragraph, ParagraphFormatting } from "../types/document";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { newContext } from "../markdown/internals";
import { renderParagraph } from "../markdown/renderParagraph";
import { convertParagraphAttrs } from "../layout-bridge/convert/paragraphAttrs";
import { toFlowBlocks } from "../layout-bridge/convert/toFlowBlocks";
import { expectParagraphAttrs } from "./attrs";
import type { ParagraphAttrs } from "./schema/nodes";
import { toProseDoc } from "./conversion/toProseDoc";
import { fromProseDoc } from "./conversion/fromProseDoc";
import { createListLabelCounter, listLabelAttrsFromRendering } from "./listLabels";
import {
  advanceVisibleListMarker,
  cloneListCounterState,
  createListCounterState,
} from "./listMarker";
import { paragraphNumberingSourcesFromBoundary } from "./numberingAttr";

// Each production import of a counting owner belongs to an exercised reader.
// Discovery below also catches aliases and new files, so this registry cannot
// silently omit another counter producer.
const PRODUCER_READERS = {
  "packages/core/src/prosemirror/listMarker.ts::advanceVisibleListMarker": "pm",
  "packages/core/src/prosemirror/listLabels.ts::listLabelAttrsFromRendering": "model",
  "packages/core/src/prosemirror/listLabels.ts::createListLabelCounter": "snapshot",
  "packages/core/src/markdown/renderParagraph.ts::renderParagraphBlock": "markdown",
  "packages/core/src/markdown/internals.ts::newContext": "markdown",
  "packages/core/src/ai-edits/snapshot.ts::createFolioAIEditSnapshotInternal": "snapshot",
  "packages/core/src/layout-bridge/convert/paragraphAttrs.ts::convertParagraphAttrs": "layout",
  "packages/core/src/prosemirror/numberedRefFields.ts::numberTargetForParagraph": "refs",
} as const;
const COUNTER_OWNERS = new Set([
  "advanceListMarker",
  "advanceVisibleListMarker",
  "createListLabelCounter",
  "listLabelAttrsFromRendering",
]);
const producerFunctions = (source: string): string[] => {
  const file = ts.createSourceFile("producer.ts", source, ts.ScriptTarget.Latest, true);
  const owners = new Set<string>();
  const namespaces = new Set<string>();
  for (const statement of file.statements) {
    if (!ts.isImportDeclaration(statement) || !ts.isStringLiteral(statement.moduleSpecifier))
      continue;
    if (!/\/(?:listMarker|listLabels)(?:\.[cm]?ts|\.js)?$/u.test(statement.moduleSpecifier.text))
      continue;
    const bindings = statement.importClause?.namedBindings;
    if (!bindings) continue;
    if (ts.isNamespaceImport(bindings)) {
      namespaces.add(bindings.name.text);
      continue;
    }
    for (const binding of bindings.elements) {
      if (COUNTER_OWNERS.has((binding.propertyName ?? binding.name).text))
        owners.add(binding.name.text);
    }
  }
  const functions = new Set<string>();
  const visit = (node: ts.Node): void => {
    if (ts.isFunctionDeclaration(node) && node.name?.text === "advanceVisibleListMarker")
      functions.add(node.name.text);
    if (
      ts.isVariableDeclaration(node) &&
      ts.isIdentifier(node.name) &&
      node.name.text === "listLabelAttrsFromRendering"
    )
      functions.add(node.name.text);
    if (
      ts.isCallExpression(node) &&
      ((ts.isIdentifier(node.expression) && owners.has(node.expression.text)) ||
        (ts.isPropertyAccessExpression(node.expression) &&
          ts.isIdentifier(node.expression.expression) &&
          namespaces.has(node.expression.expression.text) &&
          COUNTER_OWNERS.has(node.expression.name.text)))
    ) {
      let ancestor: ts.Node | undefined = node.parent;
      while (ancestor) {
        if (ts.isFunctionDeclaration(ancestor) && ancestor.name) {
          functions.add(ancestor.name.text);
          break;
        }
        if (
          ts.isVariableDeclaration(ancestor) &&
          ts.isIdentifier(ancestor.name) &&
          ancestor.initializer &&
          (ts.isArrowFunction(ancestor.initializer) ||
            ts.isFunctionExpression(ancestor.initializer))
        ) {
          functions.add(ancestor.name.text);
          break;
        }
        ancestor = ancestor.parent;
      }
      if (!ancestor) functions.add("<module>");
    }
    ts.forEachChild(node, visit);
  };
  visit(file);
  return [...functions];
};
const discoveredProducers = (sources: ReadonlyMap<string, string>): string[] =>
  [...sources].flatMap(([file, source]) =>
    producerFunctions(source).map((name) => `${file}::${name}`),
  );
const unregisteredProducers = (sources: ReadonlyMap<string, string>): string[] =>
  discoveredProducers(sources).filter((producer) => !Object.hasOwn(PRODUCER_READERS, producer));

const ROOT = path.resolve(import.meta.dir, "../../../..");
test("counter producer registry rejects a new aliased reader in an existing module", () => {
  expect(
    unregisteredProducers(
      new Map([
        [
          "packages/core/src/prosemirror/listLabels.ts",
          "import { advanceVisibleListMarker as advance } from './listMarker'; export const newReader = () => advance(attrs, streams);",
        ],
      ]),
    ),
  ).toEqual(["packages/core/src/prosemirror/listLabels.ts::newReader"]);
});
test("every production counter producer belongs to an exercised reader", () => {
  const files = execFileSync("git", ["ls-files", "packages/core/src"], { cwd: ROOT })
    .toString()
    .trim()
    .split("\n")
    .filter((file) => file.endsWith(".ts") && !file.endsWith(".test.ts"));
  const sources = new Map(files.map((file) => [file, readFileSync(path.join(ROOT, file), "utf8")]));
  expect(unregisteredProducers(sources)).toEqual([]);
  expect(discoveredProducers(sources).sort()).toEqual(Object.keys(PRODUCER_READERS).sort());
});

const streams = () => ({ final: createListCounterState(), original: createListCounterState() });
const stateSnapshot = (state: ReturnType<typeof streams>) => ({
  final: cloneListCounterState(state.final),
  original: cloneListCounterState(state.original),
});
const INFO = { id: 1636, author: "Reviewer", date: "2026-01-01T00:00:00Z" };
type Source = "absent" | "levelOnly" | "reference" | "none";
const sourceFormatting = (source: Source, numId: number, ilvl: number): ParagraphFormatting => {
  const overrides = {
    absent: undefined,
    reference: paragraphNumberingReference({ numId, ilvl }),
    levelOnly: { kind: "levelOnly", ilvl },
    none: { kind: "none" },
  } as const;
  const numPr = overrides[source];
  return { styleId: `Numbered${numId}`, ...(numPr === undefined ? {} : { numPr }) };
};
const paragraph = (formatting?: ParagraphFormatting): Paragraph => ({
  type: "paragraph",
  ...(formatting === undefined ? {} : { formatting }),
  content: [{ type: "run", content: [{ type: "text", text: "item" }] }],
});

// Explicitly enumerate source states, transition directions, and both separator
// positions in every generated document family rather than hoping random draws
// eventually reach a reader's distinct boundary semantics.
const documents = ({
  numId,
  abstractNumId,
  start,
  level,
}: {
  numId: number;
  abstractNumId: number;
  start: number;
  level: number;
}): Document[] => {
  const result: Document[] = [];
  for (const source of ["absent", "levelOnly", "reference", "none"] as const) {
    for (const direction of ["styleToDirect", "directToStyle"] as const) {
      for (const before of [false, true])
        for (const after of [false, true]) {
          const inherited = sourceFormatting(source, numId, level);
          const direct = sourceFormatting("reference", numId + 1, level);
          const original = direction === "styleToDirect" ? inherited : direct;
          const current = direction === "styleToDirect" ? direct : inherited;
          const changed = paragraph(current);
          changed.propertyChanges = [
            { type: "paragraphPropertyChange", info: INFO, previousFormatting: original },
          ];
          const deleted = paragraph(direct);
          deleted.pPrMark = { kind: "del", info: INFO };
          result.push({
            package: {
              document: {
                content: [
                  paragraph(sourceFormatting("absent", numId, level)),
                  ...(before ? [paragraph()] : []),
                  changed,
                  ...(after ? [paragraph()] : []),
                  deleted,
                  paragraph(inherited),
                  paragraph(direct),
                  paragraph(sourceFormatting("absent", numId, level)),
                ],
              },
              styles: {
                styles: [numId, numId + 1].map((id) => ({
                  type: "paragraph",
                  styleId: `Numbered${id}`,
                  pPr: { numPr: paragraphNumberingReference({ numId: id, ilvl: level }) },
                })),
              },
              numbering: {
                nums: [
                  { numId, abstractNumId },
                  {
                    numId: numId + 1,
                    abstractNumId,
                    levelOverrides: [{ ilvl: level, startOverride: start }],
                  },
                ],
                abstractNums: [
                  {
                    abstractNumId,
                    levels: Array.from({ length: 9 }, (_, ilvl) => ({
                      ilvl,
                      numFmt: "decimal",
                      lvlText: `%${ilvl + 1}.`,
                      start: 1,
                    })),
                  },
                ],
              },
            },
          });
        }
    }
  }
  return result;
};

test(
  "all numbering readers preserve labels and complete counter streams",
  () => {
    assertProperty(
      fc.property(
        fc.record({
          numId: fc.integer({ min: 1, max: 500 }),
          abstractNumId: fc.integer({ min: 1, max: 500 }),
          start: fc.integer({ min: 2, max: 20 }),
          level: fc.integer({ min: 0, max: 2 }),
        }),
        (options) => {
          const exercised = new Set<string>();
          for (const document of documents(options)) {
            const pm = toProseDoc(document);
            const converted = fromProseDoc(pm, document, { stylesheetSource: { type: "package" } });
            const owner = streams();
            const snapshot = streams();
            const model = streams();
            const markdown = streams();
            const layout = streams();
            const nextSnapshot = createListLabelCounter(snapshot);
            const nextModel = createListLabelCounter(model);
            const ctx = newContext({}, document.package.styles);
            let markdownLabel: string | undefined;
            const nextMarkdown = createListLabelCounter(markdown);
            ctx.nextListLabel = (attrs) => {
              markdownLabel = nextMarkdown(attrs);
              return markdownLabel;
            };
            const labels: (string | undefined)[] = [];
            pm.forEach((node, _offset, index) => {
              const attrs = expectParagraphAttrs(node);
              const para = converted.package.document.content[index];
              if (para?.type !== "paragraph")
                throw new Error("Generated reader fixture must contain paragraphs");
              const expected = advanceVisibleListMarker(attrs, owner).marker ?? undefined;
              labels.push(expected);
              expect(nextSnapshot(attrs), `snapshot label at paragraph ${index}`).toEqual(expected);
              const adapted = listLabelAttrsFromRendering(
                para.listRendering,
                paragraphNumberingSourcesFromBoundary({
                  numPr: para.formatting?.numPr,
                  numPrFromStyle:
                    ctx.paragraphNumbering(para.formatting?.styleId) ??
                    para.formatting?.numPrFromStyle,
                }),
              );
              const modelAttrs = {
                ...adapted,
                pPrMark: attrs.pPrMark,
                _propertyChanges: attrs._propertyChanges,
              };
              expect(nextModel(modelAttrs), `model label at paragraph ${index}`).toEqual(expected);
              const renderedMarkdown = renderParagraph(ctx, converted.package, para);
              expect(markdownLabel, `Markdown label at paragraph ${index}`).toEqual(expected);
              expect(renderedMarkdown.trimStart(), `Markdown text at paragraph ${index}`).toBe(
                expected === undefined ? "item" : `${expected} item`,
              );
              const flow = convertParagraphAttrs(attrs, {
                theme: undefined,
                fontAlternates: undefined,
                listCounterStreams: layout,
                defaultTabStopTwips: undefined,
                paragraphMarkFormatting: () => undefined,
              });
              expect(flow.listMarker, `layout label at paragraph ${index}`).toEqual(expected);
              for (const [reader, actual] of Object.entries({ snapshot, model, markdown, layout }))
                expect(
                  stateSnapshot(actual),
                  `${reader} counter streams at paragraph ${index}`,
                ).toEqual(stateSnapshot(owner));
            });
            expect(
              createFolioAIEditSnapshot(pm).blocks.map((block) =>
                block.kind === "diagnostic" ? undefined : block.displayLabel,
              ),
            ).toEqual(labels);
            expect(
              toFlowBlocks(pm)
                .filter((block) => block.kind === "paragraph")
                .map((block) => block.attrs?.listMarker),
            ).toEqual(labels);
            const field = schema.node("field", {
              fieldType: "REF",
              instruction: "REF counter-target \\n",
              displayText: "",
            });
            const first = pm.child(0);
            const firstWithRef = first.type.create(
              {
                ...first.attrs,
                bookmarks: [{ id: 1636, name: "counter-target" }] satisfies NonNullable<
                  ParagraphAttrs["bookmarks"]
                >,
              },
              first.content.append(Fragment.from(field)),
              first.marks,
            );
            const refDocument = pm.type.create(pm.attrs, [
              firstWithRef,
              ...Array.from({ length: pm.childCount - 1 }, (_, index) => pm.child(index + 1)),
            ]);
            const refs = streams();
            expect(
              resolveNumberedRefFields(refDocument, {
                listCounterState: refs.final,
                originalListCounterState: refs.original,
              }).get(field),
            ).toBe(labels.at(0)?.replace(/\.$/u, ""));
            expect(stateSnapshot(refs)).toEqual(stateSnapshot(owner));
            for (const reader of ["snapshot", "model", "markdown", "layout", "pm", "refs"])
              exercised.add(reader);
          }
          expect([...new Set(Object.values(PRODUCER_READERS))].sort()).toEqual(
            [...exercised].sort(),
          );
        },
      ),
      { numRuns: 8, id: "all numbering readers preserve labels and complete counter streams" },
    );
  },
  propertyTestTimeout(30_000),
);
