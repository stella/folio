import { expect, test } from "bun:test";
import { Fragment } from "prosemirror-model";
import { schema } from "../prosemirror/schema";
import { resolveNumberedRefFields } from "../prosemirror/numberedRefFields";
import fc from "fast-check";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";
import { normalizeCanonicalListRendering } from "./canonicalListRendering";
import { computeListMarker, type ComputeListMarkerOptions } from "../docx/listMarkerComputation";
import { computeListRendering, getCachedNumberingMap } from "../docx/numberingParser";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import { paragraphNumberingReference } from "@stll/docx-core/model";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document, Paragraph, ParagraphFormatting } from "../types/document";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { newContext } from "../markdown/internals";
import { renderParagraph } from "../markdown/renderParagraph";
import { convertParagraphAttrs } from "../layout-bridge/convert/paragraphAttrs";
import { toFlowBlocks } from "../layout-bridge/convert/toFlowBlocks";
import { expectParagraphAttrs } from "../prosemirror/attrs";
import type { ParagraphAttrs } from "../prosemirror/schema/nodes";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { createListLabelCounter, listLabelAttrsFromRendering } from "../prosemirror/listLabels";
import {
  advanceVisibleListMarker,
  cloneListCounterState,
  createListCounterState,
} from "../prosemirror/listMarker";
import { paragraphNumberingSourcesFromBoundary } from "../prosemirror/numberingAttr";

// Each production import of a counting owner belongs to an exercised reader.
// Discovery below also catches aliases and new files, so this registry cannot
// silently omit another counter producer.
const PRODUCER_READERS = {
  "packages/core/src/prosemirror/listMarker.ts::advanceVisibleListMarker": "pm",
  "packages/core/src/docx/listMarkerComputation.ts::computeListMarker": "parser",
  "packages/core/src/docx/blockContentParser.ts::p": "parser",
  "packages/core/src/docx/listModelNormalization.ts::normalizeBody": "canonical",
  "packages/core/src/prosemirror/listLabels.ts::listLabelAttrsFromRendering": "model",
  "packages/core/src/prosemirror/listLabels.ts::createListLabelCounter": "snapshot",
  "packages/core/src/markdown/renderParagraph.ts::renderParagraphBlock": "markdown",
  "packages/core/src/markdown/internals.ts::newContext": "markdown",
  "packages/core/src/ai-edits/snapshot.ts::createFolioAIEditSnapshotInternal": "snapshot",
  "packages/core/src/layout-bridge/convert/paragraphAttrs.ts::convertParagraphAttrs": "layout",
  "packages/core/src/prosemirror/numberedRefFields.ts::numberTargetForParagraph": "refs",
} as const;
const COUNTER_OWNERS = new Set([
  "computeListMarker",
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
    if (
      !/\/(?:listMarker|listLabels|listMarkerComputation)(?:\.[cm]?ts|\.js)?$/u.test(
        statement.moduleSpecifier.text,
      )
    )
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
      ["listLabelAttrsFromRendering", "computeListMarker"].includes(node.name.text)
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
        if (
          ts.isPropertyAssignment(ancestor) &&
          ts.isIdentifier(ancestor.name) &&
          (ts.isArrowFunction(ancestor.initializer) ||
            ts.isFunctionExpression(ancestor.initializer))
        ) {
          functions.add(ancestor.name.text);
          break;
        }
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
test.each([
  ["listMarker", "advanceVisibleListMarker"],
  ["listLabels", "createListLabelCounter"],
  ["listMarkerComputation", "computeListMarker"],
])("counter registry rejects aliased and namespace readers from %s", (module, helper) => {
  for (const source of [
    `import { ${helper} as advance } from './${module}'; export const newReader = () => advance(attrs, streams);`,
    `import * as owner from './${module}'; export const newReader = () => owner.${helper}(attrs, streams);`,
  ]) {
    expect(unregisteredProducers(new Map([["packages/core/src/newReader.ts", source]]))).toEqual([
      "packages/core/src/newReader.ts::newReader",
    ]);
  }
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
const PLAIN_STYLE_ID = "Plain";
const plainFormatting = (source: Source, numId: number, ilvl: number): ParagraphFormatting => ({
  ...sourceFormatting(source, numId, ilvl),
  styleId: PLAIN_STYLE_ID,
});
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
    for (const direction of ["styleToDirect", "directToStyle", "directToPlain"] as const) {
      for (const before of [false, true])
        for (const after of [false, true]) {
          const inherited = sourceFormatting(source, numId, level);
          const direct = sourceFormatting("reference", numId + 1, level);
          const transitions = {
            styleToDirect: { original: inherited, current: direct },
            directToStyle: { original: direct, current: inherited },
            directToPlain: { original: direct, current: plainFormatting(source, numId, level) },
          } as const;
          const { original, current } = transitions[direction];
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
                styles: [
                  ...[numId, numId + 1].map((id) => ({
                    type: "paragraph",
                    styleId: `Numbered${id}`,
                    pPr: { numPr: paragraphNumberingReference({ numId: id, ilvl: level }) },
                  })),
                  { type: "paragraph", styleId: PLAIN_STYLE_ID },
                ],
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
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.record({
          numId: fc.integer({ min: 1, max: 500 }),
          abstractNumId: fc.integer({ min: 1, max: 500 }),
          start: fc.integer({ min: 2, max: 20 }),
          level: fc.integer({ min: 0, max: 2 }),
          definedLevels: fc.uniqueArray(fc.integer({ min: 0, max: 8 }), {
            minLength: 1,
            maxLength: 8,
          }),
          hiddenLevel: fc.integer({ min: 0, max: 8 }),
          templateDepth: fc.integer({ min: 1, max: 9 }),
        }),
        async (options) => {
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
              const visible = nextMarkdown(attrs);
              markdownLabel = visible.label;
              return visible;
            };
            const labels: (string | undefined)[] = [];
            pm.forEach((node, _offset, index) => {
              const attrs = expectParagraphAttrs(node);
              const para = converted.package.document.content[index];
              if (para?.type !== "paragraph")
                throw new Error("Generated reader fixture must contain paragraphs");
              const expected = advanceVisibleListMarker(attrs, owner).marker ?? undefined;
              labels.push(expected);
              expect(nextSnapshot(attrs).label, `snapshot label at paragraph ${index}`).toEqual(
                expected,
              );
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
              expect(nextModel(modelAttrs).label, `model label at paragraph ${index}`).toEqual(
                expected,
              );
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
            const snapshotBlocks = createFolioAIEditSnapshot(pm).blocks;
            expect(
              snapshotBlocks.map((block) =>
                block.kind === "diagnostic" ? undefined : block.displayLabel,
              ),
            ).toEqual(labels);
            const numbering = document.package.numbering;
            const definedLevels = new Set(
              (numbering?.nums ?? []).flatMap(({ numId: id, abstractNumId: abstractId }) =>
                (
                  numbering?.abstractNums.find((entry) => entry.abstractNumId === abstractId)
                    ?.levels ?? []
                ).map(({ ilvl }) => `${id}:${ilvl}`),
              ),
            );
            for (const [index, block] of snapshotBlocks.entries()) {
              if (block.kind === "diagnostic" || block.displayLabel === undefined) continue;
              const reference = block.listReference;
              expect(
                reference === undefined
                  ? undefined
                  : definedLevels.has(`${reference.numId}:${reference.level}`),
                `list reference beside label ${block.displayLabel} at paragraph ${index}`,
              ).toBe(true);
            }
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
          await assertSparseReaders(options);
          exercised.add("parser");
          exercised.add("canonical");
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

// Dense fixtures and owner-derived expectations missed undefined-level boundaries.
const assertSparseReaders = async ({
  numId,
  abstractNumId,
  definedLevels,
  hiddenLevel,
  templateDepth,
}: {
  numId: number;
  abstractNumId: number;
  definedLevels: readonly number[];
  hiddenLevel: number;
  templateDepth: number;
}) => {
  const numberingDefinitions = {
    nums: [{ numId, abstractNumId }],
    abstractNums: [
      {
        abstractNumId,
        levels: definedLevels.map((ilvl) => ({
          ilvl,
          numFmt: "decimal" as const,
          start: 1,
          lvlText:
            Array.from({ length: templateDepth }, (_, index) => `%${index + 1}`).join(".") + ".",
          ...(ilvl === hiddenLevel ? { rPr: { hidden: true } } : {}),
        })),
      },
    ],
  };
  const numbering = getCachedNumberingMap(numberingDefinitions);
  const parser: ComputeListMarkerOptions = {
    numbering,
    listCounters: new Map<number, number[]>(),
    abstractCounters: new Map<number, number[]>(),
    restartedNumIds: new Set<number>(),
    previousList: {
      abstractNumId: null,
      fromStyle: false,
      numId: null,
    },
  };
  const reader = streams();
  const next = createListLabelCounter(reader);
  const content: Paragraph[] = [];
  const expected: (string | undefined)[] = [];
  // Every defined level brackets every possible absent parent or child level.
  for (const defined of definedLevels)
    for (let absent = 0; absent <= 8; absent += 1) {
      for (const ilvl of [defined, absent, defined]) {
        const numPr = paragraphNumberingReference({ numId, ilvl });
        const rendering = computeListRendering(numPr, numbering) ?? undefined;
        const para = {
          ...paragraph({ numPr }),
          paraId: (content.length + 1).toString(16).padStart(8, "0"),
          ...(rendering === undefined ? {} : { listRendering: { ...rendering } }),
        };
        computeListMarker(para, parser);
        const label = para.listRendering?.markerHidden ? undefined : para.listRendering?.marker;
        expect(
          next(listLabelAttrsFromRendering(rendering, { numPr, numPrFromStyle: undefined })).label,
        ).toEqual(label);
        for (const stream of [reader.final, reader.original]) {
          expect(stream.counters).toEqual(parser.listCounters);
          expect(stream.abstractCounters).toEqual(parser.abstractCounters);
          expect(stream.previousList).toEqual(parser.previousList);
          expect(stream.restartedNumIds).toEqual(parser.restartedNumIds);
        }
        content.push(para);
        expected.push(label);
      }
    }
  const document = { package: { document: { content }, numbering: numberingDefinitions } };
  const labels = (doc: Document) =>
    doc.package.document.content.map((block) =>
      block.type === "paragraph" && !block.listRendering?.markerHidden
        ? block.listRendering?.marker
        : undefined,
    );
  expect(labels(normalizeCanonicalListRendering(document).document)).toEqual(expected);
  const parsed = await parseDocx(await createDocx(document), {
    detectVariables: false,
    preloadFonts: false,
  });
  expect(labels(parsed)).toEqual(expected);
  const pm = toProseDoc(parsed);
  const layoutBlocks = toFlowBlocks(pm).filter((block) => block.kind === "paragraph");
  expect(
    layoutBlocks.map(({ attrs }) => ({
      marker: attrs?.listMarker,
      hidden: attrs?.listMarkerHidden === true,
    })),
  ).toEqual(
    parsed.package.document.content.map((block) => ({
      marker: block.type === "paragraph" ? block.listRendering?.marker : undefined,
      hidden: block.type === "paragraph" && block.listRendering?.markerHidden === true,
    })),
  );
  // Layout retains a hidden marker as metadata; the painter consults this flag.
  expect(
    layoutBlocks.map(({ attrs }) => (attrs?.listMarkerHidden ? undefined : attrs?.listMarker)),
  ).toEqual(expected);
  const converted = fromProseDoc(pm, parsed, { stylesheetSource: { type: "package" } });
  expect(labels(converted)).toEqual(expected);
  const ctx = newContext({}, parsed.package.styles);
  ctx.nextListLabel = createListLabelCounter(numbering);
  converted.package.document.content.forEach((block, index) => {
    if (block.type !== "paragraph") throw new Error("Expected paragraph");
    const label = expected.at(index);
    expect(renderParagraph(ctx, converted.package, block).trimStart()).toBe(
      label === undefined ? "item" : `${label} item`,
    );
  });
  expect(
    createFolioAIEditSnapshot(pm).blocks.map((block) =>
      block.kind === "diagnostic" ? undefined : block.displayLabel,
    ),
  ).toEqual(expected);
};

test("preserves template punctuation across readers for unresolved child placeholders", async () => {
  const source = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "12345678",
            formatting: { numPr: { kind: "reference", numId: 2, ilvl: 0 } },
            content: [{ type: "run", content: [{ type: "text", text: "Parent" }] }],
          },
        ],
      },
      numbering: {
        nums: [{ numId: 2, abstractNumId: 2 }],
        abstractNums: [
          {
            abstractNumId: 2,
            levels: [
              {
                ilvl: 0,
                numFmt: "decimal",
                start: 1,
                lvlText: "%1.%2.",
              },
            ],
          },
        ],
      },
    },
  } satisfies Document;
  const parsed = await parseDocx(await createDocx(source), { preloadFonts: false });
  const parsedParagraph = parsed.package.document.content.at(0);
  expect(
    parsedParagraph?.type === "paragraph" ? parsedParagraph.listRendering?.marker : undefined,
  ).toBe("1..");
  const canonical = normalizeCanonicalListRendering(parsed).document.package.document.content.at(0);
  expect(canonical?.type === "paragraph" ? canonical.listRendering?.marker : undefined).toBe("1..");
  const pm = toProseDoc(parsed);
  const snapshot = createFolioAIEditSnapshot(pm).blocks.at(0);
  expect(snapshot && snapshot.kind !== "diagnostic" ? snapshot.displayLabel : undefined).toBe(
    "1..",
  );
  const blocks = toFlowBlocks(pm);
  expect(blocks.at(0)?.kind).toBe("paragraph");
  expect(blocks.at(0)?.attrs?.listMarker).toBe("1..");
});
