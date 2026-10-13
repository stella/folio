import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";
import { EditorState, Plugin, PluginKey } from "prosemirror-state";
import ts from "typescript";
import { schema } from "./schema";
import { createPluginKeyOwner } from "./pluginKeyOwner";
import { createAICitationDecorationsPlugin } from "./plugins/aiCitationDecorations";
import { createAISuggestionDecorationsPlugin } from "./plugins/aiSuggestionDecorations";
import {
  anonymizationDecorationsKey,
  createAnonymizationDecorationsPlugin,
  setAnonymizationTermsMeta,
} from "./plugins/anonymizationDecorations";
import { autocompleteSuggestionPlugin } from "./plugins/autocompleteSuggestion";
import { createContentControlWidgetsPlugin } from "./plugins/contentControlWidgets";
import { createDocumentNumberingPlugin } from "./plugins/documentNumbering";
import {
  createDocumentStyleContextPlugin,
  createDocumentStylesPlugin,
} from "./plugins/documentStyles";
import { createSelectionTrackerPlugin } from "./plugins/selectionTracker";
import { createSuggestionModePlugin } from "./plugins/suggestionMode";
import { createTemplateDirectivesPlugin } from "./plugins/templateDirectives";
import { createTemplatePreviewValuesPlugin } from "./plugins/templatePreviewValues";
import { templateSlashMenuPlugin } from "./plugins/templateSlashMenu";
import { createTransactionInvariantPlugin } from "./plugins/transactionInvariants";
import { AutoBidiDetectionExtension } from "./extensions/features/AutoBidiDetectionExtension";
import { EmptyParagraphFormatExtension } from "./extensions/features/EmptyParagraphFormatExtension";
import { ImageDragExtension } from "./extensions/features/ImageDragExtension";
import { JoinedRunStyleExtension } from "./extensions/features/JoinedRunStyleExtension";
import { ParaIdAllocatorExtension } from "./extensions/features/ParaIdAllocatorExtension";
import { createParagraphChangeTrackerPlugin } from "./extensions/features/ParagraphChangeTrackerExtension";
import { RunIdentityExtension } from "./extensions/marks/RunIdentityExtension";
import type { AnyExtension } from "./extensions/types";

const extensionPlugin = (extension: AnyExtension): Plugin => {
  const plugin = extension.onSchemaReady({ schema }).plugins?.at(0);
  expect(plugin).toBeDefined();
  if (!plugin) throw new TypeError("Extension did not create its owned plugin");
  return plugin;
};
// Previous plugin fixtures exercised one editor at a time; this oracle checks
// factory key identity across concurrent editors and reinitialization.
const FACTORIES = {
  aiCitationDecorations: [createAICitationDecorationsPlugin],
  aiSuggestionDecorations: [createAISuggestionDecorationsPlugin],
  anonymizationDecorations: [() => createAnonymizationDecorationsPlugin()],
  autocompleteSuggestion: [() => autocompleteSuggestionPlugin()],
  contentControlWidgets: [() => createContentControlWidgetsPlugin()],
  documentNumbering: [() => createDocumentNumberingPlugin(null)],
  documentStyles: [
    () => createDocumentStyleContextPlugin(null),
    () => createDocumentStylesPlugin(undefined),
  ],
  selectionTracker: [() => createSelectionTrackerPlugin()],
  suggestionMode: [() => createSuggestionModePlugin()],
  templateDirectives: [() => createTemplateDirectivesPlugin()],
  templatePreviewValues: [createTemplatePreviewValuesPlugin],
  templateSlashMenu: [() => templateSlashMenuPlugin()],
  transactionInvariants: [createTransactionInvariantPlugin],
  autoBidiDetection: [() => extensionPlugin(AutoBidiDetectionExtension())],
  emptyParagraphFormat: [() => extensionPlugin(EmptyParagraphFormatExtension())],
  imageDrag: [() => extensionPlugin(ImageDragExtension())],
  pastedTableEnd: [() => extensionPlugin(JoinedRunStyleExtension())],
  paraIdAllocator: [() => extensionPlugin(ParaIdAllocatorExtension())],
  paragraphChangeTracker: [createParagraphChangeTrackerPlugin],
  runIdentityStrip: [() => extensionPlugin(RunIdentityExtension())],
};

const declaredOwners = (): string[] => {
  const owners: string[] = [];
  const sourceRoot = path.resolve(import.meta.dir, "..");
  for (const filename of new Bun.Glob("**/*.ts").scanSync({ cwd: sourceRoot })) {
    if (/\.test\.|(?:^|\/)(?:__tests__|typecheck)\//u.test(filename)) continue;
    const source = ts.createSourceFile(
      filename,
      readFileSync(path.join(sourceRoot, filename), "utf8"),
      ts.ScriptTarget.Latest,
      true,
    );
    const visit = (node: ts.Node): void => {
      if (
        ts.isCallExpression(node) &&
        ts.isIdentifier(node.expression) &&
        node.expression.text === "createPluginKeyOwner"
      ) {
        const name = node.arguments.at(0);
        expect(name && ts.isStringLiteral(name)).toBe(true);
        if (name && ts.isStringLiteral(name)) owners.push(name.text);
      }
      if (
        ts.isNewExpression(node) &&
        ts.isIdentifier(node.expression) &&
        node.expression.text === "PluginKey" &&
        filename !== "prosemirror/pluginKeyOwner.ts"
      ) {
        let ancestor: ts.Node | undefined = node.parent;
        while (ancestor && !ts.isFunctionLike(ancestor)) ancestor = ancestor.parent;
        expect(ancestor).toBeDefined();
      }
      ts.forEachChild(node, visit);
    };
    visit(source);
  }
  return owners.sort();
};

test("every declared owner has exercised factories and raw keys are factory scoped", () => {
  expect(declaredOwners()).toEqual(Object.keys(FACTORIES).sort());
});

for (const [name, factories] of Object.entries(FACTORIES)) {
  test(`${name} factories allocate distinct keys for editors and reinitialization`, () => {
    for (const create of factories) {
      const plugins = [create(), create(), create()];
      const keys = plugins.map(({ spec }) => spec.key);
      expect(keys.every((key) => key instanceof PluginKey)).toBe(true);
      expect(new Set(keys).size).toBe(plugins.length);
      const states = plugins.map((plugin) => EditorState.create({ schema, plugins: [plugin] }));
      for (let index = 0; index < states.length; index++) {
        const state = states.at(index);
        const key = keys.at(index);
        expect(state && key?.get(state)).toBe(plugins.at(index));
        const other = states.at((index + 1) % states.length);
        expect(other && key?.get(other)).toBeUndefined();
        if (state) {
          const next = state.apply(state.tr.insertText(`editor ${index}`, 1));
          expect(key?.get(next)).toBe(plugins.at(index));
          expect(state.doc.textContent).toBe("");
        }
      }
    }
  });
}

test("immutable descriptors read only their owned state across interleaved metadata updates", () => {
  const owner = createPluginKeyOwner<number>("test.counter");
  const create = () =>
    new Plugin({
      key: owner.createKey(),
      state: {
        init: () => 0,
        apply: (transaction, value: number) => value + (transaction.getMeta(owner) ?? 0),
      },
    });
  let first = EditorState.create({ schema, plugins: [create()] });
  let second = EditorState.create({ schema, plugins: [create()] });
  expect(Object.isFrozen(owner)).toBe(true);
  for (const increment of [1, 2, 3]) {
    first = first.apply(first.tr.setMeta(owner, increment));
    expect(owner.getState(second)).toBe((increment - 1) * 10);
    second = second.apply(second.tr.setMeta(owner, 10));
    expect(owner.getState(first)).toBe((increment * (increment + 1)) / 2);
  }
  expect(owner.getState(first)).toBe(6);
  expect(owner.getState(second)).toBe(30);
  const duplicated = EditorState.create({ schema, plugins: [create(), create()] });
  expect(() => owner.getState(duplicated)).toThrow("Multiple plugins belong");
});

test("real anonymization metadata and state remain editor owned", () => {
  const doc = schema.node("doc", null, [schema.node("paragraph", null, schema.text("Alice Bob"))]);
  let first = EditorState.create({ doc, plugins: [createAnonymizationDecorationsPlugin()] });
  let second = EditorState.create({ doc, plugins: [createAnonymizationDecorationsPlugin()] });
  const alice = setAnonymizationTermsMeta([{ canonical: "Alice", label: "Person A" }]);
  const bob = setAnonymizationTermsMeta([{ canonical: "Bob", label: "Person B" }]);
  first = first.apply(first.tr.setMeta(alice.key, alice.payload));
  expect(anonymizationDecorationsKey.getState(second)?.matches).toEqual([]);
  second = second.apply(second.tr.setMeta(bob.key, bob.payload));
  expect(
    anonymizationDecorationsKey.getState(first)?.matches.map(({ canonical }) => canonical),
  ).toEqual(["Alice"]);
  expect(
    anonymizationDecorationsKey.getState(second)?.matches.map(({ canonical }) => canonical),
  ).toEqual(["Bob"]);
});
