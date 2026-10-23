import { expect, test } from "bun:test";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { EditorState, Plugin, PluginKey } from "prosemirror-state";
import ts from "typescript";

import { AutoBidiDetectionExtension } from "./extensions/features/AutoBidiDetectionExtension";
import { EmptyParagraphFormatExtension } from "./extensions/features/EmptyParagraphFormatExtension";
import { ImageDragExtension } from "./extensions/features/ImageDragExtension";
import { JoinedRunStyleExtension } from "./extensions/features/JoinedRunStyleExtension";
import { createParagraphChangeTrackerPlugin } from "./extensions/features/ParagraphChangeTrackerExtension";
import { ParaIdAllocatorExtension } from "./extensions/features/ParaIdAllocatorExtension";
import { RunIdentityExtension } from "./extensions/marks/RunIdentityExtension";
import type { AnyExtension } from "./extensions/types";
import { createPluginKeyOwner } from "./pluginKeyOwner";
import { createAICitationDecorationsPlugin } from "./plugins/aiCitationDecorations";
import { createAISuggestionDecorationsPlugin } from "./plugins/aiSuggestionDecorations";
import {
  anonymizationDecorationsKey,
  createAnonymizationDecorationsPlugin,
  getAnonymizationMatches,
  setAnonymizationTermsMeta,
} from "./plugins/anonymizationDecorations";
import { autocompleteSuggestionPlugin } from "./plugins/autocompleteSuggestion";
import * as autocomplete from "./plugins/autocompleteSuggestion";
import { createContentControlWidgetsPlugin } from "./plugins/contentControlWidgets";
import { createDocumentNumberingPlugin } from "./plugins/documentNumbering";
import {
  createDocumentStyleContextPlugin,
  createDocumentStylesPlugin,
} from "./plugins/documentStyles";
import { createSelectionTrackerPlugin } from "./plugins/selectionTracker";
import { createSuggestionModePlugin } from "./plugins/suggestionMode";
import { createTemplateDirectivesPlugin } from "./plugins/templateDirectives";
import * as directives from "./plugins/templateDirectives";
import { createTemplatePreviewValuesPlugin } from "./plugins/templatePreviewValues";
import * as previewValues from "./plugins/templatePreviewValues";
import { templateSlashMenuPlugin } from "./plugins/templateSlashMenu";
import * as slashMenu from "./plugins/templateSlashMenu";
import { createTransactionInvariantPlugin } from "./plugins/transactionInvariants";
import { schema } from "./schema";

// Bun caches query-string imports by their source path. Give the real plugin
// body a distinct file, while preserving the original imports and shared peers.
const importRepeatedPlugin = async (name: string) => {
  const filename = path.join(import.meta.dir, "plugins", `${name}.ts`);
  const source = ts.createSourceFile(
    filename,
    readFileSync(filename, "utf8"),
    ts.ScriptTarget.Latest,
    true,
  );
  const copiedSource = ts.factory.updateSourceFile(
    source,
    source.statements.map((statement) => {
      if (
        !ts.isImportDeclaration(statement) ||
        !ts.isStringLiteral(statement.moduleSpecifier)
      ) {
        return statement;
      }
      const specifier = statement.moduleSpecifier.text;
      const resolved = specifier.startsWith(".")
        ? path.resolve(path.dirname(filename), specifier)
        : import.meta.resolve(specifier);
      return ts.factory.updateImportDeclaration(
        statement,
        statement.modifiers,
        statement.importClause,
        ts.factory.createStringLiteral(resolved),
        statement.attributes,
      );
    }),
  );
  const directory = mkdtempSync(
    path.join(tmpdir(), "folio-plugin-evaluation-"),
  );
  try {
    const copy = path.join(directory, `${name}.ts`);
    writeFileSync(copy, ts.createPrinter().printFile(copiedSource));
    return await import(pathToFileURL(copy).href);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
};

const extensionPlugin = (extension: AnyExtension): Plugin => {
  const plugin = extension.onSchemaReady({ schema }).plugins?.at(0);
  expect(plugin).toBeDefined();
  if (!plugin) throw new TypeError("Extension did not create its owned plugin");
  return plugin;
};
// Previous fixtures shared one module evaluation. Enumerate every factory and
// read it through a distinct descriptor as well as checking editor key isolation.
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
  emptyParagraphFormat: [
    () => extensionPlugin(EmptyParagraphFormatExtension()),
  ],
  imageDrag: [() => extensionPlugin(ImageDragExtension())],
  pastedTableEnd: [() => extensionPlugin(JoinedRunStyleExtension())],
  paraIdAllocator: [() => extensionPlugin(ParaIdAllocatorExtension())],
  paragraphChangeTracker: [createParagraphChangeTrackerPlugin],
  runIdentityStrip: [() => extensionPlugin(RunIdentityExtension())],
};

const declaredOwners = (): string[] => {
  const owners: string[] = [];
  const sourceRoot = path.resolve(import.meta.dir, "..");
  for (const filename of new Bun.Glob("**/*.ts").scanSync({
    cwd: sourceRoot,
  })) {
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
        while (ancestor && !ts.isFunctionLike(ancestor))
          ancestor = ancestor.parent;
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
      const independentlyEvaluatedOwner = createPluginKeyOwner(name);
      const keys = plugins.map(({ spec }) => spec.key);
      expect(keys.every((key) => key instanceof PluginKey)).toBe(true);
      expect(new Set(keys).size).toBe(plugins.length);
      const states = plugins.map((plugin) =>
        EditorState.create({ schema, plugins: [plugin] }),
      );
      for (let index = 0; index < states.length; index++) {
        const state = states.at(index);
        const key = keys.at(index);
        expect(state && key?.get(state)).toBe(plugins.at(index));
        const other = states.at((index + 1) % states.length);
        expect(other && key?.get(other)).toBeUndefined();
        if (state) {
          const plugin = plugins.at(index);
          expect(plugin && independentlyEvaluatedOwner.owns(plugin)).toBe(true);
          expect(independentlyEvaluatedOwner.get(state)).toBe(
            plugins.at(index),
          );
          expect(independentlyEvaluatedOwner.getState(state)).toBe(
            plugins.at(index)?.getState(state),
          );
          const next = state.apply(state.tr.insertText(`editor ${index}`, 1));
          expect(independentlyEvaluatedOwner.getState(next)).toBe(
            plugins.at(index)?.getState(next),
          );
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
        apply: (transaction, value: number) =>
          value + (transaction.getMeta(owner) ?? 0),
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
  const duplicated = EditorState.create({
    schema,
    plugins: [create(), create()],
  });
  expect(() => owner.getState(duplicated)).toThrow("Multiple plugins belong");
});

test("real anonymization metadata and state remain editor owned", () => {
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, schema.text("Alice Bob")),
  ]);
  let first = EditorState.create({
    doc,
    plugins: [createAnonymizationDecorationsPlugin()],
  });
  let second = EditorState.create({
    doc,
    plugins: [createAnonymizationDecorationsPlugin()],
  });
  const alice = setAnonymizationTermsMeta([
    { canonical: "Alice", label: "Person A" },
  ]);
  const bob = setAnonymizationTermsMeta([
    { canonical: "Bob", label: "Person B" },
  ]);
  first = first.apply(first.tr.setMeta(alice.key, alice.payload));
  expect(anonymizationDecorationsKey.getState(second)?.matches).toEqual([]);
  second = second.apply(second.tr.setMeta(bob.key, bob.payload));
  expect(
    anonymizationDecorationsKey
      .getState(first)
      ?.matches.map(({ canonical }) => canonical),
  ).toEqual(["Alice"]);
  expect(
    anonymizationDecorationsKey
      .getState(second)
      ?.matches.map(({ canonical }) => canonical),
  ).toEqual(["Bob"]);
});

test("unrelated and unbranded keys cannot impersonate a plugin family", () => {
  const owner = createPluginKeyOwner("test.counter");
  const unrelated = new Plugin({
    key: createPluginKeyOwner("test.other").createKey(),
  });
  const unbranded = new Plugin({
    key: new PluginKey("stll.folio.plugin.test.counter"),
  });
  for (const plugin of [unrelated, unbranded]) {
    expect(owner.owns(plugin)).toBe(false);
    const state = EditorState.create({ schema, plugins: [plugin] });
    expect(owner.get(state)).toBeUndefined();
    expect(owner.getState(state)).toBeUndefined();
  }
});

test("independent evaluations of a family reject duplicate plugins in one state", () => {
  const firstOwner = createPluginKeyOwner("test.repeated");
  const secondOwner = createPluginKeyOwner("test.repeated");
  expect(firstOwner).not.toBe(secondOwner);
  const first = new Plugin({ key: firstOwner.createKey() });
  const second = new Plugin({ key: secondOwner.createKey() });
  expect(first.spec.key).not.toBe(second.spec.key);
  const state = EditorState.create({ schema, plugins: [first, second] });
  for (const owner of [firstOwner, secondOwner]) {
    expect(() => owner.get(state)).toThrow("Multiple plugins belong");
  }
});

test("anonymization readers and metadata survive actual repeated module evaluation", async () => {
  const repeated: typeof import("./plugins/anonymizationDecorations") =
    await importRepeatedPlugin("anonymizationDecorations");
  expect(repeated.anonymizationDecorationsKey).not.toBe(
    anonymizationDecorationsKey,
  );
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, schema.text("Alice Bob")),
  ]);
  const firstPlugin = createAnonymizationDecorationsPlugin();
  const secondPlugin = repeated.createAnonymizationDecorationsPlugin();
  expect(firstPlugin.spec.key).not.toBe(secondPlugin.spec.key);
  let first = EditorState.create({ doc, plugins: [firstPlugin] });
  let second = EditorState.create({ doc, plugins: [secondPlugin] });
  const alice = repeated.setAnonymizationTermsMeta([
    { canonical: "Alice", label: "person" },
  ]);
  first = first.apply(first.tr.setMeta(alice.key, alice.payload));
  expect(
    repeated.getAnonymizationMatches(first).map(({ canonical }) => canonical),
  ).toEqual(["Alice"]);
  expect(repeated.getAnonymizationMatches(second)).toEqual([]);
  const bob = setAnonymizationTermsMeta([
    { canonical: "Bob", label: "person" },
  ]);
  second = second.apply(second.tr.setMeta(bob.key, bob.payload));
  expect(
    getAnonymizationMatches(second).map(({ canonical }) => canonical),
  ).toEqual(["Bob"]);
  const clear = repeated.setAnonymizationTermsMeta([]);
  first = first.apply(first.tr.setMeta(clear.key, clear.payload));
  expect(repeated.getAnonymizationMatches(first)).toEqual([]);
  expect(
    getAnonymizationMatches(second).map(({ canonical }) => canonical),
  ).toEqual(["Bob"]);
  const duplicated = EditorState.create({
    doc,
    plugins: [firstPlugin, secondPlugin],
  });
  expect(() => repeated.getAnonymizationMatches(duplicated)).toThrow(
    "Multiple plugins belong",
  );
});

test("autocomplete readers and commands survive actual repeated module evaluation", async () => {
  const repeated: typeof import("./plugins/autocompleteSuggestion") =
    await importRepeatedPlugin("autocompleteSuggestion");
  expect(repeated.autocompleteSuggestionKey).not.toBe(
    autocomplete.autocompleteSuggestionKey,
  );
  const firstPlugin = autocompleteSuggestionPlugin();
  const secondPlugin = repeated.autocompleteSuggestionPlugin();
  expect(firstPlugin.spec.key).not.toBe(secondPlugin.spec.key);
  let first = EditorState.create({ schema, plugins: [firstPlugin] });
  let second = EditorState.create({ schema, plugins: [secondPlugin] });
  first = first.apply(
    repeated.startAutocompleteSuggestion(first.tr, 1, "first"),
  );
  first = first.apply(
    repeated.appendAutocompleteToken(first.tr, "first", "Alpha"),
  );
  first = first.apply(repeated.finishAutocompleteSuggestion(first.tr, "first"));
  expect(repeated.getAutocompleteSuggestion(first)).toMatchObject({
    status: "shown",
    text: "Alpha",
    requestId: "first",
  });
  expect(repeated.getAutocompleteSuggestion(second).status).toBe("idle");
  second = second.apply(
    autocomplete.startAutocompleteSuggestion(second.tr, 1, "second"),
  );
  second = second.apply(
    autocomplete.appendAutocompleteToken(second.tr, "second", "Beta"),
  );
  second = second.apply(
    autocomplete.finishAutocompleteSuggestion(second.tr, "second"),
  );
  expect(autocomplete.getAutocompleteSuggestion(second)).toMatchObject({
    status: "shown",
    text: "Beta",
    requestId: "second",
  });
  first = first.apply(repeated.clearAutocompleteSuggestion(first.tr));
  expect(repeated.getAutocompleteSuggestion(first).status).toBe("idle");
  expect(autocomplete.getAutocompleteSuggestion(second)).toMatchObject({
    status: "shown",
    text: "Beta",
  });
});

test("slash-menu readers and commands survive actual repeated module evaluation", async () => {
  const repeated: typeof import("./plugins/templateSlashMenu") =
    await importRepeatedPlugin("templateSlashMenu");
  expect(repeated.templateSlashMenuKey).not.toBe(
    slashMenu.templateSlashMenuKey,
  );
  const firstPlugin = templateSlashMenuPlugin();
  const secondPlugin = repeated.templateSlashMenuPlugin();
  expect(firstPlugin.spec.key).not.toBe(secondPlugin.spec.key);
  let first = EditorState.create({ schema, plugins: [firstPlugin] });
  let second = EditorState.create({ schema, plugins: [secondPlugin] });
  first = first.apply(
    first.tr.insertText("/", 1).setMeta(repeated.templateSlashMenuKey, {
      type: "open",
      from: 1,
      query: "",
    }),
  );
  first = first.apply(first.tr.insertText("first", 2));
  expect(repeated.getTemplateSlashMenu(first)).toMatchObject({
    active: true,
    query: "first",
  });
  expect(repeated.getTemplateSlashMenu(second).active).toBe(false);
  second = second.apply(
    second.tr.insertText("/", 1).setMeta(slashMenu.templateSlashMenuKey, {
      type: "open",
      from: 1,
      query: "",
    }),
  );
  second = second.apply(second.tr.insertText("second", 2));
  expect(slashMenu.getTemplateSlashMenu(second)).toMatchObject({
    active: true,
    query: "second",
  });
  first = first.apply(repeated.clearTemplateSlashMenu(first.tr));
  expect(repeated.getTemplateSlashMenu(first).active).toBe(false);
  expect(slashMenu.getTemplateSlashMenu(second)).toMatchObject({
    active: true,
    query: "second",
  });
});

test("directive readers survive actual repeated module evaluation", async () => {
  const repeated: typeof import("./plugins/templateDirectives") =
    await importRepeatedPlugin("templateDirectives");
  expect(repeated.templateDirectivesKey).not.toBe(
    directives.templateDirectivesKey,
  );
  const firstPlugin = createTemplateDirectivesPlugin();
  const secondPlugin = repeated.createTemplateDirectivesPlugin();
  expect(firstPlugin.spec.key).not.toBe(secondPlugin.spec.key);
  const firstDoc = schema.node("doc", null, [
    schema.node("paragraph", null, schema.text("{{tenant.name}}")),
  ]);
  const secondDoc = schema.node("doc", null, [
    schema.node("paragraph", null, schema.text("{{landlord.name}}")),
  ]);
  let first = EditorState.create({ doc: firstDoc, plugins: [firstPlugin] });
  const second = EditorState.create({
    doc: secondDoc,
    plugins: [secondPlugin],
  });
  expect(repeated.getTemplateDirectives(first).map(({ expr }) => expr)).toEqual(
    ["tenant.name"],
  );
  expect(
    directives.getTemplateDirectives(second).map(({ expr }) => expr),
  ).toEqual(["landlord.name"]);
  first = first.apply(
    first.tr.insertText("{{company}}", 1, first.doc.content.size - 1),
  );
  expect(repeated.getTemplateDirectives(first).map(({ expr }) => expr)).toEqual(
    ["company"],
  );
  expect(
    directives.getTemplateDirectives(second).map(({ expr }) => expr),
  ).toEqual(["landlord.name"]);
});

test("preview readers and metadata survive actual repeated module evaluation", async () => {
  const repeated: typeof import("./plugins/templatePreviewValues") =
    await importRepeatedPlugin("templatePreviewValues");
  expect(repeated.templatePreviewValuesKey).not.toBe(
    previewValues.templatePreviewValuesKey,
  );
  const firstPlugin = createTemplatePreviewValuesPlugin();
  const secondPlugin = repeated.createTemplatePreviewValuesPlugin();
  expect(firstPlugin.spec.key).not.toBe(secondPlugin.spec.key);
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, schema.text("{{tenant.name}}")),
  ]);
  let first = EditorState.create({ doc, plugins: [firstPlugin] });
  let second = EditorState.create({ doc, plugins: [secondPlugin] });
  first = first.apply(
    first.tr.setMeta(repeated.templatePreviewValuesKey, {
      preview: { values: { "tenant.name": "Alice" }, mode: "highlighted" },
    }),
  );
  expect(
    repeated.templatePreviewValuesKey.getState(first)?.entries,
  ).toMatchObject([{ expr: "tenant.name", value: "Alice" }]);
  expect(repeated.templatePreviewValuesKey.getState(second)?.entries).toEqual(
    [],
  );
  second = second.apply(
    second.tr.setMeta(previewValues.templatePreviewValuesKey, {
      preview: { values: { "tenant.name": "Bob" }, mode: "highlighted" },
    }),
  );
  expect(
    previewValues.templatePreviewValuesKey.getState(second)?.entries,
  ).toMatchObject([{ expr: "tenant.name", value: "Bob" }]);
  first = first.apply(
    first.tr.setMeta(repeated.templatePreviewValuesKey, { preview: null }),
  );
  expect(repeated.templatePreviewValuesKey.getState(first)?.entries).toEqual(
    [],
  );
  expect(
    previewValues.templatePreviewValuesKey.getState(second)?.entries,
  ).toMatchObject([{ expr: "tenant.name", value: "Bob" }]);
});
