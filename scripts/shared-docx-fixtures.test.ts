import { expect, setDefaultTimeout, test } from "bun:test";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { ZipArchive } from "@stll/docx-core/zip";

import { scanTestTimeout } from "../test/scan-timeout";
import {
  buildCommentSearchDocument,
  COMMENT_SEARCH_REPLY_LENGTHS,
} from "../tests/support/commentSearchDocument";
import { buildPanelLayoutDocument } from "../tests/support/panelLayoutDocument";
import type {
  PanelLayoutReview,
  PanelLayoutSections,
} from "../tests/support/panelLayoutDocument";
import { buildScrollRootDocument } from "../tests/support/scrollRootDocument";
import {
  generateDocxFixture,
  validateDocxFixture,
} from "../tests/support/validatedDocxFixture";
import {
  declaresBindings,
  type Node,
  nodeText,
  parseSourceModule,
  staticMemberName,
  walkAst,
} from "./lib/oxc-ast";

setDefaultTimeout(scanTestTimeout(30_000));

const root = resolve(import.meta.dir, "..");
const sharedRoots = ["tests/support", "tests/visual/fixtures"];
const reviewCases = {
  none: "none",
  "changes-only": "changes-only",
  "comment-and-changes": "comment-and-changes",
} as const satisfies Record<PanelLayoutReview, PanelLayoutReview>;
const sectionCases = {
  portrait: "portrait",
  "landscape-then-portrait": "landscape-then-portrait",
} as const satisfies Record<PanelLayoutSections, PanelLayoutSections>;
const fixtureCases = [
  ...COMMENT_SEARCH_REPLY_LENGTHS.map((sentenceCount) => ({
    file: "tests/support/commentSearchDocument.ts",
    build: () => buildCommentSearchDocument(sentenceCount),
  })),
  {
    file: "tests/support/scrollRootDocument.ts",
    build: buildScrollRootDocument,
  },
  ...Object.values(reviewCases).flatMap((review) =>
    Object.values(sectionCases).map((sections) => ({
      file: "tests/support/panelLayoutDocument.ts",
      build: () => buildPanelLayoutDocument(review, sections),
    })),
  ),
];

/** The written property (or name) a destructured binding reads. */
const bindingElementName = (
  text: string,
  node: Node,
  ancestors: readonly Node[],
) => {
  const parent = ancestors.at(-1);
  if (parent?.type !== "ObjectPattern" && parent?.type !== "ArrayPattern")
    return undefined;
  if (!declaresBindings(ancestors)) return undefined;
  if (node.type === "Property")
    return node.computed ? undefined : nodeText(text, node.key);
  if (node.type === "RestElement") return nodeText(text, node.argument);
  if (
    parent.type !== "ArrayPattern" ||
    !parent.elements.some((element) => element === node)
  )
    return undefined;
  return nodeText(text, node.type === "AssignmentPattern" ? node.left : node);
};

const packageCalls = (text: string) => {
  const { program } = parseSourceModule("fixture.ts", text);
  const serializers = new Set(["createDocx"]);
  for (const statement of program.body) {
    if (statement.type !== "ImportDeclaration") continue;
    for (const binding of statement.specifiers) {
      if (binding.type !== "ImportSpecifier") continue;
      const imported =
        binding.imported.type === "Identifier"
          ? binding.imported.name
          : binding.imported.value;
      if (imported === "createDocx") serializers.add(binding.local.name);
    }
  }
  const calls: string[] = [];
  // A shorthand property or export carries its name twice; count it once.
  const serializerReferences = new Set<number>();
  walkAst(program, (node, ancestors) => {
    const member = staticMemberName(node);
    if (member === "generate") calls.push("unchecked-zip");
    if (member === "createDocx") calls.push("unchecked-serializer");
    if (bindingElementName(text, node, ancestors) === "generate") {
      calls.push("unchecked-zip");
    }
    const parent = ancestors.at(-1);
    if (
      node.type === "Identifier" &&
      serializers.has(node.name) &&
      parent?.type !== "ImportSpecifier" &&
      !(
        parent?.type === "MemberExpression" &&
        !parent.computed &&
        parent.property === node
      ) &&
      !serializerReferences.has(node.start)
    ) {
      serializerReferences.add(node.start);
      calls.push("unchecked-serializer");
    }
  });
  return calls;
};

test("every shared fixture and generator uses validated package construction", () => {
  const unguarded: string[] = [];
  const supportProducers: string[] = [];
  for (const directory of sharedRoots) {
    for (const file of new Bun.Glob("**/*.ts").scanSync({
      cwd: resolve(root, directory),
      onlyFiles: true,
    })) {
      const relative = `${directory}/${file}`;
      if (relative === "tests/support/validatedDocxFixture.ts") continue;
      const source = readFileSync(resolve(root, relative), "utf8");
      if (packageCalls(source).length > 0) unguarded.push(relative);
      if (
        directory === "tests/support" &&
        /\bgenerateDocxFixture\s*\(/u.test(source)
      ) {
        supportProducers.push(relative);
      }
    }
  }
  expect(unguarded).toEqual([]);
  expect(supportProducers.sort()).toEqual(
    [...new Set(fixtureCases.map(({ file }) => file))].sort(),
  );
});

test("fixture guard detects direct package generators", () => {
  // Fixtures are whole modules: a syntax error stops the guard.
  expect(
    packageCalls('export default zip.generate({ compression: "store" })'),
  ).toEqual(["unchecked-zip"]);
  expect(packageCalls("export default createDocx(document)")).toEqual([
    "unchecked-serializer",
  ]);
  expect(
    packageCalls(
      `import { createDocx as serialize } from "core"; export default serialize(document)`,
    ),
  ).toEqual(["unchecked-serializer"]);
  expect(packageCalls(`export default core["createDocx"](document)`)).toEqual([
    "unchecked-serializer",
  ]);
  expect(
    packageCalls(
      `const write = zip["generate"]; export default write(options)`,
    ),
  ).toEqual(["unchecked-zip"]);
  expect(
    packageCalls(
      `const { generate: write } = zip; export default write(options)`,
    ),
  ).toEqual(["unchecked-zip"]);
  expect(
    packageCalls('export default generateDocxFixture(zip, "valid")'),
  ).toEqual([]);
});

test(
  "every shared fixture variant passes schema validation at setup",
  async () => {
    for (const { file, build } of fixtureCases) {
      await validateDocxFixture(await build(), file);
    }
  },
  scanTestTimeout(30_000),
);

test("invalid fixture fails at construction with its name and schema reason", async () => {
  const zip = await ZipArchive.read(await buildScrollRootDocument());
  const part = zip.entry("word/document.xml");
  if (!part) throw new Error("Missing fixture document part");
  zip.set(
    "word/document.xml",
    (await part.text()).replace(' w:header="720"', ""),
  );
  await expect(
    generateDocxFixture(zip, "deliberately-invalid"),
  ).rejects.toThrow(
    /Invalid DOCX fixture deliberately-invalid: invalid_schema_attribute/u,
  );
});
