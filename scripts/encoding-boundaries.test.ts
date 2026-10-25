import { describe, expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";
import type { TSType } from "oxc-parser";

import { scanTestTimeout } from "../test/scan-timeout";
import {
  bindingElements,
  collectSyntaxTree,
  createLineIndex,
  isBlock,
  isFunctionLike,
  type Node,
  nodeText,
  parameterParts,
  parametersOf,
  parseLikeTypeScript,
  propertyNameText,
} from "./lib/oxc-ast";

setDefaultTimeout(scanTestTimeout(30_000));

const PACKAGES = path.resolve(import.meta.dir, "../packages");
const OWNER = "utf8-boundaries/src/encoding.ts";
const GENERATED = [
  "docx-core/src/generated/docx_kernel.js",
  "core/src/generated/text_shaper.js",
];
const EXTENSIONS = [
  "ts",
  "tsx",
  "js",
  "jsx",
  "mts",
  "cts",
  "mjs",
  "cjs",
  "vue",
];
const SOURCE_GLOB = `*/src/**/*.{${EXTENSIONS.join(",")}}`;
const SOURCE_FILES = [...new Bun.Glob(SOURCE_GLOB).scanSync({ cwd: PACKAGES })];

type Violation = { line: number; reason: string };

const property = (node: Node): string | undefined => {
  if (node.type !== "MemberExpression") return undefined;
  if (!node.computed)
    return node.property.type === "PrivateIdentifier"
      ? `#${node.property.name}`
      : node.property.name;
  if (
    node.property.type === "Literal" &&
    typeof node.property.value === "string"
  )
    return node.property.value;
  return undefined;
};

const isStringLiteral = (
  node: Node | undefined,
): node is Node & { type: "Literal"; value: string } =>
  node?.type === "Literal" && typeof node.value === "string";

const scriptText = (source: string, file: string): string =>
  file.endsWith(".vue")
    ? [...source.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gu)]
        .map((match) => match[1] ?? "")
        .join("\n")
    : source;

/** A TypeScript parameter or variable declaration. */
type Declaration = {
  readonly kind: "parameter" | "variable";
  /** Stands where TypeScript's declaration node does, under the same parent. */
  readonly node: Node;
  readonly name: Node;
  readonly type: TSType | undefined;
  readonly initializer: Node | undefined;
};

const findViolations = (text: string, file = "fixture.ts"): Violation[] => {
  const script = scriptText(text, file);
  const { program } = parseLikeTypeScript(
    file,
    script,
    file.endsWith(".tsx") || file.endsWith(".jsx") ? "tsx" : "ts",
  );
  const { nodes, parentOf } = collectSyntaxTree(program);
  const lineOf = createLineIndex(script);
  const violations: Violation[] = [];
  const codecs = new Set(["TextEncoder", "TextDecoder"]);
  const decoders = new Set(["decodeText"]);
  const encoders = new Set(["encodeUtf8"]);
  const buffers = new Set(["Buffer"]);
  const bufferFactories = new Set<string>();
  const binaryCodecs = new Set(["atob", "btoa"]);
  const textNames = new Set<string>();
  const declarations: Declaration[] = [];
  for (const node of nodes) {
    if (node.type === "ImportSpecifier") {
      const imported =
        node.imported.type === "Literal"
          ? node.imported.value
          : node.imported.name;
      if (decoders.has(imported)) decoders.add(node.local.name);
      if (encoders.has(imported)) encoders.add(node.local.name);
      if (buffers.has(imported)) buffers.add(node.local.name);
      if (binaryCodecs.has(imported)) binaryCodecs.add(node.local.name);
    }
    if (node.type === "VariableDeclarator")
      declarations.push({
        kind: "variable",
        node,
        ...parameterParts(node.id),
        initializer: node.init ?? undefined,
      });
    // TypeScript declares a catch clause's variable like any other.
    if (node.type === "CatchClause" && node.param)
      declarations.push({
        kind: "variable",
        node: node.param,
        ...parameterParts(node.param),
        initializer: undefined,
      });
    if (isFunctionLike(node))
      for (const parameter of parametersOf(node))
        declarations.push({
          kind: "parameter",
          node: parameter,
          ...parameterParts(parameter),
        });
  }
  // Follow aliases to a fixed point so an extra assignment cannot open a bypass.
  let changed = true;
  const alias = (set: Set<string>, name: string): void => {
    if (set.has(name)) return;
    set.add(name);
    changed = true;
  };
  while (changed) {
    changed = false;
    for (const declaration of declarations) {
      const { name, initializer } = declaration;
      if (
        name.type === "Identifier" &&
        declaration.type?.type === "TSStringKeyword"
      )
        alias(textNames, name.name);
      if (declaration.kind !== "variable" || !initializer) continue;
      if (name.type === "Identifier") {
        if (
          isStringLiteral(initializer) ||
          initializer.type === "TemplateLiteral"
        )
          alias(textNames, name.name);
        const target =
          initializer.type === "Identifier"
            ? initializer.name
            : property(initializer);
        if (target)
          for (const set of [
            decoders,
            encoders,
            buffers,
            bufferFactories,
            binaryCodecs,
            textNames,
          ])
            if (set.has(target)) alias(set, name.name);
        if (
          property(initializer) === "from" &&
          initializer.type === "MemberExpression" &&
          !initializer.computed &&
          initializer.object.type === "Identifier" &&
          buffers.has(initializer.object.name)
        )
          alias(bufferFactories, name.name);
      }
      if (name.type !== "ObjectPattern") continue;
      for (const element of bindingElements(name)) {
        if (element.name.type !== "Identifier") continue;
        const target =
          propertyNameText(script, element)?.replace(/["']/gu, "") ??
          element.name.name;
        for (const set of [decoders, encoders, buffers, binaryCodecs])
          if (set.has(target)) alias(set, element.name.name);
        if (
          target === "from" &&
          initializer.type === "Identifier" &&
          buffers.has(initializer.name)
        )
          alias(bufferFactories, element.name.name);
      }
    }
  }
  const add = (node: Node, reason: string): void => {
    violations.push({
      line: lineOf(node.start).line + 1,
      reason,
    });
  };
  const loopAncestor = (node: Node): boolean => {
    for (let parent = parentOf(node); parent; parent = parentOf(parent)) {
      if (
        parent.type === "ForStatement" ||
        parent.type === "ForOfStatement" ||
        parent.type === "ForInStatement" ||
        parent.type === "WhileStatement" ||
        parent.type === "DoWhileStatement"
      )
        return true;
    }
    return false;
  };
  const localScope = (node: Node | undefined): Node => {
    for (let current = node; current; current = parentOf(current)) {
      if (
        isBlock(current) ||
        isFunctionLike(current) ||
        current.type === "ForStatement" ||
        current.type === "ForOfStatement" ||
        current.type === "ForInStatement" ||
        current.type === "Program"
      )
        return current;
    }
    return program;
  };
  const localBindings = new Map<Node, Map<string, Declaration>>();
  for (const candidate of declarations) {
    if (candidate.name.type !== "Identifier") continue;
    const scope = localScope(parentOf(candidate.node));
    let scoped = localBindings.get(scope);
    if (!scoped) {
      scoped = new Map();
      localBindings.set(scope, scoped);
    }
    scoped.set(candidate.name.name, candidate);
  }
  const localDeclaration = (identifier: Node & { name: string }) => {
    for (
      let scope: Node | undefined = localScope(parentOf(identifier));
      scope;
      scope = scope === program ? undefined : localScope(parentOf(scope))
    ) {
      const declaration = localBindings.get(scope)?.get(identifier.name);
      if (declaration) return declaration;
    }
    return undefined;
  };
  const chunkOrigin = (
    expression: Node,
    seen = new Set<Node | Declaration>(),
  ): boolean => {
    if (seen.has(expression)) return false;
    seen.add(expression);
    if (
      expression.type === "ParenthesizedExpression" ||
      expression.type === "TSAsExpression" ||
      expression.type === "TSNonNullExpression"
    )
      return chunkOrigin(expression.expression, seen);
    if (expression.type === "Identifier") {
      const declaration = localDeclaration(expression);
      if (!declaration || seen.has(declaration)) return false;
      seen.add(declaration);
      if (declaration.kind === "variable") {
        const list = parentOf(declaration.node);
        if (
          list?.type === "VariableDeclaration" &&
          parentOf(list)?.type === "ForOfStatement"
        )
          return true;
        return (
          declaration.initializer !== undefined &&
          chunkOrigin(declaration.initializer, seen)
        );
      }
      const owner = parentOf(declaration.node);
      const callback = owner === undefined ? undefined : parentOf(owner);
      if (
        callback?.type !== "CallExpression" ||
        property(callback.callee) !== "on"
      )
        return false;
      const event = callback.arguments.at(0);
      return isStringLiteral(event) && event.value === "data";
    }
    if (expression.type === "CallExpression") {
      const callee = expression.callee;
      const method = property(callee);
      // A Map key selects a complete stored value; its bytes do not flow into that value.
      if (
        method === "get" &&
        callee.type === "MemberExpression" &&
        !callee.computed &&
        callee.object.type === "Identifier"
      ) {
        const declaration = localDeclaration(callee.object);
        const initializer = declaration?.initializer;
        if (
          declaration &&
          ((declaration.type &&
            /^(?:Readonly)?Map\s*</u.test(
              nodeText(script, declaration.type),
            )) ||
            (declaration.kind === "variable" &&
              initializer &&
              initializer.type === "NewExpression" &&
              initializer.callee.type === "Identifier" &&
              initializer.callee.name === "Map" &&
              localDeclaration(initializer.callee) === undefined))
        )
          return false;
      }
      if (
        (method === "slice" || method === "subarray") &&
        loopAncestor(expression)
      )
        return true;
      if (callee.type === "MemberExpression") {
        if (chunkOrigin(callee.object, seen)) return true;
      }
      return expression.arguments.some((argument) =>
        chunkOrigin(argument, seen),
      );
    }
    if (expression.type === "ConditionalExpression")
      return (
        chunkOrigin(expression.consequent, seen) ||
        chunkOrigin(expression.alternate, seen)
      );
    return false;
  };
  const chunkConsumer = (node: Node, first: Node | undefined): boolean => {
    if (first && chunkOrigin(first)) return true;
    for (let parent = parentOf(node); parent; parent = parentOf(parent)) {
      if (
        parent.type === "ForOfStatement" &&
        first &&
        first.type === "Identifier" &&
        parent.left.type === "VariableDeclaration" &&
        parent.left.declarations.some(
          ({ id }) => id.type === "Identifier" && id.name === first.name,
        )
      )
        return true;
      if (
        parent.type === "CallExpression" &&
        property(parent.callee) === "on"
      ) {
        const event = parent.arguments.at(0);
        if (isStringLiteral(event) && event.value === "data") return true;
      }
    }
    return false;
  };
  const isTextual = (first: Node | undefined): boolean =>
    first !== undefined &&
    (isStringLiteral(first) ||
      first.type === "TemplateLiteral" ||
      (first.type === "Identifier" && textNames.has(first.name)));
  for (const node of nodes) {
    if (node.type === "Identifier" && codecs.has(node.name))
      add(node, "raw text codec reference");
    if (node.type === "MemberExpression" && codecs.has(property(node) ?? ""))
      add(node, "raw text codec property");
    if (node.type !== "CallExpression") continue;
    const method = property(node.callee);
    const first = node.arguments.at(0);
    const name = node.callee.type === "Identifier" ? node.callee.name : method;
    const argument = first === undefined ? "" : nodeText(script, first);
    const options = node.arguments.at(1);
    const binaryField =
      options &&
      options.type === "ObjectExpression" &&
      options.properties.some(
        (option) =>
          option.type === "Property" &&
          option.kind === "init" &&
          !option.method &&
          !option.shorthand &&
          !option.computed &&
          nodeText(script, option.key) === "encoding" &&
          isStringLiteral(option.value) &&
          option.value.value === "binary",
      );
    if (name && binaryCodecs.has(name))
      add(node, "binary string codec must use the encoding owner");
    if (method === "fromCharCode" || method === "fromCodePoint") {
      if (/Uint8Array|bytes|buffer|chunk|view/iu.test(argument))
        add(node, "manual byte-to-text codec");
    }
    if (method === "charCodeAt") {
      for (let parent = parentOf(node); parent; parent = parentOf(parent)) {
        if (
          parent.type === "CallExpression" &&
          ((property(parent.callee) === "from" &&
            /Uint8Array/u.test(nodeText(script, parent.callee))) ||
            /^set(?:Int|Uint)(?:8|16|32)$/u.test(property(parent.callee) ?? ""))
        ) {
          add(node, "manual binary string-to-byte codec");
          break;
        }
        if (
          parent.type === "AssignmentExpression" &&
          parent.operator === "=" &&
          parent.left.type === "MemberExpression" &&
          parent.left.computed &&
          /bytes|buffer|view|encoded/iu.test(
            nodeText(script, parent.left.object),
          )
        ) {
          add(node, "manual binary string-to-byte codec");
          break;
        }
      }
    }
    if (
      name &&
      decoders.has(name) &&
      !binaryField &&
      (chunkConsumer(node, first) ||
        /(?:^|\W)(?:chunk|chunks|piece|segment)\b/iu.test(argument) ||
        (loopAncestor(node) && /\.(?:slice|subarray)\(/u.test(argument)))
    )
      add(node, "chunk decoding must use createTextReader.read");
    if (
      name &&
      encoders.has(name) &&
      (chunkConsumer(node, first) ||
        /(?:^|\W)(?:chunk|chunks|piece|segment)\b/iu.test(argument) ||
        (loopAncestor(node) && /\.slice\(/u.test(argument)))
    )
      add(node, "encode complete text before chunking");
    if (
      method === "toString" &&
      first &&
      (!isStringLiteral(first) || first.value !== "hex") &&
      !(first.type === "Literal" && typeof first.value === "number")
    )
      add(node, "textual Buffer.toString encoding");
    if (
      method === "from" &&
      node.callee.type === "MemberExpression" &&
      node.callee.object.type === "Identifier" &&
      buffers.has(node.callee.object.name)
    ) {
      if (node.arguments.length > 1 || isTextual(first))
        add(node, "textual Buffer.from encoding");
    }
    if (
      name &&
      bufferFactories.has(name) &&
      (node.arguments.length > 1 || isTextual(first))
    )
      add(node, "aliased textual Buffer.from encoding");
  }
  return violations;
};

/** wasm-bindgen decodes one complete allocation, never independent transport chunks. */
const generatedViolations = (text: string, file: string): string[] => {
  const { program } = parseLikeTypeScript(file, text, "jsx");
  const problems: string[] = [];
  let decoderConstructions = 0;
  let encoderConstructions = 0;
  let decodes = 0;
  let encodes = 0;
  let encodeInto = 0;
  const compact = (node: Node): string =>
    nodeText(text, node).replace(/\s+/gu, "");
  for (const node of collectSyntaxTree(program).nodes) {
    if (node.type === "NewExpression") {
      if (nodeText(text, node.callee) === "TextDecoder") decoderConstructions++;
      if (nodeText(text, node.callee) === "TextEncoder") encoderConstructions++;
    }
    if (node.type === "CallExpression") {
      const method = property(node.callee);
      if (method === "decode") {
        decodes++;
        const args = node.arguments.map(compact);
        if (
          args.length !== 0 &&
          (args.length !== 1 ||
            args[0] !== "getUint8ArrayMemory0().subarray(ptr,ptr+len)")
        )
          problems.push(
            "decoder no longer reads one complete memory allocation",
          );
      }
      if (method === "encode") {
        encodes++;
        if (
          node.arguments.length !== 1 ||
          compact(node.arguments[0]!) !== "arg"
        )
          problems.push("encoder does not receive the complete argument");
      }
      if (method === "encodeInto") {
        encodeInto++;
        if (
          node.arguments.length !== 2 ||
          compact(node.arguments[0]!) !== "arg" ||
          compact(node.arguments[1]!) !== "view"
        )
          problems.push("encodeInto argument changed");
      }
    }
  }
  const detected = findViolations(text, file);
  const rawReferences = detected.filter(({ reason }) =>
    reason.startsWith("raw text codec"),
  );
  if (rawReferences.length !== (file === GENERATED[0] ? 2 : 3))
    problems.push("generated codec references changed");
  if (detected.some(({ reason }) => !reason.startsWith("raw text codec")))
    problems.push(
      "additional encoding escaped the generated whole-buffer codecs",
    );
  if (decoderConstructions !== 2 || decodes !== 3)
    problems.push("generated decoder shape changed");
  if (file === GENERATED[0]) {
    if (encoderConstructions !== 0 || encodes !== 0 || encodeInto !== 0)
      problems.push("unexpected kernel encoder");
    return problems;
  }
  if (encoderConstructions !== 1 || encodes !== 2 || encodeInto !== 1)
    problems.push("generated shaper encoder shape changed");
  // TypeScript's statement is the whole `export function`, keyword included.
  const passString = program.body.find((statement) => {
    const declaration =
      (statement.type === "ExportNamedDeclaration" ||
        statement.type === "ExportDefaultDeclaration") &&
      statement.declaration
        ? statement.declaration
        : statement;
    return (
      declaration.type === "FunctionDeclaration" &&
      declaration.id?.name === "passStringToWasm0"
    );
  });
  const body = passString ? compact(passString) : "";
  // The ASCII-only prefix stops before any surrogate. The remainder is encoded
  // once into a buffer with three bytes per UTF-16 unit, sufficient even for pairs.
  if (
    !body.includes("for(;offset<len;offset++)") ||
    !body.includes("constcode=arg.charCodeAt(offset);if(code>0x7F)break;") ||
    !body.includes("arg=arg.slice(offset);") ||
    !body.includes("len=offset+arg.length*3") ||
    !body.includes("getUint8ArrayMemory0().subarray(ptr+offset,ptr+len)") ||
    !body.includes("cachedTextEncoder.encodeInto(arg,view)")
  )
    problems.push(
      "shaper no longer encodes the whole non-ASCII remainder with sufficient capacity",
    );
  return problems;
};

const excluded = (file: string): boolean =>
  /(?:^|\/)(?:__tests__|__fixtures__|tests?|fixtures?)(?:\/|$)/u.test(file) ||
  /\.(?:test|spec|fixture)\.[cm]?[jt]sx?$/u.test(file);

const sourceViolations = (file: string, source: string): string[] => {
  if (file === OWNER || excluded(file)) return [];
  if (GENERATED.includes(file)) return generatedViolations(source, file);
  return findViolations(source, file).map(
    ({ line, reason }) => `${line}: ${reason}`,
  );
};

describe("production text encoding boundaries", () => {
  test("every production package and source extension rejects local codecs and chunk decoding", async () => {
    const packages = new Set<string>();
    for (const file of SOURCE_FILES) packages.add(file.split("/")[0]!);
    const bad = [
      "new TextEncoder().encode(text)",
      "const Codec = TextDecoder; new Codec().decode(chunk)",
      "const { TextDecoder: Codec } = require('node:util'); new Codec()",
      "function consume(codec: TextDecoder) { return codec.decode(bytes); }",
      "import { TextEncoder as Codec } from 'node:util'; new Codec()",
      "new globalThis['TextEncoder']()",
      "import { decodeText as decode } from '@stll/utf8-boundaries'; for (const chunk of chunks) decode(chunk)",
      "for (let i = 0; i < bytes.length; i += 3) decodeText(bytes.subarray(i, i + 3))",
      "for (const data of input) decodeText(data)",
      "stream.on('data', data => decodeText(data))",
      "for (const chunk of chunks) { const bytes = chunk; const copied = bytes; decodeText(copied); }",
      "stream.on('data', chunk => { const bytes = chunk.subarray(0, 3); const alias = bytes; decodeText(alias); })",
      "for (const chunk of chunks) { const text = chunk; const copied = text; encodeUtf8(copied); }",
      "for (let i = 0; i < text.length; i += 3) { const part = text.slice(i, i + 3); const alias = part; encodeUtf8(alias); }",
      "for (let i = 0; i < text.length; i += 3) encodeUtf8(text.slice(i, i + 3))",
      "for (const chunk of chunks) encodeUtf8(String(chunk))",
      "Buffer.from('日本語')",
      "const text = '日本語'; const copied = text; Buffer.from(copied)",
      "const { Buffer: Bytes } = require('node:buffer'); Bytes.from(text, 'utf8')",
      "bytes.toString('utf8')",
      "const decode = atob; decode(text)",
      "globalThis.btoa(binary)",
      "Uint8Array.from(binary, character => character.charCodeAt(0))",
      "for (let i = 0; i < text.length; i++) bytes[i] = text.charCodeAt(i)",
      "for (let i = 0; i < text.length; i++) view.setUint16(i * 2, text.charCodeAt(i), true)",
      "String.fromCharCode(...bytes)",
    ];
    for (const packageName of packages)
      for (const extension of EXTENSIONS)
        for (const fixture of bad) {
          const file = `${packageName}/src/boundary-fixture.${extension}`;
          const source =
            extension === "vue" ? `<script setup>${fixture}</script>` : fixture;
          expect({
            file,
            fixture,
            rejected: sourceViolations(file, source).length > 0,
          }).toEqual({
            file,
            fixture,
            rejected: true,
          });
        }
  });

  test("whole buffers, stateful readers, and byte diagnostics remain valid", () => {
    expect(
      findViolations(
        "decodeText(bytes); encodeUtf8(text); for (const chunk of chunks) reader.read(chunk); reader.finish(); Buffer.from(bytes.subarray(1)).toString('hex'); Buffer.from(await serializeDocumentToDocx(doc));",
      ),
    ).toEqual([]);
    expect(
      findViolations(
        "const parts = new Map<string, string>(); for (const path of paths) { const xml = parts.get(path); if (xml) encodeUtf8(xml); } for (let i = 0; i < bytes.length; i += 4) decodeText(bytes.subarray(i, i + 4), { encoding: 'binary' });",
      ),
    ).toEqual([]);
    expect(
      findViolations(
        "for (let i = 0; i < bytes.length; i += 4) decodeText(bytes.subarray(i, i + 4), { encoding: 'utf-8' });",
      ).length,
    ).toBeGreaterThan(0);
    expect(
      findViolations(
        "const Map = class { get(chunk) { return chunk; } }; const parts = new Map(); for (const chunk of chunks) encodeUtf8(parts.get(chunk));",
      ).length,
    ).toBeGreaterThan(0);
  });

  test("generated codec exceptions retain their verified whole-buffer shapes", async () => {
    for (const file of GENERATED) {
      const source = await Bun.file(path.join(PACKAGES, file)).text();
      expect(sourceViolations(file, source)).toEqual([]);
      expect(
        sourceViolations(file, source + "\nnew TextDecoder().decode(chunk);")
          .length,
      ).toBeGreaterThan(0);
      expect(
        sourceViolations(
          file,
          source.replace(
            "subarray(ptr, ptr + len)",
            "subarray(ptr, ptr + len - 1)",
          ),
        ).length,
      ).toBeGreaterThan(0);
      if (file === GENERATED[1])
        expect(
          sourceViolations(file, source.replace("code > 0x7F", "code > 0xFFFF"))
            .length,
        ).toBeGreaterThan(0);
    }
  });

  test(
    "production encodings flow through the owner",
    async () => {
      const offenders: string[] = [];
      for (const file of SOURCE_FILES) {
        const source = await Bun.file(path.join(PACKAGES, file)).text();
        for (const reason of sourceViolations(file, source))
          offenders.push(`${file}:${reason}`);
      }
      expect(offenders).toEqual([]);
    },
    scanTestTimeout(30_000),
  );
});
