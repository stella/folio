import { describe, expect, test } from "bun:test";
import path from "node:path";
import ts from "typescript";

const PACKAGES = path.resolve(import.meta.dir, "../packages");
const OWNER = "docx-core/src/encoding.ts";
const GENERATED = ["docx-core/src/generated/docx_kernel.js", "core/src/generated/text_shaper.js"];
const EXTENSIONS = ["ts", "tsx", "js", "jsx", "mts", "cts", "mjs", "cjs", "vue"];
const SOURCE_GLOB = `*/src/**/*.{${EXTENSIONS.join(",")}}`;

type Violation = { line: number; reason: string };

const property = (node: ts.Expression): string | undefined => {
  if (ts.isPropertyAccessExpression(node)) return node.name.text;
  if (
    ts.isElementAccessExpression(node) &&
    node.argumentExpression &&
    ts.isStringLiteral(node.argumentExpression)
  )
    return node.argumentExpression.text;
  return undefined;
};

const scriptText = (source: string, file: string): string =>
  file.endsWith(".vue")
    ? [...source.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gu)]
        .map((match) => match[1] ?? "")
        .join("\n")
    : source;

const findViolations = (text: string, file = "fixture.ts"): Violation[] => {
  const source = ts.createSourceFile(
    file,
    scriptText(text, file),
    ts.ScriptTarget.Latest,
    true,
    file.endsWith(".tsx") || file.endsWith(".jsx") ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
  );
  const violations: Violation[] = [];
  const codecs = new Set(["TextEncoder", "TextDecoder"]);
  const decoders = new Set(["decodeText"]);
  const encoders = new Set(["encodeUtf8"]);
  const buffers = new Set(["Buffer"]);
  const bufferFactories = new Set<string>();
  const binaryCodecs = new Set(["atob", "btoa"]);
  const textNames = new Set<string>();
  const nodes: ts.Node[] = [];
  const collect = (node: ts.Node): void => {
    nodes.push(node);
    ts.forEachChild(node, collect);
  };
  collect(source);
  for (const node of nodes) {
    if (ts.isImportSpecifier(node)) {
      const imported = node.propertyName?.text ?? node.name.text;
      if (decoders.has(imported)) decoders.add(node.name.text);
      if (encoders.has(imported)) encoders.add(node.name.text);
      if (buffers.has(imported)) buffers.add(node.name.text);
      if (binaryCodecs.has(imported)) binaryCodecs.add(node.name.text);
    }
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
    for (const node of nodes) {
      if (
        (ts.isParameter(node) || ts.isVariableDeclaration(node)) &&
        ts.isIdentifier(node.name) &&
        node.type?.kind === ts.SyntaxKind.StringKeyword
      )
        alias(textNames, node.name.text);
      if (!ts.isVariableDeclaration(node) || !node.initializer) continue;
      if (ts.isIdentifier(node.name)) {
        if (
          ts.isStringLiteral(node.initializer) ||
          ts.isTemplateExpression(node.initializer) ||
          ts.isNoSubstitutionTemplateLiteral(node.initializer)
        )
          alias(textNames, node.name.text);
        const target = ts.isIdentifier(node.initializer)
          ? node.initializer.text
          : property(node.initializer);
        if (target)
          for (const set of [decoders, encoders, buffers, bufferFactories, binaryCodecs, textNames])
            if (set.has(target)) alias(set, node.name.text);
        if (
          property(node.initializer) === "from" &&
          ts.isPropertyAccessExpression(node.initializer) &&
          ts.isIdentifier(node.initializer.expression) &&
          buffers.has(node.initializer.expression.text)
        )
          alias(bufferFactories, node.name.text);
      }
      if (!ts.isObjectBindingPattern(node.name)) continue;
      for (const element of node.name.elements) {
        if (!ts.isIdentifier(element.name)) continue;
        const target =
          element.propertyName?.getText(source).replace(/["']/gu, "") ?? element.name.text;
        for (const set of [decoders, encoders, buffers, binaryCodecs])
          if (set.has(target)) alias(set, element.name.text);
        if (
          target === "from" &&
          ts.isIdentifier(node.initializer) &&
          buffers.has(node.initializer.text)
        )
          alias(bufferFactories, element.name.text);
      }
    }
  }
  const add = (node: ts.Node, reason: string): void => {
    violations.push({
      line: source.getLineAndCharacterOfPosition(node.getStart(source)).line + 1,
      reason,
    });
  };
  const loopAncestor = (node: ts.Node): boolean => {
    for (let parent = node.parent; parent; parent = parent.parent) {
      if (
        ts.isForStatement(parent) ||
        ts.isForOfStatement(parent) ||
        ts.isForInStatement(parent) ||
        ts.isWhileStatement(parent) ||
        ts.isDoStatement(parent)
      )
        return true;
    }
    return false;
  };
  const localScope = (node: ts.Node | undefined): ts.Node => {
    for (let current = node; current; current = current.parent) {
      if (
        ts.isBlock(current) ||
        ts.isFunctionLike(current) ||
        ts.isForStatement(current) ||
        ts.isForOfStatement(current) ||
        ts.isForInStatement(current) ||
        ts.isSourceFile(current)
      )
        return current;
    }
    return source;
  };
  const localBindings = new Map<
    ts.Node,
    Map<string, ts.ParameterDeclaration | ts.VariableDeclaration>
  >();
  for (const candidate of nodes) {
    if (
      !(ts.isParameter(candidate) || ts.isVariableDeclaration(candidate)) ||
      !ts.isIdentifier(candidate.name)
    )
      continue;
    const scope = localScope(candidate.parent);
    let declarations = localBindings.get(scope);
    if (!declarations) {
      declarations = new Map();
      localBindings.set(scope, declarations);
    }
    declarations.set(candidate.name.text, candidate);
  }
  const localDeclaration = (identifier: ts.Identifier) => {
    for (
      let scope: ts.Node | undefined = localScope(identifier.parent);
      scope;
      scope = scope === source ? undefined : localScope(scope.parent)
    ) {
      const declaration = localBindings.get(scope)?.get(identifier.text);
      if (declaration) return declaration;
    }
    return undefined;
  };
  const chunkOrigin = (expression: ts.Expression, seen = new Set<ts.Node>()): boolean => {
    if (seen.has(expression)) return false;
    seen.add(expression);
    if (
      ts.isParenthesizedExpression(expression) ||
      ts.isAsExpression(expression) ||
      ts.isNonNullExpression(expression)
    )
      return chunkOrigin(expression.expression, seen);
    if (ts.isIdentifier(expression)) {
      const declaration = localDeclaration(expression);
      if (!declaration || seen.has(declaration)) return false;
      seen.add(declaration);
      if (ts.isVariableDeclaration(declaration)) {
        if (
          ts.isVariableDeclarationList(declaration.parent) &&
          ts.isForOfStatement(declaration.parent.parent)
        )
          return true;
        return declaration.initializer !== undefined && chunkOrigin(declaration.initializer, seen);
      }
      const callback = declaration.parent.parent;
      if (!ts.isCallExpression(callback) || property(callback.expression) !== "on") return false;
      const event = callback.arguments.at(0);
      return event !== undefined && ts.isStringLiteral(event) && event.text === "data";
    }
    if (ts.isCallExpression(expression)) {
      const method = property(expression.expression);
      // A Map key selects a complete stored value; its bytes do not flow into that value.
      if (
        method === "get" &&
        ts.isPropertyAccessExpression(expression.expression) &&
        ts.isIdentifier(expression.expression.expression)
      ) {
        const declaration = localDeclaration(expression.expression.expression);
        if (
          declaration &&
          ((declaration.type && /^(?:Readonly)?Map\s*</u.test(declaration.type.getText(source))) ||
            (ts.isVariableDeclaration(declaration) &&
              declaration.initializer &&
              ts.isNewExpression(declaration.initializer) &&
              ts.isIdentifier(declaration.initializer.expression) &&
              declaration.initializer.expression.text === "Map" &&
              localDeclaration(declaration.initializer.expression) === undefined))
        )
          return false;
      }
      if ((method === "slice" || method === "subarray") && loopAncestor(expression)) return true;
      if (
        ts.isPropertyAccessExpression(expression.expression) ||
        ts.isElementAccessExpression(expression.expression)
      ) {
        if (chunkOrigin(expression.expression.expression, seen)) return true;
      }
      return expression.arguments.some((argument) => chunkOrigin(argument, seen));
    }
    if (ts.isConditionalExpression(expression))
      return chunkOrigin(expression.whenTrue, seen) || chunkOrigin(expression.whenFalse, seen);
    return false;
  };
  const chunkConsumer = (node: ts.Node, first: ts.Expression | undefined): boolean => {
    if (first && chunkOrigin(first)) return true;
    for (let parent = node.parent; parent; parent = parent.parent) {
      if (
        ts.isForOfStatement(parent) &&
        first &&
        ts.isIdentifier(first) &&
        ts.isVariableDeclarationList(parent.initializer) &&
        parent.initializer.declarations.some(
          ({ name }) => ts.isIdentifier(name) && name.text === first.text,
        )
      )
        return true;
      if (ts.isCallExpression(parent) && property(parent.expression) === "on") {
        const event = parent.arguments.at(0);
        if (event && ts.isStringLiteral(event) && event.text === "data") return true;
      }
    }
    return false;
  };
  for (const node of nodes) {
    if (ts.isIdentifier(node) && codecs.has(node.text)) add(node, "raw text codec reference");
    if (
      (ts.isPropertyAccessExpression(node) || ts.isElementAccessExpression(node)) &&
      codecs.has(property(node) ?? "")
    )
      add(node, "raw text codec property");
    if (!ts.isCallExpression(node)) continue;
    const method = property(node.expression);
    const first = node.arguments.at(0);
    const name = ts.isIdentifier(node.expression) ? node.expression.text : method;
    const argument = first?.getText(source) ?? "";
    const options = node.arguments.at(1);
    const binaryField =
      options &&
      ts.isObjectLiteralExpression(options) &&
      options.properties.some(
        (option) =>
          ts.isPropertyAssignment(option) &&
          option.name.getText(source) === "encoding" &&
          ts.isStringLiteral(option.initializer) &&
          option.initializer.text === "binary",
      );
    if (name && binaryCodecs.has(name))
      add(node, "binary string codec must use the encoding owner");
    if (method === "fromCharCode" || method === "fromCodePoint") {
      if (/Uint8Array|bytes|buffer|chunk|view/iu.test(argument))
        add(node, "manual byte-to-text codec");
    }
    if (method === "charCodeAt") {
      for (let parent = node.parent; parent; parent = parent.parent) {
        if (
          ts.isCallExpression(parent) &&
          ((property(parent.expression) === "from" &&
            /Uint8Array/u.test(parent.expression.getText(source))) ||
            /^set(?:Int|Uint)(?:8|16|32)$/u.test(property(parent.expression) ?? ""))
        ) {
          add(node, "manual binary string-to-byte codec");
          break;
        }
        if (
          ts.isBinaryExpression(parent) &&
          parent.operatorToken.kind === ts.SyntaxKind.EqualsToken &&
          ts.isElementAccessExpression(parent.left) &&
          /bytes|buffer|view|encoded/iu.test(parent.left.expression.getText(source))
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
      (!ts.isStringLiteral(first) || first.text !== "hex") &&
      !ts.isNumericLiteral(first)
    )
      add(node, "textual Buffer.toString encoding");
    if (
      method === "from" &&
      (ts.isPropertyAccessExpression(node.expression) ||
        ts.isElementAccessExpression(node.expression)) &&
      ts.isIdentifier(node.expression.expression) &&
      buffers.has(node.expression.expression.text)
    ) {
      if (
        node.arguments.length > 1 ||
        (first &&
          (ts.isStringLiteral(first) ||
            ts.isTemplateExpression(first) ||
            ts.isNoSubstitutionTemplateLiteral(first) ||
            (ts.isIdentifier(first) && textNames.has(first.text))))
      )
        add(node, "textual Buffer.from encoding");
    }
    if (
      name &&
      bufferFactories.has(name) &&
      (node.arguments.length > 1 ||
        (first &&
          (ts.isStringLiteral(first) ||
            ts.isTemplateExpression(first) ||
            ts.isNoSubstitutionTemplateLiteral(first) ||
            (ts.isIdentifier(first) && textNames.has(first.text)))))
    )
      add(node, "aliased textual Buffer.from encoding");
  }
  return violations;
};

/** wasm-bindgen decodes one complete allocation, never independent transport chunks. */
const generatedViolations = (text: string, file: string): string[] => {
  const source = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
  const problems: string[] = [];
  let decoderConstructions = 0;
  let encoderConstructions = 0;
  let decodes = 0;
  let encodes = 0;
  let encodeInto = 0;
  const compact = (node: ts.Node): string => node.getText(source).replace(/\s+/gu, "");
  const visit = (node: ts.Node): void => {
    if (ts.isNewExpression(node)) {
      if (node.expression.getText(source) === "TextDecoder") decoderConstructions++;
      if (node.expression.getText(source) === "TextEncoder") encoderConstructions++;
    }
    if (ts.isCallExpression(node)) {
      const method = property(node.expression);
      if (method === "decode") {
        decodes++;
        const args = node.arguments.map(compact);
        if (
          args.length !== 0 &&
          (args.length !== 1 || args[0] !== "getUint8ArrayMemory0().subarray(ptr,ptr+len)")
        )
          problems.push("decoder no longer reads one complete memory allocation");
      }
      if (method === "encode") {
        encodes++;
        if (node.arguments.length !== 1 || compact(node.arguments[0]!) !== "arg")
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
    ts.forEachChild(node, visit);
  };
  visit(source);
  const detected = findViolations(text, file);
  const rawReferences = detected.filter(({ reason }) => reason.startsWith("raw text codec"));
  if (rawReferences.length !== (file === GENERATED[0] ? 2 : 3))
    problems.push("generated codec references changed");
  if (detected.some(({ reason }) => !reason.startsWith("raw text codec")))
    problems.push("additional encoding escaped the generated whole-buffer codecs");
  if (decoderConstructions !== 2 || decodes !== 3) problems.push("generated decoder shape changed");
  if (file === GENERATED[0]) {
    if (encoderConstructions !== 0 || encodes !== 0 || encodeInto !== 0)
      problems.push("unexpected kernel encoder");
    return problems;
  }
  if (encoderConstructions !== 1 || encodes !== 2 || encodeInto !== 1)
    problems.push("generated shaper encoder shape changed");
  const passString = source.statements.find(
    (node) => ts.isFunctionDeclaration(node) && node.name?.text === "passStringToWasm0",
  );
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
  return findViolations(source, file).map(({ line, reason }) => `${line}: ${reason}`);
};

describe("production text encoding boundaries", () => {
  test("every production package and source extension rejects local codecs and chunk decoding", async () => {
    const packages = new Set<string>();
    for await (const file of new Bun.Glob(SOURCE_GLOB).scan({ cwd: PACKAGES }))
      packages.add(file.split("/")[0]!);
    const bad = [
      "new TextEncoder().encode(text)",
      "const Codec = TextDecoder; new Codec().decode(chunk)",
      "const { TextDecoder: Codec } = require('node:util'); new Codec()",
      "function consume(codec: TextDecoder) { return codec.decode(bytes); }",
      "import { TextEncoder as Codec } from 'node:util'; new Codec()",
      "new globalThis['TextEncoder']()",
      "import { decodeText as decode } from '@stll/docx-core/encoding'; for (const chunk of chunks) decode(chunk)",
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
          const source = extension === "vue" ? `<script setup>${fixture}</script>` : fixture;
          expect({ file, fixture, rejected: sourceViolations(file, source).length > 0 }).toEqual({
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
        sourceViolations(file, source + "\nnew TextDecoder().decode(chunk);").length,
      ).toBeGreaterThan(0);
      expect(
        sourceViolations(
          file,
          source.replace("subarray(ptr, ptr + len)", "subarray(ptr, ptr + len - 1)"),
        ).length,
      ).toBeGreaterThan(0);
      if (file === GENERATED[1])
        expect(
          sourceViolations(file, source.replace("code > 0x7F", "code > 0xFFFF")).length,
        ).toBeGreaterThan(0);
    }
  });

  test("production encodings flow through the owner", async () => {
    const offenders: string[] = [];
    for await (const file of new Bun.Glob(SOURCE_GLOB).scan({ cwd: PACKAGES })) {
      const source = await Bun.file(path.join(PACKAGES, file)).text();
      for (const reason of sourceViolations(file, source)) offenders.push(`${file}:${reason}`);
    }
    expect(offenders).toEqual([]);
  }, 30_000);
});
