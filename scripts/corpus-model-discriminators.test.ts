import { readdir } from "node:fs/promises";
import path from "node:path";

import { describe, expect, test } from "bun:test";
import ts from "typescript";

import { MODEL_TYPE_DISCRIMINATORS } from "./lib/corpus-signature";

/**
 * The discriminators a signature may name, derived from the model rather than
 * listed.
 *
 * A path segment names the kind of the element it steps into, and that kind is
 * read out of the parsed package at `type`. `type` is an ordinary key: a
 * package folio did not write can carry any string under it, including one
 * lifted from the document's own text. So the walk names a kind only when the
 * model declares it, and the closed set it checks against has to be the
 * model's, not a copy of it that drifted.
 *
 * Derivation, not a hand list: every array and every `Map` value in
 * `packages/docx-core/src/model` is an element position, because
 * `normalizeDocumentPackage` turns a `Map` into an array of pairs and the walk
 * steps into both the same way. The element type is expanded through the
 * aliases and unions the model writes, and each object member's literal `type`
 * is collected.
 */
const MODEL_DIR = path.join(import.meta.dir, "..", "packages", "docx-core", "src", "model");

const modelSourceFiles = async (): Promise<ts.SourceFile[]> => {
  const names = (await readdir(MODEL_DIR))
    .filter((name) => name.endsWith(".ts") && !name.includes(".test."))
    .sort();
  return Promise.all(
    names.map(async (name) =>
      ts.createSourceFile(
        name,
        await Bun.file(path.join(MODEL_DIR, name)).text(),
        ts.ScriptTarget.ESNext,
        true,
        ts.ScriptKind.TS,
      ),
    ),
  );
};

const eachNode = (node: ts.Node, visit: (child: ts.Node) => void): void => {
  visit(node);
  ts.forEachChild(node, (child) => {
    eachNode(child, visit);
  });
};

/** Every `type X = …` in the model, so a reference can be followed. */
const aliasesIn = (sources: readonly ts.SourceFile[]): Map<string, ts.TypeNode> => {
  const aliases = new Map<string, ts.TypeNode>();
  for (const source of sources) {
    eachNode(source, (node) => {
      if (ts.isTypeAliasDeclaration(node)) {
        aliases.set(node.name.text, node.type);
      }
    });
  }
  return aliases;
};

/**
 * Which type argument of a generic holds the element, by generic.
 *
 * `Map` is here because the projection normalises a map into an array of
 * `[key, value]` pairs, so a map value is an array element by the time a path
 * is built.
 */
const ELEMENT_ARGUMENT: ReadonlyMap<string, number> = new Map([
  ["Array", 0],
  ["ReadonlyArray", 0],
  ["Map", 1],
  ["ReadonlyMap", 1],
]);

/** Every type node the walk can reach as an element of an array. */
const elementPositionsIn = (sources: readonly ts.SourceFile[]): ts.TypeNode[] => {
  const positions: ts.TypeNode[] = [];
  const collect = (node: ts.Node): void => {
    if (!ts.isPropertySignature(node) || node.type === undefined) {
      return;
    }
    const declared = node.type;
    if (ts.isArrayTypeNode(declared)) {
      positions.push(declared.elementType);
      return;
    }
    if (!ts.isTypeReferenceNode(declared) || !ts.isIdentifier(declared.typeName)) {
      return;
    }
    const argument = ELEMENT_ARGUMENT.get(declared.typeName.text);
    if (argument === undefined) {
      return;
    }
    const element = (declared.typeArguments ?? []).at(argument);
    if (element !== undefined) {
      positions.push(element);
    }
  };
  for (const source of sources) {
    eachNode(source, collect);
  }
  return positions;
};

/** The string literals a `type` member declares, directly or as a union. */
const literalsOf = (declared: ts.TypeNode, into: Set<string>): void => {
  if (ts.isParenthesizedTypeNode(declared)) {
    literalsOf(declared.type, into);
    return;
  }
  if (ts.isUnionTypeNode(declared)) {
    for (const member of declared.types) {
      literalsOf(member, into);
    }
    return;
  }
  if (ts.isLiteralTypeNode(declared) && ts.isStringLiteral(declared.literal)) {
    into.add(declared.literal.text);
  }
};

const discriminatorsOfObject = (declared: ts.TypeLiteralNode, into: Set<string>): void => {
  for (const member of declared.members) {
    if (
      ts.isPropertySignature(member) &&
      member.type !== undefined &&
      ts.isIdentifier(member.name) &&
      member.name.text === "type"
    ) {
      literalsOf(member.type, into);
    }
  }
};

/**
 * `Exclude` and `Extract` narrow a union the model already declares. Both are
 * expanded through their first argument: a narrowing can only remove members,
 * and every member it could remove is reachable from the union it narrows.
 */
const NARROWING_HELPERS = new Set(["Exclude", "Extract", "NonNullable", "Readonly"]);

const expandInto = (
  declared: ts.TypeNode,
  aliases: ReadonlyMap<string, ts.TypeNode>,
  into: Set<string>,
  seen: Set<ts.TypeNode>,
): void => {
  if (seen.has(declared)) {
    return;
  }
  seen.add(declared);
  if (ts.isParenthesizedTypeNode(declared)) {
    expandInto(declared.type, aliases, into, seen);
    return;
  }
  if (ts.isUnionTypeNode(declared) || ts.isIntersectionTypeNode(declared)) {
    for (const member of declared.types) {
      expandInto(member, aliases, into, seen);
    }
    return;
  }
  if (ts.isTypeLiteralNode(declared)) {
    discriminatorsOfObject(declared, into);
    return;
  }
  if (!ts.isTypeReferenceNode(declared) || !ts.isIdentifier(declared.typeName)) {
    return;
  }
  const { text } = declared.typeName;
  const argument = declared.typeArguments?.at(0);
  if (NARROWING_HELPERS.has(text) && argument !== undefined) {
    expandInto(argument, aliases, into, seen);
    return;
  }
  const alias = aliases.get(text);
  if (alias !== undefined) {
    expandInto(alias, aliases, into, seen);
  }
};

const declaredModelDiscriminators = async (): Promise<Set<string>> => {
  const sources = await modelSourceFiles();
  const aliases = aliasesIn(sources);
  const discriminators = new Set<string>();
  const seen = new Set<ts.TypeNode>();
  for (const position of elementPositionsIn(sources)) {
    expandInto(position, aliases, discriminators, seen);
  }
  return discriminators;
};

const sorted = (values: Iterable<string>): string[] => [...values].sort();

describe("the discriminators a path segment may name", () => {
  test("are exactly the ones the model declares for an element position", async () => {
    expect(sorted(MODEL_TYPE_DISCRIMINATORS)).toEqual(sorted(await declaredModelDiscriminators()));
  });

  // The derivation is only worth trusting if it finds the model. An empty set
  // would agree with an empty constant and prove nothing.
  test("the derivation reads the model it claims to read", async () => {
    const declared = await declaredModelDiscriminators();
    expect(declared.has("run")).toBe(true);
    expect(declared.has("paragraph")).toBe(true);
    expect(declared.has("tableRow")).toBe(true);
  });
});
