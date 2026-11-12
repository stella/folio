// Assertions over what a property happened to draw.
//
// A property's generated pass searches a sequence its seed chooses, and it can
// stop after any case: a failure ends it, and under pull request CI an
// exploration finding ends it without failing (test/property-exploration.ts).
// A tally that a property's predicate writes (a counter, a Set or Map of the
// kinds it saw, a list it pushes to) therefore holds whatever that seed
// reached before the run stopped. An assertion that such a tally covers every
// kind, or reaches some count, fails on seeds that never drew a kind and on
// runs that ended early, whatever the code under test does.
//
// Coverage a property must reach is stated without the draw instead: fixed
// `examples` per kind, a generator that always yields each kind, or the same
// predicate run on a fixed sample (test/property-coverage.ts). This module
// finds assertions that read a tally written where values are drawn:
// inside an `fc.property` / `fc.asyncProperty` predicate (or a callback among
// its arguments), or in a loop over an `fc.sample` that names no seed.
//
// Syntax only. Names resolve to their declarations by lexical scope within the
// module, and functions declared in the module are followed into: a helper
// that counts into the tally it is given writes that tally, and a helper that
// asserts on the tally it is given reads it.
import {
  collectSyntaxTree,
  createLineIndex,
  type Node,
  parseSourceModule,
  staticMemberName,
  walkAst,
} from "./oxc-ast";

/** Methods that write the collection they are called on. */
const MUTATING_METHODS = new Set([
  "add",
  "set",
  "push",
  "unshift",
  "splice",
  "delete",
  "clear",
  "fill",
  "sort",
  "reverse",
  "copyWithin",
]);

const PROPERTY_BUILDERS = new Set(["property", "asyncProperty"]);
const PROPERTY_RUNNERS = new Set(["assertProperty", "assertKnownProperty"]);
/** Runs only the seeds test/property-seeds/ pins: the same cases everywhere. */
const PINNED_RUNNER = "assertPinnedProperty";
const ITERATING_METHODS = new Set(["forEach", "map", "flatMap", "reduce"]);

/** Matchers that only hold while nothing was reached: a truncated tally cannot fail them. */
const EMPTINESS_MATCHERS = new Set([
  "toBeUndefined",
  "toBeNull",
  "toBeEmpty",
  "toBeFalsy",
]);
const EQUALITY_MATCHERS = new Set(["toBe", "toEqual", "toStrictEqual"]);
const EQUALITY_ASSERTS = new Set([
  "equal",
  "strictEqual",
  "deepEqual",
  "deepStrictEqual",
]);

export type DrawCoverageFinding = {
  /** `file:line` of the assertion. */
  readonly site: string;
  /** The tallies the draw writes and the assertion reads. */
  readonly tallies: readonly string[];
};

type FunctionNode = Extract<
  Node,
  {
    type:
      "FunctionDeclaration" | "FunctionExpression" | "ArrowFunctionExpression";
  }
>;
type CallNode = Extract<Node, { type: "CallExpression" }>;

const isFunction = (node: Node): node is FunctionNode =>
  node.type === "FunctionDeclaration" ||
  node.type === "FunctionExpression" ||
  node.type === "ArrowFunctionExpression";

const unwrap = (node: Node): Node => {
  let current = node;
  while (
    current.type === "ParenthesizedExpression" ||
    current.type === "TSAsExpression" ||
    current.type === "TSSatisfiesExpression" ||
    current.type === "TSNonNullExpression" ||
    current.type === "ChainExpression" ||
    current.type === "AwaitExpression"
  ) {
    current =
      current.type === "AwaitExpression"
        ? current.argument
        : current.expression;
  }
  return current;
};

/** The identifier an access such as `a.b[c].d` or `a.get(k).d` starts from. */
const rootIdentifier = (
  node: Node,
): Extract<Node, { type: "Identifier" }> | undefined => {
  const target = unwrap(node);
  if (target.type === "Identifier") return target;
  if (target.type === "MemberExpression") return rootIdentifier(target.object);
  if (
    target.type === "CallExpression" &&
    unwrap(target.callee).type === "MemberExpression"
  )
    return rootIdentifier(target.callee);
  return undefined;
};

/** Identifiers a binding pattern declares. */
const boundIdentifiers = (pattern: Node): string[] => {
  const names: string[] = [];
  walkAst(pattern, (node, ancestors) => {
    const parent = ancestors.at(-1);
    if (node.type === "TSTypeAnnotation") return false;
    if (node.type !== "Identifier") return;
    if (parent?.type === "Property" && parent.key === node && !parent.shorthand)
      return;
    if (parent?.type === "AssignmentPattern" && parent.right === node) return;
    if (parent?.type === "MemberExpression") return;
    names.push(node.name);
  });
  return names;
};

const declaredByStatement = (statement: Node, into: Set<string>): void => {
  let declaration: Node | null | undefined = statement;
  if (
    statement.type === "ExportNamedDeclaration" ||
    statement.type === "ExportDefaultDeclaration"
  )
    declaration = statement.declaration;
  if (declaration === null || declaration === undefined) return;
  if (declaration.type === "VariableDeclaration")
    for (const declarator of declaration.declarations)
      for (const name of boundIdentifiers(declarator.id)) into.add(name);
  if (
    (declaration.type === "FunctionDeclaration" ||
      declaration.type === "ClassDeclaration") &&
    declaration.id
  )
    into.add(declaration.id.name);
  if (declaration.type === "ImportDeclaration")
    for (const specifier of declaration.specifiers)
      into.add(specifier.local.name);
};

/** The names a scope node declares directly; `undefined` for nodes that open no scope. */
const scopeNames = (node: Node): Set<string> | undefined => {
  const names = new Set<string>();
  switch (node.type) {
    case "Program":
    case "BlockStatement":
    case "StaticBlock":
      for (const statement of node.body) declaredByStatement(statement, names);
      return names;
    case "SwitchStatement":
      for (const branch of node.cases)
        for (const statement of branch.consequent)
          declaredByStatement(statement, names);
      return names;
    case "FunctionDeclaration":
    case "FunctionExpression":
    case "ArrowFunctionExpression":
      for (const parameter of node.params)
        for (const name of boundIdentifiers(parameter)) names.add(name);
      if (node.type === "FunctionExpression" && node.id)
        names.add(node.id.name);
      return names;
    case "ForStatement":
      if (node.init?.type === "VariableDeclaration")
        declaredByStatement(node.init, names);
      return names;
    case "ForInStatement":
    case "ForOfStatement":
      if (node.left.type === "VariableDeclaration")
        declaredByStatement(node.left, names);
      return names;
    case "CatchClause":
      if (node.param)
        for (const name of boundIdentifiers(node.param)) names.add(name);
      return names;
    default:
      return undefined;
  }
};

const within = (inner: Node, outer: Node): boolean =>
  inner.start >= outer.start && inner.end <= outer.end;

/** A declaration: the scope that declares a name, and the name. */
type Binding = { readonly scope: Node; readonly name: string };
const keyOf = ({ scope, name }: Binding): string =>
  `${scope.type}@${String(scope.start)}-${String(scope.end)}:${name}`;

type Facts = {
  /** Bindings the function writes, its own parameters included. */
  readonly writes: Set<string>;
  /** Bindings an assertion in the function reads. */
  readonly asserts: Set<string>;
};

/** `{ seed: … }` in the parameters of a runner call: the run is the same everywhere. */
const namesSeed = (params: Node | undefined): boolean => {
  if (params === undefined) return false;
  let target = unwrap(params);
  if (
    target.type === "CallExpression" &&
    unwrap(target.callee).type === "Identifier"
  )
    target =
      target.arguments[0] === undefined ? target : unwrap(target.arguments[0]);
  return (
    target.type === "ObjectExpression" &&
    target.properties.some(
      (property) =>
        property.type === "Property" &&
        !property.computed &&
        property.key.type === "Identifier" &&
        property.key.name === "seed",
    )
  );
};

const isFastCheckCall = (node: Node, names: ReadonlySet<string>): boolean => {
  if (node.type !== "CallExpression") return false;
  const callee = unwrap(node.callee);
  return (
    callee.type === "MemberExpression" &&
    names.has(staticMemberName(callee) ?? "") &&
    rootIdentifier(callee)?.name === "fc"
  );
};

const isExpectCall = (node: Node): node is CallNode => {
  if (node.type !== "CallExpression") return false;
  const callee = unwrap(node.callee);
  if (callee.type === "Identifier") return callee.name === "expect";
  // `expect.soft(...)`
  return (
    callee.type === "MemberExpression" &&
    unwrap(callee.object).type === "Identifier" &&
    rootIdentifier(callee)?.name === "expect"
  );
};

/** `assert(...)`, `assert.ok(...)`, `assert.equal(...)` and the rest of node:assert. */
const isAssertCall = (node: Node): node is CallNode => {
  if (node.type !== "CallExpression") return false;
  const callee = unwrap(node.callee);
  if (callee.type === "Identifier") return callee.name === "assert";
  return (
    callee.type === "MemberExpression" &&
    rootIdentifier(callee)?.name === "assert"
  );
};

/** `0`, `false`, `[]`, `{}`, `new Set()`, `new Map()`. */
const isEmptyValue = (node: Node | undefined): boolean => {
  if (node === undefined) return false;
  const value = unwrap(node);
  if (value.type === "Literal")
    return value.value === 0 || value.value === false;
  if (value.type === "ArrayExpression") return value.elements.length === 0;
  if (value.type === "ObjectExpression") return value.properties.length === 0;
  return (
    value.type === "NewExpression" &&
    value.arguments.length === 0 &&
    value.callee.type === "Identifier" &&
    (value.callee.name === "Set" || value.callee.name === "Map")
  );
};

/**
 * Collect the facts of one module: where values are drawn, what those places
 * write, and which assertions read it.
 */
export const drawCoverageAssertions = (
  file: string,
  sourceText: string,
): DrawCoverageFinding[] => {
  const { program } = parseSourceModule(file, sourceText);
  const lineOf = createLineIndex(sourceText);
  const { nodes, parentOf } = collectSyntaxTree(program);

  const declared = new Map<Node, Set<string> | undefined>();
  const namesOf = (node: Node) => {
    if (!declared.has(node)) declared.set(node, scopeNames(node));
    return declared.get(node);
  };
  const resolve = (
    identifier: Node & { name: string },
  ): Binding | undefined => {
    for (
      let scope = parentOf(identifier);
      scope !== undefined;
      scope = parentOf(scope)
    ) {
      if (namesOf(scope)?.has(identifier.name))
        return { scope, name: identifier.name };
    }
    return undefined;
  };
  const scopeOfKey = new Map<string, Node>();
  const bindingKey = (binding: Binding): string => {
    const key = keyOf(binding);
    scopeOfKey.set(key, binding.scope);
    return key;
  };
  const resolveKey = (node: Node | undefined): string | undefined => {
    const identifier = node === undefined ? undefined : rootIdentifier(node);
    const binding = identifier === undefined ? undefined : resolve(identifier);
    return binding === undefined ? undefined : bindingKey(binding);
  };

  // Functions by the binding that names them.
  const functionOf = new Map<string, FunctionNode>();
  for (const node of nodes) {
    if (node.type === "FunctionDeclaration" && node.id) {
      const binding = resolve(node.id);
      if (binding) functionOf.set(bindingKey(binding), node);
    }
    if (
      node.type === "VariableDeclarator" &&
      node.id.type === "Identifier" &&
      node.init
    ) {
      const init = unwrap(node.init);
      const binding = resolve(node.id);
      if (isFunction(init) && binding)
        functionOf.set(bindingKey(binding), init);
    }
  }
  const calledFunction = (call: CallNode): FunctionNode | undefined => {
    const callee = unwrap(call.callee);
    if (callee.type !== "Identifier") return undefined;
    const key = resolveKey(callee);
    return key === undefined ? undefined : functionOf.get(key);
  };
  const parameterKeys = (fn: FunctionNode): (string | undefined)[] =>
    fn.params.map((parameter) => {
      const target =
        parameter.type === "AssignmentPattern" ? parameter.left : parameter;
      return target.type === "Identifier"
        ? bindingKey({ scope: fn, name: target.name })
        : undefined;
    });
  /** Bindings declared inside `owner` (its parameters included). */
  const isLocalTo = (key: string, owner: Node): boolean => {
    const scope = scopeOfKey.get(key);
    return scope !== undefined && within(scope, owner);
  };

  // Places where values are drawn, filled in below; their bodies are skipped
  // when reading assertions, and identifiers inside them are not reads.
  const drawSites = new Set<Node>();
  const insideDraw = (node: Node): boolean => {
    for (let at: Node | undefined = node; at; at = parentOf(at))
      if (drawSites.has(at)) return true;
    return false;
  };

  /** Bindings an assertion chain reads, outside the draws it may contain. */
  const readsOf = (chain: Node, { functions = true } = {}): string[] => {
    const keys: string[] = [];
    walkAst(chain, (child, ancestors) => {
      if (drawSites.has(child)) return false;
      if (!functions && isFunction(child)) return false;
      if (child.type !== "Identifier") return;
      const parent = ancestors.at(-1);
      if (
        parent?.type === "MemberExpression" &&
        parent.property === child &&
        !parent.computed
      )
        return;
      if (
        parent?.type === "Property" &&
        parent.key === child &&
        !parent.computed &&
        !parent.shorthand
      )
        return;
      const binding = resolve(child);
      if (binding) keys.push(bindingKey(binding));
    });
    return keys;
  };

  /** The whole `expect(x).not.toEqual(y)` chain an `expect(x)` call starts. */
  const assertionChain = (call: CallNode): Node => {
    let top: Node = call;
    for (let parent = parentOf(top); parent; parent = parentOf(top)) {
      if (
        (parent.type === "MemberExpression" && parent.object === top) ||
        (parent.type === "CallExpression" && parent.callee === top) ||
        parent.type === "ChainExpression" ||
        parent.type === "AwaitExpression"
      ) {
        top = parent;
        continue;
      }
      break;
    }
    return top;
  };

  /** An assertion that only holds while nothing was reached. */
  const assertsEmptiness = (call: CallNode, chain: Node): boolean => {
    if (isAssertCall(call)) {
      const callee = unwrap(call.callee);
      const method =
        callee.type === "MemberExpression"
          ? staticMemberName(callee)
          : undefined;
      return (
        method !== undefined &&
        EQUALITY_ASSERTS.has(method) &&
        isEmptyValue(call.arguments[1])
      );
    }
    if (chain.type !== "CallExpression") return false;
    const matcher = unwrap(chain.callee);
    if (matcher.type !== "MemberExpression") return false;
    // `.not` turns an emptiness matcher into a coverage one.
    for (
      let at: Node = matcher.object;
      at.type === "MemberExpression";
      at = at.object
    )
      if (staticMemberName(at) === "not") return false;
    const name = staticMemberName(matcher) ?? "";
    if (EMPTINESS_MATCHERS.has(name)) return true;
    if (name === "toHaveLength") return isEmptyValue(chain.arguments[0]);
    return EQUALITY_MATCHERS.has(name) && isEmptyValue(chain.arguments[0]);
  };

  const facts = new Map<Node, Facts>();
  /**
   * What `scope` (a function, or a loop over a sample) writes and asserts on,
   * following calls into functions of this module.
   */
  const factsOf = (scope: Node, visiting: Set<Node>): Facts => {
    const known = facts.get(scope);
    if (known !== undefined) return known;
    const writes = new Set<string>();
    const asserts = new Set<string>();
    if (visiting.has(scope)) return { writes, asserts };
    visiting.add(scope);
    const addWrite = (target: Node) => {
      const key = resolveKey(target);
      if (key !== undefined) writes.add(key);
    };
    const throughCall = (call: CallNode) => {
      const target = calledFunction(call);
      if (target === undefined || target === scope) return;
      const inner = factsOf(target, visiting);
      const params = parameterKeys(target);
      const map = (key: string, into: Set<string>) => {
        const index = params.indexOf(key);
        if (index !== -1) {
          const argument = call.arguments[index];
          const mapped =
            argument === undefined ? undefined : resolveKey(argument);
          if (mapped !== undefined) into.add(mapped);
          return;
        }
        if (!isLocalTo(key, target)) into.add(key);
      };
      for (const key of inner.writes) map(key, writes);
      for (const key of inner.asserts) map(key, asserts);
    };
    walkAst(scope, (node) => {
      if (node.type === "UpdateExpression") addWrite(node.argument);
      else if (node.type === "AssignmentExpression") addWrite(node.left);
      else if (node.type === "CallExpression") {
        const callee = unwrap(node.callee);
        if (
          callee.type === "MemberExpression" &&
          MUTATING_METHODS.has(staticMemberName(callee) ?? "")
        )
          addWrite(callee.object);
        if (isExpectCall(node) || isAssertCall(node)) {
          const chain = assertionChain(node);
          if (!assertsEmptiness(node, chain))
            for (const key of readsOf(chain)) asserts.add(key);
        }
        throughCall(node);
      }
    });
    visiting.delete(scope);
    const result = { writes, asserts };
    facts.set(scope, result);
    return result;
  };

  // Draws and the tallies they write.
  const tallies = new Map<string, string>();
  const addTallies = (scope: Node) => {
    for (const key of factsOf(scope, new Set()).writes)
      if (!isLocalTo(key, scope))
        tallies.set(key, key.slice(key.lastIndexOf(":") + 1));
  };
  for (const node of nodes) {
    if (node.type !== "CallExpression") continue;
    if (isFastCheckCall(node, PROPERTY_BUILDERS)) {
      // A runner given a seed of its own replays the same cases everywhere.
      const runner = parentOf(node);
      if (runner?.type === "CallExpression" && runner.arguments[0] === node) {
        const callee = unwrap(runner.callee);
        const name = callee.type === "Identifier" ? callee.name : "";
        if (name === PINNED_RUNNER) continue;
        if (
          (isFastCheckCall(runner, new Set(["assert", "check"])) ||
            PROPERTY_RUNNERS.has(name)) &&
          namesSeed(runner.arguments[1])
        )
          continue;
      }
      const last = node.arguments.at(-1);
      const predicate = last === undefined ? undefined : unwrap(last);
      if (predicate?.type === "Identifier") {
        const key = resolveKey(predicate);
        const fn = key === undefined ? undefined : functionOf.get(key);
        if (fn) {
          drawSites.add(fn);
          addTallies(fn);
        }
      } else if (predicate?.type === "CallExpression") {
        // A factory: what it returns writes what it is given.
        const factory = calledFunction(predicate);
        if (factory) {
          drawSites.add(factory);
          const inner = factsOf(factory, new Set());
          parameterKeys(factory).forEach((param, index) => {
            const argument = predicate.arguments[index];
            const mapped =
              argument === undefined ? undefined : resolveKey(argument);
            if (
              param !== undefined &&
              mapped !== undefined &&
              inner.writes.has(param)
            )
              tallies.set(mapped, mapped.slice(mapped.lastIndexOf(":") + 1));
          });
          for (const key of inner.writes)
            if (!isLocalTo(key, factory))
              tallies.set(key, key.slice(key.lastIndexOf(":") + 1));
        }
      }
      // The predicate and every callback among the arguments (`arb.map(...)`)
      // run once per drawn value.
      for (const argument of node.arguments)
        walkAst(argument, (child) => {
          if (!isFunction(child)) return;
          drawSites.add(child);
          addTallies(child);
          return false;
        });
      continue;
    }
    if (
      isFastCheckCall(node, new Set(["sample"])) &&
      !namesSeed(node.arguments[1])
    ) {
      const parent = parentOf(node);
      if (parent?.type === "ForOfStatement" && parent.right === node) {
        drawSites.add(parent.body);
        addTallies(parent.body);
      }
      if (
        parent?.type === "MemberExpression" &&
        parent.object === node &&
        ITERATING_METHODS.has(staticMemberName(parent) ?? "")
      ) {
        const call = parentOf(parent);
        const callback =
          call?.type === "CallExpression" ? call.arguments[0] : undefined;
        if (callback !== undefined && isFunction(unwrap(callback))) {
          drawSites.add(unwrap(callback));
          addTallies(unwrap(callback));
        }
      }
    }
  }
  // A property a local function builds over what it is given, run by a
  // runner: `assertProperty(law(tally), ...)`, or through `const p = law(tally)`.
  const runnerName = (call: CallNode): string => {
    const callee = unwrap(call.callee);
    if (callee.type === "Identifier") return callee.name;
    return isFastCheckCall(call, new Set(["assert", "check"])) ? "fc" : "";
  };
  const builtProperty = (node: Node | undefined): CallNode | undefined => {
    const target = node === undefined ? undefined : unwrap(node);
    if (target?.type === "CallExpression") return target;
    if (target?.type !== "Identifier") return undefined;
    const key = resolveKey(target);
    for (const declarator of nodes) {
      if (
        declarator.type === "VariableDeclarator" &&
        declarator.id.type === "Identifier" &&
        declarator.init?.type === "CallExpression" &&
        resolveKey(declarator.id) === key
      )
        return declarator.init;
    }
    return undefined;
  };
  for (const node of nodes) {
    if (node.type !== "CallExpression") continue;
    const name = runnerName(node);
    if (!(PROPERTY_RUNNERS.has(name) || name === "fc")) continue;
    if (namesSeed(node.arguments[1])) continue;
    const built = builtProperty(node.arguments[0]);
    const factory = built === undefined ? undefined : calledFunction(built);
    if (built === undefined || factory === undefined) continue;
    parameterKeys(factory).forEach((param, index) => {
      const argument = built.arguments[index];
      const mapped = argument === undefined ? undefined : resolveKey(argument);
      if (param !== undefined && mapped !== undefined && tallies.has(param))
        tallies.set(mapped, tallies.get(param) ?? param);
    });
  }
  if (tallies.size === 0) return [];

  // Values read from a tally are the tally: `const n = tally.get(kind)`,
  // `for (const n of tally.values())`.
  const readsTally = (node: Node | null | undefined): string | undefined => {
    if (node === null || node === undefined) return undefined;
    for (const key of readsOf(node, { functions: false })) {
      const name = tallies.get(key);
      if (name !== undefined) return name;
    }
    return undefined;
  };
  for (let grew = true; grew;) {
    grew = false;
    for (const node of nodes) {
      let pattern: Node | undefined;
      let source: Node | null | undefined;
      if (node.type === "VariableDeclarator") {
        pattern = node.id;
        source = node.init;
      } else if (
        (node.type === "ForOfStatement" || node.type === "ForInStatement") &&
        node.left.type === "VariableDeclaration"
      ) {
        pattern = node.left.declarations[0]?.id;
        source = node.right;
      }
      if (pattern === undefined || insideDraw(node)) continue;
      const name = readsTally(source);
      if (name === undefined) continue;
      walkAst(pattern, (child) => {
        if (child.type !== "Identifier") return;
        const binding = resolve(child);
        if (binding === undefined) return;
        const key = bindingKey(binding);
        if (!tallies.has(key)) {
          tallies.set(key, name);
          grew = true;
        }
      });
    }
  }

  // Assertions outside every draw that read a tally, directly or through a
  // helper of this module.
  const findings: DrawCoverageFinding[] = [];
  const report = (node: Node, keys: Iterable<string>) => {
    const hit = [...new Set(keys)].flatMap((key) => {
      const name = tallies.get(key);
      return name === undefined ? [] : [name];
    });
    if (hit.length === 0) return;
    const { line } = lineOf(node.start);
    findings.push({
      site: `${file}:${String(line + 1)}`,
      tallies: [...new Set(hit)].toSorted(),
    });
  };
  for (const node of nodes) {
    if (node.type !== "CallExpression" || insideDraw(node)) continue;
    if (isExpectCall(node) || isAssertCall(node)) {
      const chain = assertionChain(node);
      if (!assertsEmptiness(node, chain)) report(node, readsOf(chain));
      continue;
    }
    const target = calledFunction(node);
    if (target === undefined || drawSites.has(target)) continue;
    const inner = factsOf(target, new Set());
    const params = parameterKeys(target);
    const keys: string[] = [];
    for (const key of inner.asserts) {
      const index = params.indexOf(key);
      if (index !== -1) {
        const argument = node.arguments[index];
        const mapped =
          argument === undefined ? undefined : resolveKey(argument);
        if (mapped !== undefined) keys.push(mapped);
      } else if (!isLocalTo(key, target)) keys.push(key);
    }
    report(node, keys);
  }
  return findings.toSorted((a, b) =>
    a.site.localeCompare(b.site, "en", { numeric: true }),
  );
};
