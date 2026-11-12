// Bound every test that starts a child process by what the child does.
//
// A test's child runs through `runChild` from `test/subprocess.ts`, which
// reports whether it exited (with its status and output), hung or was killed.
// Its only clock is a hang guard from the shared scaled budget, and
// `subprocessTestTimeout(children)` gives the test room for its children's
// guards, so a hung child is reported by its own test. A bare wall-clock limit
// instead fails a slow runner on a child that would have finished, and the
// runner's 5 s default reports a hung child as a generic test timeout.
//
// In a test file that starts a child (through `Bun.spawn`, `Bun.spawnSync`,
// `spawn`, `spawnSync`, `execFile`, `execFileSync`, `Bun.$` or `runChild`):
//
// Flagged:
//   setDefaultTimeout(30_000)                       (a bare default)
//   test("x", () => runChild(cmd), 5_000)           (a bare test limit)
//   test("x", { timeout: LIMIT_MS }, () => ...)     (a named number is bare too)
//   Bun.spawnSync(cmd, { timeout: 5_000 })          (a bare child limit)
//   runChild(cmd, { hangTimeoutMs: 5_000 })         (a bare hang guard)
//   test("x", () => check())                        (check starts a child, and the
//                                                    file's last default is not
//                                                    subprocessTestTimeout)
//   beforeAll(() => runChild(cmd))                  (the same for a hook)
//
// Not flagged:
//   setDefaultTimeout(subprocessTestTimeout())      (and a test that relies on it)
//   test("x", () => check(), subprocessTestTimeout(2))
//   test("x", { timeout: subprocessTestTimeout() }, () => check())
//   test("y", () => {}, propertyTestTimeout(5_000)) (a test without a child)

type AstNode = Record<string, unknown> & { type: string };

type MessageId = "bareLimit" | "unboundedChildTest";

type RuleContext = {
  report: (descriptor: { node: unknown; messageId: MessageId }) => void;
};

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

// `$` is the shell template tag (`Bun.$` or an imported `$`).
const SUBPROCESS_METHODS: ReadonlySet<string> = new Set([
  "$",
  "execFile",
  "execFileSync",
  "runChild",
  "spawn",
  "spawnSync",
]);
const SUBPROCESS_BUDGET = "subprocessTestTimeout";
const TEST_FUNCTIONS: ReadonlySet<string> = new Set(["it", "test"]);
const HOOK_FUNCTIONS: ReadonlySet<string> = new Set([
  "afterAll",
  "afterEach",
  "beforeAll",
  "beforeEach",
]);
const TIMEOUT_KEYS: ReadonlySet<string> = new Set(["hangTimeoutMs", "timeout"]);

const children = (node: AstNode): AstNode[] => {
  const found: AstNode[] = [];
  for (const [key, value] of Object.entries(node)) {
    if (key === "parent") continue;
    if (Array.isArray(value)) {
      for (const item of value) if (isAstNode(item)) found.push(item);
    } else if (isAstNode(value)) {
      found.push(value);
    }
  }
  return found;
};

/** Visit `root` and its descendants; a visitor returning false skips the children. */
const walk = (root: AstNode, visit: (node: AstNode) => boolean): void => {
  const stack = [root];
  for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
    if (visit(node)) stack.push(...children(node));
  }
};

const identifierName = (node: unknown): string | undefined =>
  isAstNode(node) &&
  node.type === "Identifier" &&
  typeof node["name"] === "string"
    ? node["name"]
    : undefined;

/** `name` for `name(...)`, and the property for `a.b.name(...)`. */
const callName = (callee: unknown): string | undefined => {
  if (!isAstNode(callee)) return undefined;
  if (callee.type === "Identifier") return identifierName(callee);
  if (callee.type !== "MemberExpression" || callee["computed"] === true) {
    return undefined;
  }
  return identifierName(callee["property"]);
};

/** `test` for `test(...)`, `test.each(...)(...)` and `test.skip(...)`. */
const calleeRoot = (callee: unknown): string | undefined => {
  if (!isAstNode(callee)) return undefined;
  if (callee.type === "Identifier") return identifierName(callee);
  if (callee.type === "MemberExpression") return calleeRoot(callee["object"]);
  if (callee.type === "CallExpression") return calleeRoot(callee["callee"]);
  return undefined;
};

const calledName = (node: AstNode): string | undefined => {
  if (node.type === "CallExpression") return callName(node["callee"]);
  if (node.type === "TaggedTemplateExpression") return callName(node["tag"]);
  return undefined;
};

const callsAny = (root: AstNode, names: ReadonlySet<string>): boolean => {
  let found = false;
  walk(root, (node) => {
    if (found) return false;
    const name = calledName(node);
    if (name !== undefined && names.has(name)) found = true;
    return !found;
  });
  return found;
};

const argumentsOf = (node: AstNode): unknown[] =>
  Array.isArray(node["arguments"]) ? node["arguments"] : [];

const isBudgetCall = (node: unknown): boolean =>
  isAstNode(node) &&
  node.type === "CallExpression" &&
  callName(node["callee"]) === SUBPROCESS_BUDGET;

/** A number written in place: a literal, arithmetic on literals or a constant naming one. */
const isBareNumber = (
  node: unknown,
  constants: ReadonlySet<string>,
): boolean => {
  if (!isAstNode(node)) return false;
  switch (node.type) {
    case "Literal":
      return typeof node["value"] === "number";
    case "BinaryExpression":
      return (
        isBareNumber(node["left"], constants) &&
        isBareNumber(node["right"], constants)
      );
    case "Identifier":
      return constants.has(identifierName(node) ?? "");
    case "TSAsExpression":
    case "TSSatisfiesExpression":
    case "ParenthesizedExpression":
      return isBareNumber(node["expression"], constants);
    default:
      return false;
  }
};

const statementsOf = (program: AstNode): AstNode[] =>
  Array.isArray(program["body"]) ? program["body"].filter(isAstNode) : [];

/** Top-level `const NAME = <number>` bindings, which a limit may name. */
const numericConstants = (program: AstNode): ReadonlySet<string> => {
  const names = new Set<string>();
  for (const statement of statementsOf(program)) {
    const declaration =
      statement.type === "ExportNamedDeclaration"
        ? statement["declaration"]
        : statement;
    if (!isAstNode(declaration) || declaration.type !== "VariableDeclaration") {
      continue;
    }
    const declarators = declaration["declarations"];
    if (!Array.isArray(declarators)) continue;
    for (const declarator of declarators) {
      if (!isAstNode(declarator)) continue;
      const name = identifierName(declarator["id"]);
      if (name !== undefined && isBareNumber(declarator["init"], names)) {
        names.add(name);
      }
    }
  }
  return names;
};

/**
 * Names whose call starts a child: the subprocess methods and every function
 * in the file that calls one of them, directly or through another.
 */
const childStarters = (program: AstNode): ReadonlySet<string> => {
  const functions = new Map<string, AstNode>();
  walk(program, (node) => {
    if (node.type === "FunctionDeclaration") {
      const name = identifierName(node["id"]);
      if (name !== undefined) functions.set(name, node);
    } else if (node.type === "VariableDeclarator") {
      const name = identifierName(node["id"]);
      const init = node["init"];
      if (
        name !== undefined &&
        isAstNode(init) &&
        (init.type === "ArrowFunctionExpression" ||
          init.type === "FunctionExpression")
      ) {
        functions.set(name, init);
      }
    }
    return true;
  });
  const starters = new Set(SUBPROCESS_METHODS);
  let grew = true;
  while (grew) {
    grew = false;
    for (const [name, body] of functions) {
      if (!starters.has(name) && callsAny(body, starters)) {
        starters.add(name);
        grew = true;
      }
    }
  }
  return starters;
};

const timeoutOption = (options: unknown): unknown => {
  if (!isAstNode(options) || options.type !== "ObjectExpression") {
    return options;
  }
  const properties = options["properties"];
  if (!Array.isArray(properties)) return undefined;
  for (const property of properties) {
    if (
      isAstNode(property) &&
      property.type === "Property" &&
      property["computed"] !== true &&
      identifierName(property["key"]) === "timeout"
    ) {
      return property["value"];
    }
  }
  return undefined;
};

/** A test's callback and limit, from `(name, fn, limit)` or `(name, options, fn)`. */
const testParts = (
  args: readonly unknown[],
): { callback: unknown; limit: unknown } => {
  const [, second, third] = args;
  if (isAstNode(second) && second.type === "ObjectExpression") {
    return { callback: third, limit: timeoutOption(second) };
  }
  return {
    callback: second,
    limit: third === undefined ? undefined : timeoutOption(third),
  };
};

/** Report every limit and child-starting test in a file that is not bounded by its children. */
const checkProgram = (program: AstNode, context: RuleContext): void => {
  if (!callsAny(program, SUBPROCESS_METHODS)) return;
  const starters = childStarters(program);
  const constants = numericConstants(program);
  // Bun applies the last top-level default to every test that states none.
  let lastDefault: unknown;
  for (const statement of statementsOf(program)) {
    const expression = statement["expression"];
    if (
      statement.type === "ExpressionStatement" &&
      isAstNode(expression) &&
      expression.type === "CallExpression" &&
      callName(expression["callee"]) === "setDefaultTimeout"
    ) {
      lastDefault = argumentsOf(expression)[0];
    }
  }
  const defaultIsBudget = isBudgetCall(lastDefault);
  const bounded = (limit: unknown): boolean =>
    limit === undefined ? defaultIsBudget : isBudgetCall(limit);
  const startsChild = (callback: unknown): boolean =>
    isAstNode(callback) && callsAny(callback, starters);
  walk(program, (node) => {
    if (
      node.type === "Property" &&
      node["computed"] !== true &&
      TIMEOUT_KEYS.has(identifierName(node["key"]) ?? "") &&
      isBareNumber(node["value"], constants)
    ) {
      context.report({ node, messageId: "bareLimit" });
      return true;
    }
    if (node.type !== "CallExpression") return true;
    const args = argumentsOf(node);
    if (
      callName(node["callee"]) === "setDefaultTimeout" &&
      isBareNumber(args[0], constants)
    ) {
      context.report({ node, messageId: "bareLimit" });
      return true;
    }
    const root = calleeRoot(node["callee"]);
    if (root !== undefined && HOOK_FUNCTIONS.has(root)) {
      const [callback, limit] = args;
      if (startsChild(callback) && !bounded(limit)) {
        context.report({ node, messageId: "unboundedChildTest" });
      }
      return true;
    }
    if (root === undefined || !TEST_FUNCTIONS.has(root)) return true;
    const { callback, limit } = testParts(args);
    if (limit !== undefined && isBareNumber(limit, constants)) {
      // A limit in an options object is reported as its `timeout` property.
      if (limit === args[2])
        context.report({ node: limit, messageId: "bareLimit" });
    } else if (startsChild(callback) && !bounded(limit)) {
      context.report({ node, messageId: "unboundedChildTest" });
    }
    return true;
  });
};

export default {
  meta: { name: "folio-subprocess-bounds" },
  rules: {
    "bounded-child-tests": {
      meta: {
        type: "problem",
        messages: {
          bareLimit:
            "A test file that starts a child process states no bare wall-clock " +
            "limit: a slow runner fails a child that would have finished. Run " +
            "the child through `runChild` from `test/subprocess.ts` (its hang " +
            "guard comes from the shared scaled budget) and bound the test with " +
            "`subprocessTestTimeout(children)`, or a test without a child with " +
            "`scanTestTimeout` or `propertyTestTimeout`.",
          unboundedChildTest:
            "This test starts a child process but neither its own limit nor the " +
            "file's last `setDefaultTimeout` is `subprocessTestTimeout`, so a " +
            "hung child ends as a generic test timeout instead of its own " +
            "report. Give it `subprocessTestTimeout(children)` from " +
            "`test/subprocess.ts`, as its limit or as the file's last default.",
        },
      },
      create(context: RuleContext) {
        return {
          Program: (node: unknown) => {
            if (isAstNode(node)) checkProgram(node, context);
          },
        };
      },
    },
  },
};
