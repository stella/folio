/**
 * A block id held across accepting or rejecting changes is read again first.
 *
 * A block's id is its paragraph's `w14:paraId`. Resolving a change that joins
 * a paragraph into the next one leaves the paragraph whose mark stays, the
 * next one, with its id (the reference rule), so an id read before the
 * resolution can name a paragraph that is gone. A caller reads it again after
 * the resolution, by its words, through `heldBlockIds` (or binds the value a
 * fresh read returns), instead of reusing the old binding.
 *
 * A binding holds a block id when its value reads `.id` of a reader's blocks,
 * calls `idOf` or `heldBlockIds`, or is built from another such binding. A
 * resolution is a call of `acceptAll`, `rejectAll`, `acceptChange`,
 * `rejectChange`, `acceptSuggestion` or `rejectSuggestion`, or of a function
 * of the file that makes one. A read of a held binding is reported when a
 * resolution in the function declaring it can run between its last write and
 * the read: earlier in source order, or in a loop the read shares with the
 * resolution and no write of the binding stands in. A read inside a nested
 * function may run after any resolution, so it is reported unless a write
 * follows every resolution.
 */

type AstNode = Record<string, unknown> & {
  type: string;
  parent?: unknown;
  range?: unknown;
};
type LexicalVariable = {
  name: string;
  defs: readonly { node: unknown }[];
  references: readonly {
    identifier: unknown;
    writeExpr?: unknown;
    isRead: () => boolean;
    isWrite: () => boolean;
  }[];
};
type LexicalScope = {
  upper: LexicalScope | null;
  set: Map<string, LexicalVariable>;
};
type RuleContext = {
  filename: string;
  sourceCode: { getScope: (node: AstNode) => LexicalScope };
  report: (descriptor: {
    node: unknown;
    messageId: "staleBlockId";
    data: { name: string };
  }) => void;
};

const RESOLUTIONS: ReadonlySet<string> = new Set([
  "acceptAll",
  "rejectAll",
  "acceptChange",
  "rejectChange",
  "acceptSuggestion",
  "rejectSuggestion",
]);
/** Calls that read block ids from a reader. */
const ID_READERS: ReadonlySet<string> = new Set(["idOf", "heldBlockIds"]);

const isNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof value.type === "string";

const nameOf = (node: unknown): string | undefined => {
  if (!isNode(node)) return undefined;
  if (node.type === "Identifier" && typeof node.name === "string")
    return node.name;
  if (node.type === "Literal" && typeof node.value === "string")
    return node.value;
  return undefined;
};

const startOf = (node: unknown): number => {
  if (!isNode(node) || !Array.isArray(node.range)) return -1;
  const start: unknown = node.range.at(0);
  return typeof start === "number" ? start : -1;
};

const endOf = (node: unknown): number => {
  if (!isNode(node) || !Array.isArray(node.range)) return -1;
  const end: unknown = node.range.at(1);
  return typeof end === "number" ? end : -1;
};

const contains = (outer: AstNode, inner: unknown): boolean => {
  let current: unknown = inner;
  while (isNode(current)) {
    if (current === outer) return true;
    current = current.parent;
  }
  return false;
};

const FUNCTIONS: ReadonlySet<string> = new Set([
  "FunctionDeclaration",
  "FunctionExpression",
  "ArrowFunctionExpression",
]);
const LOOPS: ReadonlySet<string> = new Set([
  "ForStatement",
  "ForInStatement",
  "ForOfStatement",
  "WhileStatement",
  "DoWhileStatement",
]);

const functionOf = (node: unknown): AstNode | undefined => {
  let current: unknown = isNode(node) ? node.parent : undefined;
  while (isNode(current)) {
    if (FUNCTIONS.has(current.type) || current.type === "Program")
      return current;
    current = current.parent;
  }
  return undefined;
};

/** The outermost function around `node` that `owner` declares directly. */
const nestedFunctionIn = (
  node: AstNode,
  owner: AstNode | undefined,
): AstNode | undefined => {
  let current: unknown = node.parent;
  let nested: AstNode | undefined;
  while (isNode(current) && current !== owner) {
    if (FUNCTIONS.has(current.type)) nested = current;
    current = current.parent;
  }
  return nested;
};

/** Whether `fn` is passed straight to a call, so it runs there. */
const isCallback = (fn: AstNode): boolean =>
  isNode(fn.parent) &&
  fn.parent.type === "CallExpression" &&
  Array.isArray(fn.parent.arguments) &&
  fn.parent.arguments.includes(fn);

/** The loops around `node`, inside its function. */
const loopsOf = (node: unknown): AstNode[] => {
  const loops: AstNode[] = [];
  let current: unknown = isNode(node) ? node.parent : undefined;
  while (isNode(current) && !FUNCTIONS.has(current.type)) {
    if (LOOPS.has(current.type)) loops.push(current);
    current = current.parent;
  }
  return loops;
};

/**
 * Every node under `node`, itself included, short of nested functions: a
 * function value runs later, so what it reads is not the binding's value.
 * A callback passed straight to a call runs now, so `callbacks` walks into
 * those.
 */
const walk = (
  node: unknown,
  visit: (child: AstNode) => void,
  callbacks = false,
): void => {
  if (!isNode(node)) return;
  if (FUNCTIONS.has(node.type) && !(callbacks && isCallback(node))) return;
  visit(node);
  for (const [key, value] of Object.entries(node)) {
    if (key === "parent") continue;
    if (Array.isArray(value))
      for (const item of value) walk(item, visit, callbacks);
    else if (isNode(value)) walk(value, visit, callbacks);
  }
};

const isMember = (node: AstNode, name: string): boolean =>
  node.type === "MemberExpression" && nameOf(node.property) === name;

/**
 * Whether `node` reads block ids now: it calls an id reader, or reads a
 * reader's `.blocks` and an `.id`, in callbacks it passes as well.
 */
const readsBlockIds = (node: unknown): boolean => {
  let reader = false;
  let blocks = false;
  let id = false;
  walk(
    node,
    (child) => {
      if (
        child.type === "CallExpression" &&
        ID_READERS.has(nameOf(child.callee) ?? "")
      )
        reader = true;
      if (isMember(child, "blocks")) blocks = true;
      if (isMember(child, "id")) id = true;
    },
    true,
  );
  return reader || (blocks && id);
};

export default {
  meta: { name: "folio-held-block-ids" },
  rules: {
    "reread-after-resolution": {
      meta: {
        type: "problem",
        messages: {
          staleBlockId:
            "'{{name}}' holds a block id read before accepting or rejecting changes, which can join its paragraph away; read it again through heldBlockIds after the resolution.",
        },
      },
      create(context: RuleContext) {
        const filename = context.filename.replaceAll("\\", "/");
        if (filename.endsWith("/folio-held-block-ids.ts")) return {};
        const memberResolutions: AstNode[] = [];
        const calls: AstNode[] = [];
        const declarators: AstNode[] = [];
        const functions: AstNode[] = [];
        const variableOf = (
          identifier: AstNode,
        ): LexicalVariable | undefined => {
          if (typeof identifier.name !== "string") return undefined;
          let scope: LexicalScope | null =
            context.sourceCode.getScope(identifier);
          while (scope) {
            const variable = scope.set.get(identifier.name);
            if (variable) return variable;
            scope = scope.upper;
          }
          return undefined;
        };
        return {
          CallExpression: (node: unknown) => {
            if (!isNode(node)) return;
            calls.push(node);
            if (
              isNode(node.callee) &&
              node.callee.type === "MemberExpression" &&
              RESOLUTIONS.has(nameOf(node.callee.property) ?? "")
            )
              memberResolutions.push(node);
          },
          FunctionDeclaration: (node: unknown) => {
            if (isNode(node)) functions.push(node);
          },
          VariableDeclarator: (node: unknown) => {
            if (
              isNode(node) &&
              isNode(node.id) &&
              node.id.type === "Identifier"
            )
              declarators.push(node);
          },
          "Program:exit": () => {
            if (memberResolutions.length === 0) return;
            // Functions of the file that resolve, by name, to a fixed point.
            const named = new Map<string, AstNode>();
            for (const fn of functions)
              if (isNode(fn.id) && typeof fn.id.name === "string")
                named.set(fn.id.name, fn);
            for (const declarator of declarators)
              if (
                isNode(declarator.id) &&
                typeof declarator.id.name === "string" &&
                isNode(declarator.init) &&
                FUNCTIONS.has(declarator.init.type)
              )
                named.set(declarator.id.name, declarator.init);
            const resolving = new Set<AstNode>();
            const resolutions = [...memberResolutions];
            for (let grew = true; grew;) {
              grew = false;
              for (const call of resolutions) {
                const fn = functionOf(call);
                if (fn && fn.type !== "Program" && !resolving.has(fn)) {
                  resolving.add(fn);
                  grew = true;
                }
              }
              if (!grew) break;
              for (const call of calls) {
                const target = named.get(nameOf(call.callee) ?? "");
                if (
                  target &&
                  resolving.has(target) &&
                  !resolutions.includes(call)
                )
                  resolutions.push(call);
              }
            }
            const held = new Set<LexicalVariable>();
            // Bindings built from held bindings hold ids too, to a fixed point.
            const heldInit = (init: unknown): boolean => {
              if (readsBlockIds(init)) return true;
              let found = false;
              walk(init, (child) => {
                if (found || child.type !== "Identifier") return;
                const variable = variableOf(child);
                if (variable && held.has(variable)) found = true;
              });
              return found;
            };
            for (let grew = true; grew;) {
              grew = false;
              for (const declarator of declarators) {
                if (!isNode(declarator.id)) continue;
                const variable = variableOf(declarator.id);
                if (!variable || held.has(variable)) continue;
                if (heldInit(declarator.init)) {
                  held.add(variable);
                  grew = true;
                }
              }
            }
            for (const variable of held) {
              const definition = variable.defs.at(0)?.node;
              const writes = [
                definition,
                ...variable.references
                  .filter((reference) => reference.isWrite())
                  .map((reference) => reference.identifier),
              ]
                .filter(isNode)
                .toSorted((a, b) => startOf(a) - startOf(b));
              const owner = functionOf(definition);
              const ownResolutions = resolutions.filter(
                (call) => functionOf(call) === owner,
              );
              for (const reference of variable.references) {
                if (!reference.isRead() || !isNode(reference.identifier))
                  continue;
                let read: AstNode = reference.identifier;
                if (functionOf(read) !== owner) {
                  const nested = nestedFunctionIn(read, owner);
                  if (nested && isCallback(nested)) {
                    // A callback runs where it is passed.
                    read = nested;
                  } else {
                    // Captured: it may run after any resolution.
                    if (
                      ownResolutions.some((call) =>
                        writes.every(
                          (write) =>
                            functionOf(write) !== owner ||
                            startOf(write) < startOf(call),
                        ),
                      )
                    )
                      context.report({
                        node: reference.identifier,
                        messageId: "staleBlockId",
                        data: { name: variable.name },
                      });
                    continue;
                  }
                }
                const lastWrite = writes
                  .filter(
                    (write) =>
                      functionOf(write) === owner &&
                      startOf(write) < startOf(read),
                  )
                  .at(-1);
                const stale = ownResolutions.some((call) => {
                  const after =
                    lastWrite === undefined ||
                    startOf(call) > startOf(lastWrite);
                  // A call's own arguments are read before it resolves.
                  if (endOf(call) <= startOf(read) && after) return true;
                  // A loop runs the call again before the read.
                  return loopsOf(call).some(
                    (loop) =>
                      contains(loop, read) &&
                      writes.every((write) => !contains(loop, write)),
                  );
                });
                if (stale)
                  context.report({
                    node: reference.identifier,
                    messageId: "staleBlockId",
                    data: { name: variable.name },
                  });
              }
            }
          },
        };
      },
    },
  },
};
