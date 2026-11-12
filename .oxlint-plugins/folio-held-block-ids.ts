/**
 * A block id held across accepting or rejecting changes is read again first.
 *
 * A block's id is its paragraph's `w14:paraId`. Resolving a change that joins
 * a paragraph into the next one leaves the paragraph whose mark stays, the
 * next one, under its own id (the reference rule), so an id read before the
 * resolution can name a paragraph that is gone. A caller reads it again after
 * the resolution through `heldBlockIds`, instead of reusing the old binding.
 *
 * A binding holds a block id read from a reader (the identifier at the root
 * of the read) when its value reads that reader's `.blocks` and an `.id`,
 * calls `idOf` or `heldBlockIds` with it, or is built from another such
 * binding. A resolution of a reader is a call of `acceptAll`, `rejectAll`,
 * `acceptChange`, `rejectChange`, `acceptSuggestion` or `rejectSuggestion` on
 * it, or a call passing it to a function of the file that makes one (a call
 * passing no reader resolves any). A read of a held binding is reported when
 * a resolution of its reader, in the function declaring the binding, can run
 * between the binding's last write and the read: earlier in source order and
 * not cut off by a `return` or `throw` after it, or in a loop the read shares
 * with the resolution and no write of the binding stands in. A read inside a
 * callback counts where the callback is passed, and one inside a function a
 * binding names counts where that binding is read. Passing the binding to an
 * id reader is the read again, not a use.
 */

type AstNode = Record<string, unknown> & {
  type: string;
  parent?: unknown;
  range?: unknown;
};
type LexicalReference = {
  identifier: unknown;
  isRead: () => boolean;
  isWrite: () => boolean;
};
type LexicalVariable = {
  name: string;
  defs: readonly { node: unknown }[];
  references: readonly LexicalReference[];
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
/** Calls that read block ids from the reader they are passed first. */
const ID_READERS: ReadonlySet<string> = new Set(["idOf", "heldBlockIds"]);
/** A resolution passing no reader: it may resolve any. */
const ANY_READER = "";

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

const isMember = (node: AstNode, name: string): boolean =>
  node.type === "MemberExpression" && nameOf(node.property) === name;

const rangeOf = (node: unknown): [number, number] => {
  if (!isNode(node) || !Array.isArray(node.range)) return [-1, -1];
  const [start, end]: unknown[] = node.range;
  return [
    typeof start === "number" ? start : -1,
    typeof end === "number" ? end : -1,
  ];
};
const startOf = (node: unknown): number => rangeOf(node)[0];
const endOf = (node: unknown): number => rangeOf(node)[1];

const contains = (outer: AstNode, inner: unknown): boolean => {
  let current: unknown = inner;
  while (isNode(current)) {
    if (current === outer) return true;
    current = current.parent;
  }
  return false;
};

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
 * The block a `return` or `throw` after `call` leaves, inside its function:
 * code outside it never runs after the call.
 */
const exitedBlockOf = (call: AstNode): AstNode | undefined => {
  let child: unknown = call;
  let current: unknown = call.parent;
  while (isNode(current) && !FUNCTIONS.has(current.type)) {
    if (current.type === "BlockStatement" && Array.isArray(current.body)) {
      const at = current.body.indexOf(child);
      if (
        current.body
          .slice(at + 1)
          .some(
            (statement) =>
              isNode(statement) &&
              (statement.type === "ReturnStatement" ||
                statement.type === "ThrowStatement"),
          )
      )
        return current;
    }
    child = current;
    current = current.parent;
  }
  return undefined;
};

/**
 * Every node under `node`, itself included, short of nested functions: a
 * function value runs later, so what it reads is not the binding's value.
 * With `callbacks`, a function passed straight to a call is walked too: it
 * runs now.
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

/** The identifier at the root of a member or call chain. */
const rootOf = (node: unknown): string | undefined => {
  let current: unknown = node;
  while (isNode(current)) {
    if (current.type === "Identifier") return nameOf(current);
    if (current.type === "ThisExpression") return "this";
    if (current.type === "MemberExpression") current = current.object;
    else if (current.type === "CallExpression") current = current.callee;
    else if (
      current.type === "ChainExpression" ||
      current.type === "TSNonNullExpression" ||
      current.type === "ParenthesizedExpression"
    )
      current = current.expression;
    else return undefined;
  }
  return undefined;
};

/** The readers `node` reads block ids from now, if it reads any. */
const blockIdReadersOf = (node: unknown): Set<string> => {
  const readers = new Set<string>();
  const blockReaders = new Set<string>();
  let id = false;
  walk(
    node,
    (child) => {
      if (
        child.type === "CallExpression" &&
        ID_READERS.has(nameOf(child.callee) ?? "") &&
        Array.isArray(child.arguments)
      )
        readers.add(rootOf(child.arguments.at(0)) ?? ANY_READER);
      if (isMember(child, "blocks"))
        blockReaders.add(rootOf(child.object) ?? ANY_READER);
      if (isMember(child, "id")) id = true;
    },
    true,
  );
  if (id) for (const reader of blockReaders) readers.add(reader);
  return readers;
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
            // Each resolution, with the readers it resolves.
            const resolutions = new Map<AstNode, Set<string>>();
            for (const call of memberResolutions) {
              const callee = call.callee;
              if (isNode(callee))
                resolutions.set(
                  call,
                  new Set([rootOf(callee.object) ?? ANY_READER]),
                );
            }
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
            for (let grew = true; grew;) {
              grew = false;
              for (const call of resolutions.keys()) {
                const fn = functionOf(call);
                if (fn && fn.type !== "Program" && !resolving.has(fn)) {
                  resolving.add(fn);
                  grew = true;
                }
              }
              for (const call of calls) {
                const target = named.get(nameOf(call.callee) ?? "");
                if (!target || !resolving.has(target) || resolutions.has(call))
                  continue;
                const passed = Array.isArray(call.arguments)
                  ? call.arguments.flatMap((argument) => {
                      const root = rootOf(argument);
                      return root === undefined ? [] : [root];
                    })
                  : [];
                resolutions.set(
                  call,
                  new Set(passed.length === 0 ? [ANY_READER] : passed),
                );
                grew = true;
              }
            }
            // Held bindings and their readers, to a fixed point.
            const held = new Map<LexicalVariable, Set<string>>();
            for (let grew = true; grew;) {
              grew = false;
              for (const declarator of declarators) {
                if (!isNode(declarator.id)) continue;
                const variable = variableOf(declarator.id);
                if (!variable || held.has(variable)) continue;
                const readers = blockIdReadersOf(declarator.init);
                walk(
                  declarator.init,
                  (child) => {
                    if (child.type !== "Identifier") return;
                    const source = variableOf(child);
                    for (const reader of (source && held.get(source)) ?? [])
                      readers.add(reader);
                  },
                  true,
                );
                if (readers.size > 0) {
                  held.set(variable, readers);
                  grew = true;
                }
              }
            }
            for (const [variable, readers] of held) {
              const definition = variable.defs.at(0)?.node;
              const owner = functionOf(definition);
              const writes = [
                definition,
                ...variable.references
                  .filter((reference) => reference.isWrite())
                  .map((reference) => reference.identifier),
              ]
                .filter(isNode)
                .filter((write) => functionOf(write) === owner)
                .toSorted((a, b) => startOf(a) - startOf(b));
              const ownResolutions = [...resolutions].flatMap(
                ([call, resolved]) =>
                  functionOf(call) === owner &&
                  (resolved.has(ANY_READER) ||
                    readers.has(ANY_READER) ||
                    [...resolved].some((reader) => readers.has(reader)))
                    ? [call]
                    : [],
              );
              if (ownResolutions.length === 0) continue;
              /** Whether a resolution can run between the last write and `site`. */
              const staleAt = (site: AstNode): boolean => {
                const lastWrite = writes
                  .filter((write) => startOf(write) < startOf(site))
                  .at(-1);
                return ownResolutions.some((call) => {
                  const exited = exitedBlockOf(call);
                  const reaches = !exited || contains(exited, site);
                  // A call's own arguments are read before it resolves.
                  if (
                    reaches &&
                    endOf(call) <= startOf(site) &&
                    (lastWrite === undefined ||
                      startOf(call) > startOf(lastWrite))
                  )
                    return true;
                  // A loop runs the call again before the read.
                  return loopsOf(call).some(
                    (loop) =>
                      contains(loop, site) &&
                      writes.every((write) => !contains(loop, write)),
                  );
                });
              };
              /** The places a read inside `fn` runs, in `owner`. */
              const sitesOf = (fn: AstNode): AstNode[] | undefined => {
                if (isCallback(fn)) return [fn];
                if (
                  !isNode(fn.parent) ||
                  fn.parent.type !== "VariableDeclarator" ||
                  !isNode(fn.parent.id)
                )
                  return undefined;
                const binding = variableOf(fn.parent.id);
                if (!binding) return undefined;
                const sites: AstNode[] = [];
                for (const reference of binding.references) {
                  if (!reference.isRead() || !isNode(reference.identifier))
                    continue;
                  const site = reference.identifier;
                  if (functionOf(site) === owner) sites.push(site);
                  else {
                    const nested = nestedFunctionIn(site, owner);
                    if (!nested || !isCallback(nested)) return undefined;
                    sites.push(nested);
                  }
                }
                return sites;
              };
              for (const reference of variable.references) {
                if (!reference.isRead() || !isNode(reference.identifier))
                  continue;
                const read = reference.identifier;
                // Passing it to an id reader is the read again.
                if (
                  isNode(read.parent) &&
                  read.parent.type === "CallExpression" &&
                  ID_READERS.has(nameOf(read.parent.callee) ?? "")
                )
                  continue;
                let sites: AstNode[] | undefined = [read];
                if (functionOf(read) !== owner) {
                  const nested = nestedFunctionIn(read, owner);
                  sites = nested ? sitesOf(nested) : undefined;
                }
                // A read that may run anywhere is stale unless a write
                // follows every resolution.
                const stale =
                  sites === undefined
                    ? ownResolutions.some((call) =>
                        writes.every((write) => startOf(write) < startOf(call)),
                      )
                    : sites.some(staleAt);
                if (stale)
                  context.report({
                    node: read,
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
