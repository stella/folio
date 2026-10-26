import {
  CANONICAL_CAPABILITIES,
  CANONICAL_GAP,
} from "../../packages/core/src/types/canonicalCapabilities";
import { censusUndescribedCanonicalCommands } from "./canonical-command-census";
import {
  createLineIndex,
  type Node,
  nodeText,
  parseSourceModule,
  type Program,
  walkAst,
} from "./oxc-ast";

/** Legacy PM execution still owns these sources when canonical sessions refuse them. */
export const PUBLIC_OPERATION_SOURCE_GAPS = Object.entries(
  CANONICAL_CAPABILITIES,
)
  .filter(
    ([, capability]) =>
      capability.owner === "document-operations" &&
      capability.defaultSessionMutation === "pm-public-operations",
  )
  .map(([gap]) => gap);

export type CanonicalSource = { file: string; source: string };

const gapNames = new Map(Object.entries(CANONICAL_GAP));
const gapIds = new Set(Object.keys(CANONICAL_CAPABILITIES));
const sessionName = /experimentalSession|ExperimentalSession|sessionRef/;

type ScriptPart = { code: string; directive: boolean };

const scriptParts = ({ file, source }: CanonicalSource): ScriptPart[] => {
  if (!file.endsWith(".vue")) return [{ code: source, directive: false }];
  const parts = [
    ...source.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/g),
  ].map((match) => ({ code: match[1] ?? "", directive: false }));
  // Vue directives are JS expressions too; do not leave template-only gates unmeasured.
  for (const match of source.matchAll(
    /(?:v-[\w:-]+|:[\w-]+|@[\w.-]+)="([^"]*)"/g,
  ))
    parts.push({ code: match[1] ?? "", directive: true });
  return parts;
};

/**
 * Vue parts parse as TypeScript. A directive is a handler body or a bare
 * expression such as an object literal; a candidate that is neither is text
 * the attribute pattern matched outside a template (a comment, a string), so
 * it carries no session gate to measure.
 */
const parsePart = (
  file: string,
  { code, directive }: ScriptPart,
): { code: string; program: Program } | undefined => {
  const name = file.endsWith(".vue") ? `${file}.ts` : file;
  try {
    return { code, program: parseSourceModule(name, code).program };
  } catch (error) {
    if (!directive) throw error;
  }
  const expression = `(${code})`;
  try {
    return {
      code: expression,
      program: parseSourceModule(name, expression).program,
    };
  } catch {
    return undefined;
  }
};

const isRefusal = (name: string) =>
  name === "CanonicalSessionRefusalError" || name === "CanonicalSessionError";

/** Optional chains are plain member and call nodes in the shapes compared here. */
const unwrapChain = (node: Node): Node =>
  node.type === "ChainExpression" ? node.expression : node;

const isStaticMember = (node: Node) =>
  node.type === "MemberExpression" && !node.computed;

const isPrefixUnary = (node: Node) =>
  (node.type === "UnaryExpression" &&
    ["!", "-", "+", "~"].includes(node.operator)) ||
  (node.type === "UpdateExpression" && node.prefix);

/** An object literal member written as `name: value`, a shorthand or a method. */
const isObjectLiteralMember = (node: Node, parent: Node | undefined) =>
  node.type === "Property" && parent?.type === "ObjectExpression";

/** The written name of a non-computed property or method key. */
const keyText = (text: string, node: Node): string | undefined => {
  if (
    node.type !== "Property" &&
    node.type !== "MethodDefinition" &&
    node.type !== "TSAbstractMethodDefinition"
  )
    return undefined;
  return node.computed ? undefined : nodeText(text, node.key);
};

const isPropertyAssignment = (node: Node) =>
  node.type === "Property" &&
  node.kind === "init" &&
  !node.method &&
  !node.shorthand;

const isShorthandProperty = (node: Node) =>
  node.type === "Property" && node.shorthand;

/** Enumerate the known PM mutation primitives, wherever they are installed or executed. */
const mutationGap = (
  node: Node,
  parent: Node | undefined,
  text: string,
  file: string,
  imports: Map<string, string>,
): string | undefined => {
  if (node.type === "CallExpression") {
    const name = nodeText(text, node.callee);
    if (name === "fromProseDoc" || imports.get(name)?.endsWith(":fromProseDoc"))
      return CANONICAL_GAP.save;
    if (imports.get(name) === "prosemirror-history:history")
      return CANONICAL_GAP.history;
    if (
      name === "columnResizing" ||
      name === "tableEditing" ||
      imports.get(name) === "prosemirror-tables:columnResizing" ||
      imports.get(name) === "prosemirror-tables:tableEditing"
    )
      return CANONICAL_GAP.tableGeometry;
    if (
      name === "applyFolioAIEditOperations" &&
      file.endsWith("/document-operations.ts")
    )
      return CANONICAL_GAP.documentOperations;
  }
  const method =
    ((node.type === "MethodDefinition" ||
      node.type === "TSAbstractMethodDefinition") &&
      node.kind === "method") ||
    (isObjectLiteralMember(node, parent) &&
      node.type === "Property" &&
      (node.method || isPropertyAssignment(node)));
  if (method) {
    const name = keyText(text, node);
    if (name === "appendTransaction") {
      if (file.endsWith("/suggestionMode.ts"))
        return CANONICAL_GAP.suggestionPlugin;
      if (file.endsWith("/ParaIdAllocatorExtension.ts"))
        return CANONICAL_GAP.paragraphIdentity;
      if (file.endsWith("/ParagraphChangeTrackerExtension.ts"))
        return CANONICAL_GAP.paragraphTracker;
    }
    if (name === "fromBuffer" && file.endsWith("/headless.ts"))
      return CANONICAL_GAP.publicHeadlessSession;
    if (name === "captureReviewerState" && file.endsWith("/headless.ts"))
      return CANONICAL_GAP.aiSnapshots;
  }
  return undefined;
};

export const inspectCanonicalSources = (
  sources: readonly CanonicalSource[],
) => {
  const failures: string[] = [];
  const sites = new Map<string, Set<string>>();
  const branches: Record<string, number> = {};
  const record = (gap: string, file: string) => {
    if (!gapIds.has(gap)) {
      failures.push(`${file}: unknown canonical gap ${gap}`);
      return;
    }
    const files = sites.get(gap) ?? new Set<string>();
    files.add(file);
    sites.set(gap, files);
  };
  for (const source of sources) {
    const { file } = source;
    if (file.endsWith("/canonicalCapabilities.ts")) continue;
    const counts = { branches: 0 };
    for (const scriptPart of scriptParts(source)) {
      const parsed = parsePart(file, scriptPart);
      if (parsed === undefined) continue;
      const { code: part, program } = parsed;
      const text = (node: Node) => nodeText(part, node);
      const lineOf = createLineIndex(part);
      const imports = new Map<string, string>();
      const namespaces = new Set<string>();
      for (const statement of program.body) {
        if (statement.type !== "ImportDeclaration") continue;
        const moduleName = statement.source.value;
        for (const binding of statement.specifiers) {
          if (
            binding.type === "ImportNamespaceSpecifier" &&
            /(?:^|\/)canonicalCapabilities(?:\.ts)?$/.test(moduleName)
          )
            namespaces.add(binding.local.name);
          if (binding.type !== "ImportSpecifier") continue;
          const imported =
            binding.imported.type === "Identifier"
              ? binding.imported.name
              : binding.imported.value;
          imports.set(binding.local.name, `${moduleName}:${imported}`);
          if (
            [
              "usesCanonicalSession",
              "CANONICAL_GAP",
              "CanonicalSessionError",
              "CanonicalSessionRefusalError",
            ].includes(imported) &&
            imported !== binding.local.name
          )
            failures.push(
              `${file}: canonical guard primitives cannot be aliased`,
            );
        }
      }
      // Canonical persistence has one model-only owner. Consumers must capture
      // its snapshot from the controller rather than assembling PM save signals.
      const saveSnapshots = new Set<string>();
      walkAst(program, (node) => {
        if (
          node.type === "VariableDeclarator" &&
          node.id.type === "Identifier" &&
          node.init
        ) {
          const initializer = unwrapChain(node.init);
          if (
            initializer.type === "CallExpression" &&
            /(?:\.|\?\.)captureCanonicalSave$/.test(text(initializer.callee))
          )
            saveSnapshots.add(node.id.name);
        }
      });
      // An identifier token is visited once, even where oxc repeats it as
      // both halves of a shorthand import or property.
      const seenIdentifiers = new Set<number>();
      walkAst(program, (node, ancestors) => {
        const parent = ancestors.at(-1);
        if (
          node.type === "VariableDeclarator" &&
          node.id.type === "Identifier" &&
          node.id.name === "serializeCanonicalSave" &&
          !file.endsWith("/docx/canonicalSave.ts")
        )
          failures.push(
            `${file}: canonical serialization must use its core owner`,
          );
        if (
          file.endsWith("/docx/canonicalSave.ts") &&
          node.type === "Identifier" &&
          !seenIdentifiers.has(node.start) &&
          (parent?.type === "ImportSpecifier" ||
            (parent?.type === "CallExpression" && parent.callee === node)) &&
          [
            "fromProseDoc",
            "getChangedParagraphIds",
            "hasStructuralChanges",
            "hasUntrackedChanges",
          ].includes(node.name)
        )
          failures.push(
            `${file}: canonical serialization cannot derive authority from PM`,
          );
        if (node.type === "Identifier") seenIdentifiers.add(node.start);
        if (
          node.type === "CallExpression" &&
          text(node.callee) === "serializeCanonicalSave"
        ) {
          const argument = node.arguments.at(0);
          const snapshot =
            argument?.type === "ObjectExpression"
              ? argument.properties.find(
                  (property) => keyText(part, property) === "snapshot",
                )
              : undefined;
          let name: string | undefined;
          if (
            snapshot &&
            isShorthandProperty(snapshot) &&
            snapshot.type === "Property" &&
            snapshot.key.type === "Identifier"
          )
            name = snapshot.key.name;
          if (
            snapshot &&
            isPropertyAssignment(snapshot) &&
            snapshot.type === "Property" &&
            snapshot.value.type === "Identifier"
          )
            name = snapshot.value.name;
          if (!name || !saveSnapshots.has(name))
            failures.push(
              `${file}: canonical serialization requires a controller save snapshot`,
            );
        }
      });
      const memberKey = (node: Node) => {
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
      const primitiveName = (node: Node): string | undefined => {
        if (node.type === "Identifier") return node.name;
        if (
          node.type === "MemberExpression" &&
          node.object.type === "Identifier" &&
          namespaces.has(node.object.name)
        )
          return memberKey(node);
        return undefined;
      };
      const gapKey = (node: Node | undefined): string | undefined => {
        if (node?.type !== "MemberExpression") return undefined;
        return primitiveName(node.object) === "CANONICAL_GAP"
          ? memberKey(node)
          : undefined;
      };
      const lines = part.split("\n");
      const aliases = new Set<string>();
      const safeSelectors = new Set<string>();
      const gapRelays = new Set<string>();
      const selectorsIn = (node: Node): boolean => {
        if (sessionName.test(text(node))) return true;
        let found = false;
        walkAst(node, (child) => {
          if (
            (child.type === "Identifier" || child.type === "JSXIdentifier") &&
            aliases.has(child.name)
          )
            found = true;
        });
        return found;
      };
      walkAst(program, (node) => {
        if (
          node.type !== "VariableDeclarator" ||
          node.id.type !== "Identifier" ||
          !node.init
        )
          return;
        const initial = node.init;
        if (
          initial.type === "ArrowFunctionExpression" &&
          initial.params.some((parameter) => {
            const binding =
              parameter.type === "AssignmentPattern"
                ? parameter.left
                : parameter;
            const type =
              binding.type === "TSParameterProperty"
                ? undefined
                : binding.typeAnnotation?.typeAnnotation;
            if (type?.type !== "TSTypeReference") return false;
            const typeName = text(type.typeName);
            return (
              typeName === "CanonicalGap" ||
              imports.get(typeName)?.endsWith(":CanonicalGap")
            );
          })
        )
          gapRelays.add(node.id.name);
        const value = unwrapChain(initial);
        if (
          (value.type === "Identifier" ||
            isStaticMember(value) ||
            (value.type === "CallExpression" &&
              /toValue|[gG]etExperimentalSession/.test(text(value.callee)))) &&
          selectorsIn(initial)
        )
          aliases.add(node.id.name);
        if (
          initial.type === "ArrowFunctionExpression" &&
          initial.expression &&
          unwrapChain(initial.body).type === "CallExpression"
        ) {
          const body = unwrapChain(initial.body);
          if (
            body.type === "CallExpression" &&
            primitiveName(body.callee) === "usesCanonicalSession"
          )
            safeSelectors.add(node.id.name);
        }
      });
      const markers = [...part.matchAll(/\/\/ canonical-gap: ([\w.-]+)/g)];
      const usedMarkers = new Set<number>();
      const markerFor = (node: Node, gap: string) => {
        const { line } = lineOf(node.start);
        const marker = markers.find((match) => {
          const markerLine = lineOf(match.index).line;
          const between = lines.slice(markerLine + 1, line);
          return (
            match[1] === gap &&
            markerLine < line &&
            between.every((entry) => /^\s*\/\/ canonical-gap:/.test(entry))
          );
        });
        if (!marker)
          failures.push(
            `${file}:${line + 1}: ${gap} source needs its ledger marker`,
          );
        else {
          usedMarkers.add(marker.index);
          record(gap, file);
        }
      };
      walkAst(program, (node, ancestors) => {
        if (
          node.type === "BinaryExpression" &&
          ["===", "!==", "==", "!="].includes(node.operator) &&
          (selectorsIn(node) ||
            (!file.startsWith("packages/core/") &&
              [node.left, node.right].some(
                (operand) =>
                  text(operand) === '"canonical"' ||
                  text(operand) === "'canonical'",
              )))
        ) {
          // The operands themselves identify a selector, rather than nested helper calls.
          const operands = [node.left, node.right];
          if (
            operands.some((operand) => {
              const call = unwrapChain(operand);
              return (
                call.type !== "CallExpression" ||
                primitiveName(call.callee) !== "usesCanonicalSession"
              );
            }) &&
            operands.some(
              (operand) =>
                text(operand) === '"canonical"' ||
                text(operand) === "'canonical'" ||
                text(operand) === "undefined",
            )
          ) {
            counts.branches++;
            failures.push(
              `${file}: raw session comparison must use usesCanonicalSession with a ledger id`,
            );
          }
        }
        const key = gapKey(node);
        if (key !== undefined) {
          const gap = gapNames.get(key);
          if (gap === undefined)
            failures.push(`${file}: unknown canonical gap key ${key}`);
          else record(gap, file);
        }
        const conditionOf = () => {
          if (node.type === "IfStatement") return node.test;
          if (node.type === "ConditionalExpression") return node.test;
          return undefined;
        };
        const condition = conditionOf();
        const conditionValue = condition && unwrapChain(condition);
        if (
          condition &&
          conditionValue &&
          (conditionValue.type === "Identifier" ||
            isStaticMember(conditionValue) ||
            isPrefixUnary(conditionValue)) &&
          selectorsIn(condition) &&
          !text(condition).includes("usesCanonicalSession") &&
          ![...safeSelectors].some((selector) =>
            text(condition).includes(selector),
          )
        )
          failures.push(`${file}: bare session gate must name a ledger id`);
        if (
          node.type === "CallExpression" &&
          text(node.callee) === "refuseCanonicalModelEdit"
        ) {
          const gap = node.arguments.at(0);
          if (gapKey(gap && unwrapChain(gap)) === undefined)
            failures.push(`${file}: model refusal must name a ledger id`);
        }
        if (
          node.type === "CallExpression" &&
          primitiveName(node.callee) === "usesCanonicalSession"
        ) {
          counts.branches++;
          if (node.arguments.length !== 2)
            failures.push(`${file}: session branch needs a ledger id`);
          const gap = node.arguments.at(1);
          if (gapKey(gap && unwrapChain(gap)) === undefined)
            failures.push(
              `${file}: session branch must name a constant ledger id`,
            );
        }
        if (
          node.type === "CallExpression" &&
          text(node.callee) === "handleSessionRefusal" &&
          node.arguments.length !== 2
        )
          failures.push(
            `${file}: session refusal handler must receive its gap id`,
          );
        if (
          node.type === "CallExpression" &&
          gapRelays.has(text(node.callee))
        ) {
          const messageArgument = node.arguments.at(0);
          const gapArgument = node.arguments.at(1);
          const message = messageArgument && unwrapChain(messageArgument);
          const gap = gapArgument && unwrapChain(gapArgument);
          if (
            message?.type === "MemberExpression" &&
            memberKey(message) === "message" &&
            memberKey(message.object) === "error"
          ) {
            if (
              !gap ||
              gap.type !== "MemberExpression" ||
              memberKey(gap) !== "gap" ||
              text(gap.object) !== text(message.object)
            )
              failures.push(
                `${file}: canonical error relay must forward the original gap id`,
              );
          }
        }
        if (
          node.type === "NewExpression" &&
          isRefusal(primitiveName(node.callee) ?? "")
        ) {
          const payload = node.arguments.at(0);
          if (
            payload?.type !== "ObjectExpression" ||
            !payload.properties.some(
              (property) =>
                (isPropertyAssignment(property) ||
                  isShorthandProperty(property)) &&
                keyText(part, property) === "gap",
            )
          )
            failures.push(`${file}: canonical refusal must carry its gap id`);
        }
        const mutation = mutationGap(
          node,
          ancestors.at(-1),
          part,
          file,
          imports,
        );
        if (mutation) markerFor(node, mutation);
        if (mutation === CANONICAL_GAP.documentOperations) {
          for (const gap of PUBLIC_OPERATION_SOURCE_GAPS) markerFor(node, gap);
        }
        if (
          node.type === "IfStatement" &&
          text(node.test) === "intents === undefined"
        )
          markerFor(node, CANONICAL_GAP.commands);
      });
      for (const marker of markers)
        if (!usedMarkers.has(marker.index))
          failures.push(`${file}: orphan canonical gap marker ${marker[1]}`);
    }
    if (counts.branches > 0) branches[file] = counts.branches;
  }
  for (const gap of gapIds)
    if (!sites.has(gap))
      failures.push(`Ledger entry ${gap} has no source site`);
  const remainingCommands = new Map<string, string[]>();
  for (const [gap, capability] of Object.entries(CANONICAL_CAPABILITIES)) {
    if (!("remainingCommands" in capability)) continue;
    remainingCommands.set(
      gap,
      censusUndescribedCanonicalCommands(
        sources.filter(
          ({ file }) =>
            file.startsWith("packages/core/src/prosemirror/") &&
            file.endsWith(".ts"),
        ),
      ),
    );
  }
  return { failures, sites, branches, remainingCommands };
};

export const checkCanonicalBaseline = (
  current: Record<string, number>,
  baseline: Record<string, number>,
) => {
  const failures: string[] = [];
  for (const file of new Set([
    ...Object.keys(current),
    ...Object.keys(baseline),
  ])) {
    const actual = current[file] ?? 0;
    const allowed = baseline[file] ?? 0;
    if (actual > allowed)
      failures.push(
        `${file}: session branches increased from ${allowed} to ${actual}`,
      );
    if (actual < allowed)
      failures.push(
        `${file}: record the session branch decrease from ${allowed} to ${actual}`,
      );
  }
  return failures;
};

type CanonicalCutoverDocsOptions = {
  sites: Map<string, Set<string>>;
  remainingCommands: Map<string, string[]>;
};

export const canonicalCutoverDocs = ({
  sites,
  remainingCommands,
}: CanonicalCutoverDocsOptions) => {
  const headers = [
    "Id",
    "Owner",
    "Kind",
    "Adapters",
    "Remaining work",
    "Source files",
  ];
  const rows = Object.entries(CANONICAL_CAPABILITIES).map(
    ([id, capability]) => [
      id,
      capability.owner,
      capability.kind,
      capability.adapters.join(", "),
      capability.summary,
      [...(sites.get(id) ?? [])]
        .sort()
        .map((file) => `\`${file}\``)
        .join("<br>"),
    ],
  );
  const widths = headers.map((header, column) =>
    Math.max(
      3,
      header.length,
      ...rows.map((row) => row.at(column)?.length ?? 0),
    ),
  );
  const tableRow = (cells: string[]) =>
    `| ${cells.map((cell, column) => cell.padEnd(widths.at(column) ?? cell.length)).join(" | ")} |`;
  return [
    "# Canonical session cutover",
    "",
    "Generated by `bun scripts/check-canonical-cutover.ts --write` from the core capability ledger and source sites. Do not edit by hand.",
    "",
    "Routing entries include supported canonical behavior with a remaining legacy branch. Refusal entries identify gated capabilities; mutation-source entries identify PM producers still to retire. Ledger entries identify remaining work and are deleted when their sites are retired; an entry does not imply that every operation in its area is unsupported.",
    "",
    "Session branch counts are held per file in `scripts/canonical-cutover-baseline.json`. Removing a branch requires recording the decrease; CI rejects increases against both the committed baseline and its merge-base version.",
    "",
    tableRow(headers),
    tableRow(widths.map((width) => "-".repeat(width))),
    ...rows.map(tableRow),
    "",
    "## Remaining command registrations",
    "",
    "Derived from the extension command registry. Entries lack a structural proof that every factory return path attaches a canonical descriptor; this conservative census is not a runtime support policy. CI requires the committed list to match the census and rejects new entries against the PR-base source census.",
    "",
    ...[...remainingCommands].flatMap(([gap, commands]) =>
      [`### ${gap} (${commands.length})`, ""].concat(
        commands.map((command) => `- \`${command}\``),
        [""],
      ),
    ),
  ].join("\n");
};
