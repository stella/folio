import ts from "typescript";
import {
  CANONICAL_CAPABILITIES,
  CANONICAL_GAP,
} from "../../packages/core/src/types/canonicalCapabilities";

export type CanonicalSource = { file: string; source: string };

const gapNames = new Map(Object.entries(CANONICAL_GAP));
const gapIds = new Set(Object.keys(CANONICAL_CAPABILITIES));
const sessionName = /experimentalSession|ExperimentalSession|sessionRef/;

const scriptParts = ({ file, source }: CanonicalSource) => {
  if (!file.endsWith(".vue")) return [source];
  const parts = [...source.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/g)].map(
    (match) => match[1] ?? "",
  );
  // Vue directives are JS expressions too; do not leave template-only gates unmeasured.
  for (const match of source.matchAll(/(?:v-[\w:-]+|:[\w-]+|@[\w.-]+)="([^"]*)"/g))
    parts.push(match[1] ?? "");
  return parts;
};

const isRefusal = (name: string) =>
  name === "CanonicalSessionRefusalError" || name === "CanonicalSessionError";

/** Enumerate the known PM mutation primitives, wherever they are installed or executed. */
const mutationGap = (
  node: ts.Node,
  file: string,
  imports: Map<string, string>,
): string | undefined => {
  if (ts.isCallExpression(node)) {
    const name = node.expression.getText();
    if (name === "fromProseDoc" || imports.get(name)?.endsWith(":fromProseDoc"))
      return CANONICAL_GAP.save;
    if (imports.get(name) === "prosemirror-history:history") return CANONICAL_GAP.history;
    if (
      name === "columnResizing" ||
      name === "tableEditing" ||
      imports.get(name) === "prosemirror-tables:columnResizing" ||
      imports.get(name) === "prosemirror-tables:tableEditing"
    )
      return CANONICAL_GAP.tableGeometry;
    if (name === "applyFolioAIEditOperations" && file.endsWith("/document-operations.ts"))
      return CANONICAL_GAP.documentOperations;
  }
  if (ts.isMethodDeclaration(node) || ts.isPropertyAssignment(node)) {
    const name = node.name.getText();
    if (name === "appendTransaction") {
      if (file.endsWith("/suggestionMode.ts")) return CANONICAL_GAP.suggestionPlugin;
      if (file.endsWith("/ParaIdAllocatorExtension.ts")) return CANONICAL_GAP.paragraphIdentity;
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

export const inspectCanonicalSources = (sources: readonly CanonicalSource[]) => {
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
    for (const part of scriptParts(source)) {
      const tree = ts.createSourceFile(
        file,
        part,
        ts.ScriptTarget.Latest,
        true,
        file.endsWith(".tsx") ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
      );
      const imports = new Map<string, string>();
      for (const statement of tree.statements) {
        if (!ts.isImportDeclaration(statement) || !ts.isStringLiteral(statement.moduleSpecifier))
          continue;
        const bindings = statement.importClause?.namedBindings;
        if (bindings && ts.isNamedImports(bindings))
          for (const binding of bindings.elements) {
            const imported = binding.propertyName?.text ?? binding.name.text;
            imports.set(binding.name.text, `${statement.moduleSpecifier.text}:${imported}`);
            if (
              [
                "usesCanonicalSession",
                "CANONICAL_GAP",
                "CanonicalSessionError",
                "CanonicalSessionRefusalError",
              ].includes(imported) &&
              imported !== binding.name.text
            )
              failures.push(`${file}: canonical guard primitives cannot be aliased`);
          }
      }
      const lines = part.split("\n");
      const aliases = new Set<string>();
      const safeSelectors = new Set<string>();
      const selectorsIn = (node: ts.Node): boolean => {
        if (sessionName.test(node.getText())) return true;
        let found = false;
        const walk = (child: ts.Node) => {
          if (ts.isIdentifier(child) && aliases.has(child.text)) found = true;
          ts.forEachChild(child, walk);
        };
        walk(node);
        return found;
      };
      const collectAliases = (node: ts.Node) => {
        if (ts.isVariableDeclaration(node) && ts.isIdentifier(node.name) && node.initializer) {
          const initial = node.initializer;
          if (
            (ts.isIdentifier(initial) ||
              ts.isPropertyAccessExpression(initial) ||
              (ts.isCallExpression(initial) &&
                /toValue|[gG]etExperimentalSession/.test(initial.expression.getText()))) &&
            selectorsIn(initial)
          )
            aliases.add(node.name.text);
          if (
            ts.isArrowFunction(initial) &&
            initial.body.getText().startsWith("usesCanonicalSession(")
          )
            safeSelectors.add(node.name.text);
        }
        ts.forEachChild(node, collectAliases);
      };
      collectAliases(tree);
      const markers = [...part.matchAll(/\/\/ canonical-gap: ([\w.-]+)/g)];
      const usedMarkers = new Set<number>();
      const markerFor = (node: ts.Node, gap: string) => {
        const line = tree.getLineAndCharacterOfPosition(node.getStart()).line;
        const marker = markers.find((match) => {
          const markerLine = tree.getLineAndCharacterOfPosition(match.index).line;
          const between = lines.slice(markerLine + 1, line);
          return (
            match[1] === gap &&
            markerLine < line &&
            between.every((text) => /^\s*\/\/ canonical-gap:/.test(text))
          );
        });
        if (!marker) failures.push(`${file}:${line + 1}: ${gap} source needs its ledger marker`);
        else {
          usedMarkers.add(marker.index);
          record(gap, file);
        }
      };
      const visit = (node: ts.Node) => {
        if (
          ts.isBinaryExpression(node) &&
          (selectorsIn(node) ||
            (!file.startsWith("packages/core/") &&
              [node.left, node.right].some(
                (operand) =>
                  operand.getText() === '"canonical"' || operand.getText() === "'canonical'",
              ))) &&
          [
            ts.SyntaxKind.EqualsEqualsEqualsToken,
            ts.SyntaxKind.ExclamationEqualsEqualsToken,
            ts.SyntaxKind.EqualsEqualsToken,
            ts.SyntaxKind.ExclamationEqualsToken,
          ].includes(node.operatorToken.kind)
        ) {
          // The operands themselves identify a selector, rather than nested helper calls.
          const operands = [node.left, node.right];
          if (
            operands.some(
              (operand) =>
                !ts.isCallExpression(operand) ||
                operand.expression.getText() !== "usesCanonicalSession",
            ) &&
            operands.some(
              (operand) =>
                operand.getText() === '"canonical"' ||
                operand.getText() === "'canonical'" ||
                operand.getText() === "undefined",
            )
          ) {
            counts.branches++;
            failures.push(
              `${file}: raw session comparison must use usesCanonicalSession with a ledger id`,
            );
          }
        }
        if (ts.isPropertyAccessExpression(node) && node.expression.getText() === "CANONICAL_GAP") {
          const gap = gapNames.get(node.name.text);
          if (gap === undefined)
            failures.push(`${file}: unknown canonical gap key ${node.name.text}`);
          else record(gap, file);
        }
        const conditionOf = () => {
          if (ts.isIfStatement(node)) return node.expression;
          if (ts.isConditionalExpression(node)) return node.condition;
          return undefined;
        };
        const condition = conditionOf();
        if (
          condition &&
          (ts.isIdentifier(condition) ||
            ts.isPropertyAccessExpression(condition) ||
            ts.isPrefixUnaryExpression(condition)) &&
          selectorsIn(condition) &&
          !condition.getText().includes("usesCanonicalSession") &&
          ![...safeSelectors].some((selector) => condition.getText().includes(selector))
        )
          failures.push(`${file}: bare session gate must name a ledger id`);
        if (ts.isCallExpression(node) && node.expression.getText() === "refuseCanonicalModelEdit") {
          const gap = node.arguments.at(0);
          if (
            !gap ||
            !ts.isPropertyAccessExpression(gap) ||
            gap.expression.getText() !== "CANONICAL_GAP"
          )
            failures.push(`${file}: model refusal must name a ledger id`);
        }
        if (ts.isCallExpression(node) && node.expression.getText() === "usesCanonicalSession") {
          counts.branches++;
          if (node.arguments.length !== 2)
            failures.push(`${file}: session branch needs a ledger id`);
          const gap = node.arguments.at(1);
          if (
            !gap ||
            !(ts.isPropertyAccessExpression(gap) && gap.expression.getText() === "CANONICAL_GAP")
          )
            failures.push(`${file}: session branch must name a constant ledger id`);
        }
        if (ts.isNewExpression(node) && isRefusal(node.expression.getText())) {
          const payload = node.arguments?.at(0);
          if (
            !payload ||
            !ts.isObjectLiteralExpression(payload) ||
            !payload.properties.some(
              (property) =>
                (ts.isPropertyAssignment(property) || ts.isShorthandPropertyAssignment(property)) &&
                property.name.getText() === "gap",
            )
          )
            failures.push(`${file}: canonical refusal must carry its gap id`);
        }
        const mutation = mutationGap(node, file, imports);
        if (mutation) markerFor(node, mutation);
        if (mutation === CANONICAL_GAP.documentOperations) {
          for (const gap of [
            CANONICAL_GAP.publicComments,
            CANONICAL_GAP.publicSuggestedMode,
            CANONICAL_GAP.publicTableProjection,
            CANONICAL_GAP.publicUnsupportedInline,
          ])
            markerFor(node, gap);
        }
        if (ts.isIfStatement(node) && node.expression.getText() === "intents === undefined")
          markerFor(node, CANONICAL_GAP.commands);
        ts.forEachChild(node, visit);
      };
      visit(tree);
      for (const marker of markers)
        if (!usedMarkers.has(marker.index))
          failures.push(`${file}: orphan canonical gap marker ${marker[1]}`);
    }
    if (counts.branches > 0) branches[file] = counts.branches;
  }
  for (const gap of gapIds)
    if (!sites.has(gap)) failures.push(`Ledger entry ${gap} has no source site`);
  return { failures, sites, branches };
};

export const checkCanonicalBaseline = (
  current: Record<string, number>,
  baseline: Record<string, number>,
) => {
  const failures: string[] = [];
  for (const file of new Set([...Object.keys(current), ...Object.keys(baseline)])) {
    const actual = current[file] ?? 0;
    const allowed = baseline[file] ?? 0;
    if (actual > allowed)
      failures.push(`${file}: session branches increased from ${allowed} to ${actual}`);
    if (actual < allowed)
      failures.push(`${file}: record the session branch decrease from ${allowed} to ${actual}`);
  }
  return failures;
};

export const canonicalCutoverDocs = (sites: Map<string, Set<string>>) => {
  const headers = ["Id", "Owner", "Kind", "Adapters", "Remaining work", "Source files"];
  const rows = Object.entries(CANONICAL_CAPABILITIES).map(([id, capability]) => [
    id,
    capability.owner,
    capability.kind,
    capability.adapters.join(", "),
    capability.summary,
    [...(sites.get(id) ?? [])]
      .sort()
      .map((file) => `\`${file}\``)
      .join("<br>"),
  ]);
  const widths = headers.map((header, column) =>
    Math.max(3, header.length, ...rows.map((row) => row.at(column)?.length ?? 0)),
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
  ].join("\n");
};
