import type { Expression } from "oxc-parser";

import { createLineIndex, nodeText, parseSourceModule } from "./oxc-ast";

export type MissingExportType = {
  readonly column: number;
  readonly line: number;
  readonly name: string;
};

const unwrapExpression = (expression: Expression): Expression => {
  if (
    expression.type === "ParenthesizedExpression" ||
    expression.type === "TSSatisfiesExpression"
  ) {
    return unwrapExpression(expression.expression);
  }
  return expression;
};

const isConstAssertedAggregate = (
  text: string,
  expression: Expression,
): boolean => {
  const candidate = unwrapExpression(expression);
  if (
    candidate.type !== "TSAsExpression" ||
    nodeText(text, candidate.typeAnnotation) !== "const"
  )
    return false;
  const value = unwrapExpression(candidate.expression);
  return value.type === "ArrayExpression" || value.type === "ObjectExpression";
};

/** Find exported const-asserted arrays and objects whose public type is inferred. */
export const exportedConstAssertionsWithoutTypes = (
  sourceText: string,
  fileName = "source.ts",
): readonly MissingExportType[] => {
  const { program } = parseSourceModule(fileName, sourceText);
  const lineOf = createLineIndex(sourceText);
  const missing: MissingExportType[] = [];

  for (const statement of program.body) {
    if (
      statement.type !== "ExportNamedDeclaration" ||
      statement.declaration?.type !== "VariableDeclaration"
    )
      continue;
    for (const declaration of statement.declaration.declarations) {
      if (
        declaration.id.type !== "Identifier" ||
        declaration.id.typeAnnotation ||
        !declaration.init ||
        !isConstAssertedAggregate(sourceText, declaration.init)
      ) {
        continue;
      }
      const position = lineOf(declaration.id.start);
      missing.push({
        column: position.column + 1,
        line: position.line + 1,
        name: declaration.id.name,
      });
    }
  }

  return missing;
};
