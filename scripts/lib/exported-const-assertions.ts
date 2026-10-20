import ts from "typescript";

export type MissingExportType = {
  readonly column: number;
  readonly line: number;
  readonly name: string;
};

const isExported = (node: ts.Node): boolean =>
  ts.canHaveModifiers(node) &&
  (ts.getModifiers(node)?.some(({ kind }) => kind === ts.SyntaxKind.ExportKeyword) ?? false);

const unwrapExpression = (expression: ts.Expression): ts.Expression => {
  if (ts.isParenthesizedExpression(expression) || ts.isSatisfiesExpression(expression)) {
    return unwrapExpression(expression.expression);
  }
  return expression;
};

const isConstAssertedAggregate = (expression: ts.Expression): boolean => {
  const candidate = unwrapExpression(expression);
  if (!ts.isAsExpression(candidate) || candidate.type.getText() !== "const") return false;
  const value = unwrapExpression(candidate.expression);
  return ts.isArrayLiteralExpression(value) || ts.isObjectLiteralExpression(value);
};

/** Find exported const-asserted arrays and objects whose public type is inferred. */
export const exportedConstAssertionsWithoutTypes = (
  sourceText: string,
  fileName = "source.ts",
): readonly MissingExportType[] => {
  const source = ts.createSourceFile(fileName, sourceText, ts.ScriptTarget.Latest, true);
  const missing: MissingExportType[] = [];

  for (const statement of source.statements) {
    if (!ts.isVariableStatement(statement) || !isExported(statement)) continue;
    for (const declaration of statement.declarationList.declarations) {
      if (
        declaration.type ||
        !declaration.initializer ||
        !ts.isIdentifier(declaration.name) ||
        !isConstAssertedAggregate(declaration.initializer)
      ) {
        continue;
      }
      const position = source.getLineAndCharacterOfPosition(declaration.name.getStart(source));
      missing.push({
        column: position.character + 1,
        line: position.line + 1,
        name: declaration.name.text,
      });
    }
  }

  return missing;
};
