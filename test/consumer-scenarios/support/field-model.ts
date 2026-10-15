import assert from "node:assert/strict";
import { isDeepStrictEqual } from "node:util";

/** Field syntax is a carrier; instruction, cached result and review data are meaning. */
export const canonicalFieldModel = (field: Record<string, unknown>): Record<string, unknown> => {
  if (field.type !== "simpleField" && field.type !== "complexField") return field;
  assert.ok(typeof field.instruction === "string", "A field requires an instruction");
  const result = field.type === "simpleField" ? field.content : field.fieldResult;
  assert.ok(Array.isArray(result), "A field requires result content");
  const first = result.at(0);
  const resultFormatting =
    first && typeof first === "object" && "formatting" in first ? first.formatting : undefined;
  const code = field.fieldCode;
  let instruction = "";
  const plainInstruction =
    Array.isArray(code) &&
    code.every((run: unknown) => {
      if (
        run === null ||
        typeof run !== "object" ||
        !("type" in run) ||
        run.type !== "run" ||
        !("content" in run) ||
        !Array.isArray(run.content)
      )
        return false;
      if (Object.keys(run).some((key) => !["type", "content", "formatting"].includes(key)))
        return false;
      if ("formatting" in run && !isDeepStrictEqual(run.formatting, resultFormatting)) return false;
      return run.content.every((item: unknown) => {
        if (
          item === null ||
          typeof item !== "object" ||
          !("type" in item) ||
          item.type !== "instrText" ||
          !("text" in item) ||
          typeof item.text !== "string" ||
          Object.keys(item).some((key) => key !== "type" && key !== "text")
        )
          return false;
        instruction += item.text;
        return true;
      });
    });
  const entries = Object.entries(field).filter(
    ([key]) =>
      !["type", "content", "fieldResult", "fieldCode", "instruction"].includes(key) &&
      !(key === "formatting" && isDeepStrictEqual(field.formatting, resultFormatting)),
  );
  return Object.fromEntries([
    ...entries,
    ["type", "field"],
    ["instruction", field.instruction.trim()],
    ["result", result],
    ...(code !== undefined && (!plainInstruction || instruction.trim() !== field.instruction.trim())
      ? [["instructionParts", code]]
      : []),
  ]);
};
