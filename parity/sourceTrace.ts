#!/usr/bin/env bun
/**
 * Static source-pipeline trace for one text query.
 *
 * Parses a DOCX and searches the parsed document model, ProseMirror document,
 * and FlowBlock conversion output. It does not render Word or Folio. The goal
 * is a deterministic causal-chain artifact for a divergence's source data.
 */
import path from "node:path";

import { parseDocx } from "../packages/core/src/docx/parser";
import { toFlowBlocks } from "../packages/core/src/layout-bridge/convert/toFlowBlocks";
import type { FlowBlock, Run } from "../packages/core/src/layout-engine/types";
import { toProseDoc } from "../packages/core/src/prosemirror/conversion/toProseDoc";
import { normalizeLineText } from "./textNorm";

type Flags = {
  doc?: string | undefined;
  text?: string | undefined;
  limit: number;
};

type PMNode = ReturnType<typeof toProseDoc>;

type JsonRecord = Record<string, unknown>;

const usage = (): string =>
  [
    "Usage:",
    "  bun parity/sourceTrace.ts --doc file.docx --text 'Optional Currencies'",
    "",
    "Options:",
    "  --doc <path>      DOCX to parse",
    "  --text <text>     normalized substring to search",
    "  --limit <n>       max matches per layer (default: 8)",
  ].join("\n");

const parseArgs = (argv: string[]): Flags => {
  const flags: Flags = { limit: 8 };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "--doc") {
      flags.doc = argv[++i];
    } else if (arg === "--text") {
      flags.text = argv[++i];
    } else if (arg === "--limit") {
      flags.limit = Number.parseInt(argv[++i] ?? "", 10);
    } else if (arg === "--help" || arg === "-h") {
      console.log(usage());
      process.exit(0);
    } else if (!flags.doc) {
      flags.doc = arg;
    } else if (!flags.text) {
      flags.text = arg;
    } else {
      throw new Error(`Unknown argument: ${arg}`);
    }
  }
  if (!flags.doc) {
    throw new Error("--doc is required");
  }
  if (!flags.text) {
    throw new Error("--text is required");
  }
  if (!Number.isInteger(flags.limit) || flags.limit <= 0) {
    throw new Error("--limit must be a positive integer");
  }
  return flags;
};

const truncate = (value: string, max = 240): string =>
  value.length <= max ? value : `${value.slice(0, max - 1)}…`;

const isRecord = (value: unknown): value is JsonRecord =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const pickFormatting = (value: unknown): unknown => {
  if (!isRecord(value)) {
    return value;
  }
  const keys = [
    "styleId",
    "alignment",
    "spacing",
    "spacingExplicit",
    "indentLeft",
    "indentFirstLine",
    "hangingIndent",
    "tabs",
    "runProperties",
    "numPr",
    "pageBreakBefore",
    "keepNext",
    "keepLines",
    "outlineLevel",
  ];
  return Object.fromEntries(keys.flatMap((key) => (key in value ? [[key, value[key]]] : [])));
};

const textFromDocContent = (content: unknown): string => {
  if (!Array.isArray(content)) {
    return "";
  }
  let text = "";
  for (const item of content) {
    if (!isRecord(item)) {
      continue;
    }
    if (item["type"] === "run") {
      text += textFromDocContent(item["content"]);
    } else if (item["type"] === "text" && typeof item["text"] === "string") {
      text += item["text"];
    } else if (
      (item["type"] === "complexField" || item["type"] === "simpleField") &&
      Array.isArray(item["fieldResult"])
    ) {
      text += textFromDocContent(item["fieldResult"]);
    } else if (item["type"] === "simpleField" && Array.isArray(item["content"])) {
      text += textFromDocContent(item["content"]);
    } else if (item["type"] === "tab") {
      text += "\t";
    }
  }
  return text;
};

const summarizeDocContent = (content: unknown): unknown[] => {
  if (!Array.isArray(content)) {
    return [];
  }
  return content.flatMap((item): unknown[] => {
    if (!isRecord(item)) {
      return [];
    }
    const type = item["type"];
    if (type === "run") {
      return summarizeDocContent(item["content"]);
    }
    if (type === "text") {
      return [{ type, text: truncate(String(item["text"] ?? "")) }];
    }
    if (type === "tab") {
      return [{ type }];
    }
    if (type === "complexField" || type === "simpleField") {
      return [
        {
          type,
          fieldType: item["fieldType"],
          instruction: item["instruction"],
          resultText: truncate(textFromDocContent(item["fieldResult"] ?? item["content"])),
          fldLock: item["fldLock"],
          dirty: item["dirty"],
        },
      ];
    }
    return [{ type }];
  });
};

const pmMarks = (node: PMNode): unknown[] =>
  node.marks.map((mark) => ({ type: mark.type.name, attrs: mark.attrs }));

const summarizePmChild = (node: PMNode, pos: number): unknown => ({
  pos,
  type: node.type.name,
  text: node.isText ? truncate(node.text ?? "") : undefined,
  attrs: node.attrs,
  marks: pmMarks(node),
});

const summarizeRun = (run: Run): unknown => {
  const base = {
    kind: run.kind,
    pmStart: run.pmStart,
    pmEnd: run.pmEnd,
    fontFamily: "fontFamily" in run ? run.fontFamily : undefined,
    fontSize: "fontSize" in run ? run.fontSize : undefined,
    bold: "bold" in run ? run.bold : undefined,
    allCaps: "allCaps" in run ? run.allCaps : undefined,
    smallCaps: "smallCaps" in run ? run.smallCaps : undefined,
  };
  if (run.kind === "text") {
    return { ...base, text: truncate(run.text) };
  }
  if (run.kind === "field") {
    return {
      ...base,
      fieldType: run.fieldType,
      instruction: run.instruction,
      fallback: run.fallback,
      fldLock: run.fldLock,
    };
  }
  if (run.kind === "tab") {
    return base;
  }
  if (run.kind === "lineBreak") {
    return base;
  }
  if (run.kind === "image") {
    return { ...base, width: run.width, height: run.height, wrapType: run.wrapType };
  }
  if (run.kind === "math") {
    return { ...base, plainText: run.plainText };
  }
  return base;
};

const blockText = (block: FlowBlock): string => {
  if (block.kind === "paragraph") {
    return block.runs
      .map((run) => {
        if (run.kind === "text") return run.text;
        if (run.kind === "field") return run.fallback ?? "";
        if (run.kind === "tab") return "\t";
        return "";
      })
      .join("");
  }
  if (block.kind === "table") {
    return block.rows
      .flatMap((row) => row.cells)
      .flatMap((cell) => cell.blocks)
      .map(blockText)
      .join(" ");
  }
  return "";
};

const normalizedSearchText = (value: string): string =>
  normalizeLineText(value).toLocaleLowerCase();

const main = async (): Promise<void> => {
  const flags = parseArgs(process.argv.slice(2));
  const docPath = path.resolve(flags.doc!);
  const needle = normalizedSearchText(flags.text!);
  const parsed = await parseDocx(await Bun.file(docPath).arrayBuffer(), {
    preloadFonts: false,
    detectVariables: false,
  });
  const pmDoc = toProseDoc(parsed);
  const flowBlocks = toFlowBlocks(pmDoc, {
    theme: parsed.package.theme ?? null,
  });

  const docMatches: unknown[] = [];
  for (const [index, block] of parsed.package.document.content.entries()) {
    if (block.type !== "paragraph") {
      continue;
    }
    const text = textFromDocContent(block.content);
    if (!normalizedSearchText(text).includes(needle)) {
      continue;
    }
    docMatches.push({
      index,
      type: block.type,
      text: truncate(text),
      formatting: pickFormatting(block.formatting),
      content: summarizeDocContent(block.content),
    });
    if (docMatches.length >= flags.limit) {
      break;
    }
  }

  const pmMatches: unknown[] = [];
  pmDoc.descendants((node, pos) => {
    if (pmMatches.length >= flags.limit || node.type.name !== "paragraph") {
      return undefined;
    }
    if (!normalizedSearchText(node.textContent).includes(needle)) {
      return undefined;
    }
    const children: unknown[] = [];
    node.descendants((child, childPos) => {
      if (child.isText || child.type.name === "field" || child.type.name === "tab") {
        children.push(summarizePmChild(child, pos + 1 + childPos));
      }
      return undefined;
    });
    pmMatches.push({
      pos,
      text: truncate(node.textContent),
      attrs: node.attrs,
      children,
    });
    return undefined;
  });

  const flowMatches = flowBlocks
    .map((block, index) => ({ block, index, text: blockText(block) }))
    .filter(({ text }) => normalizedSearchText(text).includes(needle))
    .slice(0, flags.limit)
    .map(({ block, index, text }) => ({
      index,
      id: "id" in block ? block.id : undefined,
      kind: block.kind,
      text: truncate(text),
      attrs: block.kind === "paragraph" ? block.attrs : undefined,
      runs: block.kind === "paragraph" ? block.runs.map(summarizeRun) : undefined,
    }));

  console.log(
    JSON.stringify(
      {
        doc: docPath,
        query: flags.text,
        normalizedQuery: needle,
        counts: {
          documentMatches: docMatches.length,
          pmMatches: pmMatches.length,
          flowMatches: flowMatches.length,
        },
        documentMatches: docMatches,
        pmMatches,
        flowMatches,
      },
      null,
      2,
    ),
  );
};

try {
  await main();
} catch (error) {
  const err = error instanceof Error ? error : new Error(String(error));
  console.error(`parity source trace failed: ${err.message}`);
  console.error(usage());
  process.exitCode = 2;
}
