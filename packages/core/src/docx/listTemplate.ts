import type { CounterFormat } from "../types/document";
import { formatOoxmlCounter } from "./ooxmlCounterFormatter";

export type ResolvedListComponent = {
  level: number;
  relativeStart: number;
  start: number;
  end: number;
};

export type ResolvedListTemplate = {
  value: string;
  components: readonly ResolvedListComponent[];
};

type ResolveListTemplateOptions = {
  template: string;
  counters: readonly number[];
  levelFormats?: readonly (CounterFormat | undefined)[] | undefined;
  forceDecimal?: boolean | undefined;
};

export function resolveListTemplate(options: ResolveListTemplateOptions): string {
  return resolveListTemplateWithComponents(options).value;
}

export function resolveListTemplateWithComponents({
  template,
  counters,
  levelFormats,
  forceDecimal = false,
}: ResolveListTemplateOptions): ResolvedListTemplate {
  const components: ResolvedListComponent[] = [];
  let value = "";
  let sourceEnd = 0;
  for (const match of template.matchAll(/%(?<digit>\d)(?<punct>[.):\]])?/gu)) {
    const sourceStart = match.index;
    const relativeStart = value.length;
    value += template.slice(sourceEnd, sourceStart);
    sourceEnd = sourceStart + match[0].length;

    const index = Number.parseInt(match.groups?.["digit"] ?? "", 10) - 1;
    const counter = counters[index];
    if (index < 0 || counter === undefined || !Number.isFinite(counter)) {
      value += match.groups?.["punct"] ?? "";
      continue;
    }
    const formatted = formatOoxmlCounter(counter, forceDecimal ? "decimal" : levelFormats?.[index]);
    if (!formatted) {
      continue;
    }
    const start = value.length;
    value += formatted;
    components.push({ level: index, relativeStart, start, end: value.length });
    value += match.groups?.["punct"] ?? "";
  }
  value += template.slice(sourceEnd);
  return { value, components };
}
