import { createHash } from "node:crypto";

export const LAYOUT_INTERACTION_AXES = {
  section: ["single", "continuous", "nextPage", "twoColumn"],
  anchorFrame: ["inline", "page", "margin", "column", "paragraph"],
  wrap: ["inline", "none", "square", "topBottom"],
  flow: ["normal", "keepNext", "keepLines", "hardPageBreak", "renderedPageBreak"],
  table: ["none", "fixed", "autofit", "merged", "splitRow"],
  typography: ["latin", "rtl", "cjk", "tabs", "numbering"],
} as const;

type LayoutInteractionAxes = typeof LAYOUT_INTERACTION_AXES;
export type LayoutInteractionScenario = {
  readonly [K in keyof LayoutInteractionAxes]: LayoutInteractionAxes[K][number];
};

export type LayoutInteractionCase = LayoutInteractionScenario & {
  readonly id: string;
};

export const LAYOUT_INTERACTION_AXIS_NAMES = [
  "section",
  "anchorFrame",
  "wrap",
  "flow",
  "table",
  "typography",
] as const satisfies readonly (keyof LayoutInteractionAxes)[];

const isValidScenario = ({ anchorFrame, wrap }: LayoutInteractionScenario): boolean =>
  (anchorFrame === "inline") === (wrap === "inline");

const scenarios = (): readonly LayoutInteractionScenario[] => {
  const result: LayoutInteractionScenario[] = [];
  for (const section of LAYOUT_INTERACTION_AXES.section) {
    for (const anchorFrame of LAYOUT_INTERACTION_AXES.anchorFrame) {
      for (const wrap of LAYOUT_INTERACTION_AXES.wrap) {
        for (const flow of LAYOUT_INTERACTION_AXES.flow) {
          for (const table of LAYOUT_INTERACTION_AXES.table) {
            for (const typography of LAYOUT_INTERACTION_AXES.typography) {
              const scenario = { section, anchorFrame, wrap, flow, table, typography };
              if (isValidScenario(scenario)) result.push(scenario);
            }
          }
        }
      }
    }
  }
  return result;
};

const pairKeys = (scenario: LayoutInteractionScenario): readonly string[] => {
  const pairs: string[] = [];
  for (let leftIndex = 0; leftIndex < LAYOUT_INTERACTION_AXIS_NAMES.length; leftIndex += 1) {
    const left = LAYOUT_INTERACTION_AXIS_NAMES[leftIndex];
    if (!left) continue;
    for (
      let rightIndex = leftIndex + 1;
      rightIndex < LAYOUT_INTERACTION_AXIS_NAMES.length;
      rightIndex += 1
    ) {
      const right = LAYOUT_INTERACTION_AXIS_NAMES[rightIndex];
      if (!right) continue;
      pairs.push(`${left}=${scenario[left]}|${right}=${scenario[right]}`);
    }
  }
  return pairs;
};

const CANDIDATES = scenarios().map((scenario) => ({ scenario, pairs: pairKeys(scenario) }));
const VALID_PAIRS = new Set(CANDIDATES.flatMap(({ pairs }) => pairs));

const caseId = (scenario: LayoutInteractionScenario): string =>
  `mx-${createHash("sha256").update(JSON.stringify(scenario)).digest("hex").slice(0, 10)}`;

export const validLayoutInteractionPairs = (): ReadonlySet<string> => new Set(VALID_PAIRS);

export const layoutInteractionCasePairs = (
  scenario: LayoutInteractionScenario,
): ReadonlySet<string> => new Set(pairKeys(scenario));

const generateLayoutInteractionMatrix = (): readonly LayoutInteractionCase[] => {
  const uncovered = new Set(VALID_PAIRS);
  const selected: LayoutInteractionCase[] = [];

  while (uncovered.size > 0) {
    let best: LayoutInteractionScenario | undefined;
    let bestPairs: readonly string[] = [];
    let bestCoverage = 0;
    for (const { scenario, pairs } of CANDIDATES) {
      let coverage = 0;
      for (const pair of pairs) {
        if (uncovered.has(pair)) coverage += 1;
      }
      if (coverage > bestCoverage) {
        best = scenario;
        bestPairs = pairs;
        bestCoverage = coverage;
      }
    }
    if (!best || bestCoverage === 0) {
      throw new TypeError("layout interaction matrix could not cover every valid pair");
    }
    selected.push({ id: caseId(best), ...best });
    for (const pair of bestPairs) uncovered.delete(pair);
  }

  return selected;
};

const LAYOUT_INTERACTION_MATRIX = generateLayoutInteractionMatrix();

export const buildLayoutInteractionMatrix = (): readonly LayoutInteractionCase[] =>
  LAYOUT_INTERACTION_MATRIX.map(({ id, section, anchorFrame, wrap, flow, table, typography }) => ({
    id,
    section,
    anchorFrame,
    wrap,
    flow,
    table,
    typography,
  }));
