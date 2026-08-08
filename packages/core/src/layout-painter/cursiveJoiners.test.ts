/**
 * Joiner planning has two failure modes that both matter: skipping a boundary
 * that needed repair (a cursive word paints broken), and repairing one that did
 * not (a stray span, and a joiner where Word would have none). The face rule is
 * what separates them, so these drive it through a stand-in `applyRunStyles`
 * shaped like the painter's.
 *
 * Characters are `\u` escapes: a pasted glyph once corrupted a character class
 * elsewhere in this package.
 */

import { describe, expect, test } from "bun:test";

import {
  createJoinerSpan,
  JOINER_DATASET_KEY,
  planCursiveJoiners,
  withCursiveJoiners,
} from "./cursiveJoiners";

const MEEM = "م";
const KAF = "ك";
const TEH = "ت";
const BEH = "ب";
const ALEF = "ا";
const ZWJ = "‍";

type TestRun = {
  text?: string;
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  /** A shaping-neutral property: differs without changing the face. */
  color?: string;
};

/** Stands in for the painter's `applyRunStyles`, including a neutral property. */
const applyRunStyles = (element: HTMLElement, run: TestRun): void => {
  const style = element.style as unknown as Record<string, string>;
  if (run.fontFamily) style["fontFamily"] = run.fontFamily;
  if (run.fontSize) style["fontSize"] = `${run.fontSize}px`;
  if (run.bold) style["fontWeight"] = "700";
  if (run.italic) style["fontStyle"] = "italic";
  if (run.color) style["color"] = run.color;
};

/** The painter's minimal fake Document: no HTMLElement, no insertBefore. */
class FakeElement {
  dataset: Record<string, string> = {};
  style: Record<string, string> = {};
  children: FakeElement[] = [];
  textContent = "";
  append(...children: FakeElement[]): void {
    this.children.push(...children);
  }
}

const doc = {
  createElement: () => new FakeElement(),
} as unknown as Document;

const plan = (runs: TestRun[]) =>
  planCursiveJoiners<TestRun>({
    runs,
    textOf: (run) => run.text,
    applyRunStyles,
    doc,
  });

const BASE: TestRun = { fontFamily: "Arial", fontSize: 32 };
const run = (text: string, extra: Partial<TestRun> = {}): TestRun => ({
  ...BASE,
  text,
  ...extra,
});

/** Total painted text of a run sequence, joiners included. */
function paintedText(runs: TestRun[]): string {
  const joinerPlan = plan(runs);
  let out = "";
  for (const item of runs) {
    const runEl = doc.createElement("span");
    (runEl as unknown as FakeElement).textContent = item.text ?? "";
    for (const el of withCursiveJoiners(runEl, item, joinerPlan, applyRunStyles, doc)) {
      out += (el as unknown as FakeElement).textContent;
    }
  }
  return out;
}

describe("planCursiveJoiners", () => {
  test("repairs a bold split inside a cursive word", () => {
    const runs = [run(MEEM + KAF), run(TEH + BEH, { bold: true })];
    const joinerPlan = plan(runs);
    expect(joinerPlan.get(runs[0]!)).toEqual({ leading: false, trailing: true });
    expect(joinerPlan.get(runs[1]!)).toEqual({ leading: true, trailing: false });
    // One joiner each side, so both halves shape into a joined form.
    expect(paintedText(runs)).toBe(MEEM + KAF + ZWJ + ZWJ + TEH + BEH);
  });

  test("repairs italic, size and family changes too", () => {
    for (const extra of [
      { italic: true },
      { fontSize: 40 },
      { fontFamily: "Courier New" },
    ] as Partial<TestRun>[]) {
      const runs = [run(MEEM + KAF), run(TEH + BEH, extra)];
      expect(plan(runs).size).toBe(2);
    }
  });

  test("leaves a same-face boundary alone — the browser already shapes across it", () => {
    // The tracked-change and comment-anchor shape: a shaping-neutral property
    // differs, the face does not.
    const runs = [run(MEEM + KAF), run(TEH + BEH, { color: "#c00" })];
    expect(plan(runs).size).toBe(0);
    expect(paintedText(runs)).toBe(MEEM + KAF + TEH + BEH);
  });

  test("an unknown style property counts as a face change, so it fails safe", () => {
    // A property this module has never heard of must push toward inserting a
    // harmless joiner, never toward silently skipping the repair.
    const applyWithUnknown = (element: HTMLElement, item: TestRun): void => {
      applyRunStyles(element, item);
      if (item.bold) {
        (element.style as unknown as Record<string, string>)["fontStretch"] = "condensed";
      }
    };
    const runs = [run(MEEM + KAF), run(TEH + BEH, { bold: true })];
    const joinerPlan = planCursiveJoiners<TestRun>({
      runs,
      textOf: (item) => item.text,
      applyRunStyles: applyWithUnknown,
      doc,
    });
    expect(joinerPlan.size).toBe(2);
  });

  test("does not join where the letters were never connected", () => {
    // A right-joining letter before the cut connects backwards only, so a bold
    // change after alef needs no joiner even though the face changes.
    expect(plan([run(KAF + ALEF), run(TEH + BEH, { bold: true })]).size).toBe(0);
  });

  test("does not join across a word boundary", () => {
    expect(plan([run(MEEM + KAF + " "), run(TEH + BEH, { bold: true })]).size).toBe(0);
  });

  test("does not join across a run that paints no text", () => {
    // An image or tab between two words means they were never connected.
    const runs = [run(MEEM + KAF), { ...BASE }, run(TEH + BEH, { bold: true })];
    expect(plan(runs).size).toBe(0);
  });

  test("never touches a Latin line, however it is styled", () => {
    const runs = [run("smlouva "), run("o dílo", { bold: true }), run(" 2026", { fontSize: 12 })];
    expect(plan(runs).size).toBe(0);
    expect(paintedText(runs)).toBe("smlouva o dílo 2026");
  });

  test("repairs every affected boundary on a line with several", () => {
    const runs = [run(MEEM), run(KAF, { bold: true }), run(TEH), run(BEH, { bold: true })];
    const joinerPlan = plan(runs);
    expect(joinerPlan.size).toBe(4);
    // The middle runs need a joiner on BOTH sides.
    expect(joinerPlan.get(runs[1]!)).toEqual({ leading: true, trailing: true });
    expect(joinerPlan.get(runs[2]!)).toEqual({ leading: true, trailing: true });
    expect(paintedText(runs)).toBe(MEEM + ZWJ + ZWJ + KAF + ZWJ + ZWJ + TEH + ZWJ + ZWJ + BEH);
  });

  test("a single run, or none, is left alone", () => {
    expect(plan([]).size).toBe(0);
    expect(plan([run(MEEM + KAF + TEH + BEH)]).size).toBe(0);
  });

  // A real CSSStyleDeclaration is not a plain object: `Object.entries` on one
  // returns its INDEXED entries, so reading it that way compares the list of
  // property names rather than their values, and every colour-only boundary
  // looks like a face change. That shipped once and only the browser oracle
  // caught it, so the real shape is covered here too.
  test("reads a real CSSStyleDeclaration, not just the plain-object fake", () => {
    class RealishStyle {
      private readonly properties = new Map<string, string>();
      setProperty(property: string, value: string): void {
        this.properties.set(property, value);
      }
      getPropertyValue(property: string): string {
        return this.properties.get(property) ?? "";
      }
      item(index: number): string {
        return [...this.properties.keys()][index] ?? "";
      }
      get length(): number {
        return this.properties.size;
      }
    }
    class RealishElement {
      dataset: Record<string, string> = {};
      style = new RealishStyle();
      textContent = "";
    }
    const realishDoc = {
      createElement: () => new RealishElement(),
    } as unknown as Document;

    // Mirrors the painter: kebab-case names, colour set as both a property and
    // a custom property, exactly as `applyRunStyles` does.
    const applyRealish = (element: HTMLElement, item: TestRun): void => {
      const style = element.style;
      if (item.fontFamily) style.setProperty("font-family", item.fontFamily);
      if (item.fontSize) style.setProperty("font-size", `${item.fontSize}px`);
      if (item.bold) style.setProperty("font-weight", "700");
      if (item.color) {
        style.setProperty("color", item.color);
        style.setProperty("--doc-run-color", item.color);
      }
    };
    const planWith = (runs: TestRun[]) =>
      planCursiveJoiners<TestRun>({
        runs,
        textOf: (item) => item.text,
        applyRunStyles: applyRealish,
        doc: realishDoc,
      });

    // Colour only: same face, so the browser already joins and no joiner belongs.
    expect(planWith([run(MEEM + KAF), run(TEH + BEH, { color: "#c00" })]).size).toBe(0);
    // Bold: a genuine face change, still repaired.
    expect(planWith([run(MEEM + KAF), run(TEH + BEH, { bold: true })]).size).toBe(2);
  });
});

describe("createJoinerSpan", () => {
  test("wears the run's own painted styles", () => {
    // Without the matching face, shaping would not cross into the joiner and
    // the repair would be a no-op that still looked applied.
    const bold = createJoinerSpan(run(TEH, { bold: true }), applyRunStyles, doc);
    const regular = createJoinerSpan(run(KAF), applyRunStyles, doc);
    expect((bold.style as unknown as Record<string, string>)["fontWeight"]).toBe("700");
    expect((regular.style as unknown as Record<string, string>)["fontWeight"]).toBeUndefined();
  });

  test("is furniture: joiner flag set, no pm positions", () => {
    const span = createJoinerSpan(run(TEH), applyRunStyles, doc);
    expect(span.dataset[JOINER_DATASET_KEY]).toBe("true");
    expect(span.dataset["pmStart"]).toBeUndefined();
    expect(span.dataset["pmEnd"]).toBeUndefined();
    expect(span.textContent).toBe(ZWJ);
  });
});

describe("withCursiveJoiners", () => {
  test("returns the element alone when nothing needs repair", () => {
    const runs = [run("plain"), run("text", { bold: true })];
    const runEl = doc.createElement("span");
    expect(withCursiveJoiners(runEl, runs[0]!, plan(runs), applyRunStyles, doc)).toEqual([runEl]);
  });
});
