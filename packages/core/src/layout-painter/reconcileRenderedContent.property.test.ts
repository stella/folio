import { expect, test } from "bun:test";
import fc from "fast-check";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import { assertProperty } from "../../../../test/property-testing";
import { reconcileRenderedContent } from "./reconcileRenderedContent";

type Line = {
  pmStart: number;
  text: string;
  bold: boolean;
  spans: number;
};

/** A painted paragraph: a block with its position and some spans of text. */
const paragraphElement = (line: Line, shift: number): HTMLElement => {
  const paragraph = document.createElement("div");
  paragraph.className = "layout-paragraph";
  paragraph.dataset["pmStart"] = String(line.pmStart + shift);
  paragraph.dataset["pmEnd"] = String(line.pmStart + shift + line.text.length);
  for (let index = 0; index < line.spans; index += 1) {
    const span = document.createElement("span");
    span.dataset["pmStart"] = String(line.pmStart + shift + index);
    if (line.bold) span.style.setProperty("font-weight", "700");
    span.textContent = `${line.text}${index}`;
    paragraph.append(span);
  }
  return paragraph;
};

const contentElement = (lines: readonly Line[], shifts: readonly number[]) => {
  const content = document.createElement("div");
  content.className = "layout-page-content";
  for (const [index, line] of lines.entries())
    content.append(paragraphElement(line, shifts[index] ?? 0));
  return content;
};

const line = fc.record({
  pmStart: fc.integer({ min: 0, max: 500 }),
  text: fc.constantFrom("alpha", "beta", "gamma", "x"),
  bold: fc.boolean(),
  spans: fc.integer({ min: 0, max: 3 }),
});

/** What an edit does to one painted line in the fresh rendering. */
const change = fc.constantFrom(
  "same",
  "shifted",
  "text",
  "bold",
  "spans",
) satisfies fc.Arbitrary<"same" | "shifted" | "text" | "bold" | "spans">;

test("reconciling painted content gives the fresh rendering and keeps the lines it shows the same", () => {
  const releaseDom = acquireHarnessDom();
  try {
    assertProperty(
      fc.property(
        fc.array(fc.tuple(line, change), { minLength: 0, maxLength: 8 }),
        fc.integer({ min: -3, max: 3 }),
        fc.boolean(),
        (lines, delta, grows) => {
          const before = lines.map(([painted]) => painted);
          const after = lines.map(([painted, kind]): Line => {
            switch (kind) {
              case "same":
              case "shifted":
                return painted;
              case "text":
                return { ...painted, text: `${painted.text}!` };
              case "bold":
                return { ...painted, bold: !painted.bold };
              case "spans":
                return { ...painted, spans: painted.spans + 1 };
              default: {
                const exhaustive: never = kind;
                return exhaustive;
              }
            }
          });
          const shifts = lines.map(([, kind]) => (kind === "same" ? 0 : delta));
          if (grows)
            after.push({ pmStart: 0, text: "new", bold: false, spans: 1 });
          const painted = contentElement(
            before,
            before.map(() => 0),
          );
          const fresh = contentElement(after, shifts);
          const expected = fresh.cloneNode(true);
          const kept = Array.from(painted.childNodes);
          const reconciled = reconcileRenderedContent(painted, fresh);
          if (grows) {
            // A different number of lines is the caller's to replace whole.
            expect(reconciled).toBe(false);
            return;
          }
          expect(reconciled).toBe(true);
          // The content is now exactly what painting it afresh gives.
          expect(painted.isEqualNode(expected)).toBe(true);
          // And a line the edit only shifted, or left alone, is the node that
          // was painted before.
          for (const [index, [, kind]] of lines.entries()) {
            const reused = painted.childNodes[index] === kept[index];
            if (kind === "same" || kind === "shifted")
              expect(reused).toBe(true);
          }
        },
      ),
      {
        id: "reconciling painted content gives the fresh rendering and keeps the lines it shows the same",
      },
    );
  } finally {
    releaseDom();
  }
});
