import { expect, test } from "bun:test";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import { applyPageShellStyles } from "./pageShellStyles";

type Write = { property: string; before: string; after: string };

/** The style writes `apply` makes to a shell that already holds its styles. */
const writesOnStyledShell = (apply: (element: HTMLElement) => void) => {
  const releaseDom = acquireHarnessDom();
  try {
    const element = document.createElement("div");
    apply(element);
    const style = element.style;
    const writes: Write[] = [];
    const recorder = new Proxy(style, {
      get: (target, property) => Reflect.get(target, property, target),
      set: (target, property, value) => {
        const before = String(Reflect.get(target, property, target));
        const done = Reflect.set(target, property, value, target);
        writes.push({
          property: String(property),
          before,
          after: String(Reflect.get(target, property, target)),
        });
        return done;
      },
    });
    Object.defineProperty(element, "style", { value: recorder });
    apply(element);
    return { writes, style };
  } finally {
    releaseDom();
  }
};

test("styling a page shell again changes none of its values on the way", () => {
  const { writes, style } = writesOnStyledShell((element) =>
    applyPageShellStyles(element, 816, 1056, {}),
  );
  const properties = writes.map(({ property }) => property);
  expect(properties.length).toBe(new Set(properties).size);
  expect(writes.filter(({ before, after }) => before !== after)).toEqual([]);
  expect(style.color).toBe("var(--doc-canvas-text, #000)");
  expect([style.width, style.height]).toEqual(["816px", "1056px"]);
});
