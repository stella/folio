import { expect, test } from "bun:test";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import { applyTextAdvanceGaps } from "./textAdvanceGapsDom";

/** The run's text nodes, in order. */
const textNodes = (host: HTMLElement): string[] => {
  const walker = host.ownerDocument.createTreeWalker(host, 4);
  const out: string[] = [];
  let node = walker.nextNode();
  while (node !== null) {
    out.push(node.nodeValue ?? "");
    node = walker.nextNode();
  }
  return out;
};

const withRun = (text: string, check: (host: HTMLElement) => void): void => {
  const releaseDom = acquireHarnessDom();
  try {
    const host = document.createElement("span");
    host.textContent = text;
    check(host);
  } finally {
    releaseDom();
  }
};

// A collapsed range in an empty text node has no geometry, so a caret mapped
// onto one lands at the page origin instead of the source boundary.
test("a gap covering a whole run leaves no empty text node around it", () => {
  withRun("𠀀", (host) => {
    applyTextAdvanceGaps(host, [{ offset: 2, advancePx: 4 }]);
    expect(textNodes(host)).toEqual(["𠀀"]);
    expect(host.innerHTML).toBe('<span style="letter-spacing: 4px;">𠀀</span>');
  });
});

test("adjacent gaps leave no empty text node between them", () => {
  withRun("Aé́漢1", (host) => {
    applyTextAdvanceGaps(host, [
      { offset: 3, advancePx: 3.5 },
      { offset: 4, advancePx: 4 },
    ]);
    expect(textNodes(host)).toEqual(["A", "é́", "漢", "1"]);
  });
});
