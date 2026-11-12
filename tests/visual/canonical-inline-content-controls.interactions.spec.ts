import { expect } from "@playwright/test";

import { parseDocx } from "../../packages/core/src/docx/parser";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { schema } from "../../packages/core/src/prosemirror/schema";
import type { Document } from "../../packages/core/src/types/document";
import {
  clearRefusals,
  expectNoRefusals,
  loadReady,
  reactPort,
  select,
  snapshot,
  vuePort,
} from "../parity/canonicalInteractionAssertions";
import { canonicalLoadFixture } from "../parity/canonicalLoadFixture";
import { test } from "../parity/canonicalTest";

const run = (text: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text }],
});

const controlDocument = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [
            run("Before "),
            {
              type: "inlineSdt",
              properties: { sdtType: "richText", tag: "clause", id: 7 },
              content: [run("Clause")],
            },
            run(" after"),
          ],
        },
      ],
    },
  },
});

/** The editor position `offset` characters into the text node `text`. */
const positionIn = (json: unknown, text: string, offset: number) => {
  let found = -1;
  schema.nodeFromJSON(json).descendants((node, position) => {
    if (found >= 0) return false;
    if (node.isText && node.text === text) found = position + offset;
    return true;
  });
  expect(found).toBeGreaterThan(0);
  return found;
};

test("typing inside an inline content control keeps the text in the control", async ({
  page,
}) => {
  test.setTimeout(60_000);
  const source = await canonicalLoadFixture(
    await createDocx(controlDocument()),
  );
  for (const port of [reactPort, vuePort]) {
    await page.goto(`http://localhost:${port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    await clearRefusals(page);
    await loadReady(page, source);
    await expectNoRefusals(page);
    const opened = await snapshot(page);
    await select(page, positionIn(opened?.projectionJSON, "Clause", 3));
    await page.keyboard.type("X");
    const typed = await snapshot(page);
    expect(typed?.text).toBe("Before ClaXuse after");
    expect(typed?.projectionMatchesCanonical).toBe(true);
    await expectNoRefusals(page);
    const saved = await page.evaluate(async () => {
      const buffer = await globalThis.__folioCanonical?.save();
      return buffer ? [...new Uint8Array(buffer)] : null;
    });
    if (!saved) throw new TypeError("The canonical document did not save.");
    const reopened = await parseDocx(new Uint8Array(saved), {
      preloadFonts: false,
      detectVariables: false,
    });
    const paragraph = reopened.package.document.content.at(0);
    if (paragraph?.type !== "paragraph")
      throw new TypeError("The paragraph is absent.");
    const controls = paragraph.content.filter(
      (item) => item.type === "inlineSdt",
    );
    expect(controls).toHaveLength(1);
    expect(JSON.stringify(controls[0])).toContain("ClaXuse");
  }
});
