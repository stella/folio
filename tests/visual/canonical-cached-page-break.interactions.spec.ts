import { expect } from "@playwright/test";

import { parseDocx } from "../../packages/core/src/docx/parser";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
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

const runContentTypes = (
  content: Awaited<
    ReturnType<typeof parseDocx>
  >["package"]["document"]["content"],
) =>
  content.flatMap((block) =>
    block.type === "paragraph"
      ? block.content.flatMap((item) =>
          item.type === "run" ? item.content.map(({ type }) => type) : [],
        )
      : [],
  );

test("a paginated document opens, edits beside its cached page break and keeps it", async ({
  page,
}) => {
  test.setTimeout(60_000);
  const document = createEmptyDocument({ initialText: "Alpha" });
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph")
    throw new TypeError("Expected the source paragraph.");
  // A paginating producer records where it broke the page, in a run before the text.
  paragraph.content.unshift({
    type: "run",
    content: [{ type: "renderedPageBreak" }],
  });
  const source = await canonicalLoadFixture(await createDocx(document));
  for (const port of [reactPort, vuePort]) {
    await page.goto(`http://localhost:${port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    await clearRefusals(page);
    await loadReady(page, source);
    await expectNoRefusals(page);
    // The break is one zero-width position before the text: "Alpha" ends at 7.
    await select(page, 7);
    await page.keyboard.type("x");
    const typed = await snapshot(page);
    expect(typed?.text).toBe("Alphax");
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
    const content = reopened.package.document.content;
    expect(runContentTypes(content)).toContain("renderedPageBreak");
    expect(
      content
        .flatMap((block) =>
          block.type === "paragraph"
            ? block.content.flatMap((item) =>
                item.type === "run"
                  ? item.content.flatMap((leaf) =>
                      leaf.type === "text" ? [leaf.text] : [],
                    )
                  : [],
              )
            : [],
        )
        .join(""),
    ).toBe("Alphax");
  }
});
