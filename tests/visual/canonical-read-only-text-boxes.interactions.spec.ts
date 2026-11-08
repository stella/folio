import { expect } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../../packages/core/src/docx/rezip";
import { schema } from "../../packages/core/src/prosemirror/schema";
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

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="${R}" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;
const BOX_TEXT =
  '<w:p w14:paraId="20000001"><w:r><w:t>Inside</w:t></w:r></w:p>';
const drawing = `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="4000000" cy="2000000"/><wp:docPr id="1" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${BOX_TEXT}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const boxedDocument = async () => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>` +
      `<w:p w14:paraId="10000001"><w:r><w:t>Host</w:t></w:r>${drawing}</w:p>` +
      '<w:p w14:paraId="10000002"><w:r><w:t>Next</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

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

test("a text box opens read-only: typing inside it changes nothing, typing beside it edits", async ({
  page,
}) => {
  test.setTimeout(60_000);
  const source = await canonicalLoadFixture(await boxedDocument());
  for (const port of [reactPort, vuePort]) {
    await page.goto(`http://localhost:${port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    await clearRefusals(page);
    await loadReady(page, source);
    await expectNoRefusals(page);
    const opened = await snapshot(page);
    expect(opened?.text).toBe("HostInsideNext");
    const content = JSON.stringify(opened?.document?.package.document.content);

    // The box's text is not part of the story the session edits, so the
    // keystroke reaches no paragraph and changes nothing.
    await select(page, positionIn(opened?.projectionJSON, "Inside", 3));
    await page.keyboard.type("x");
    const untouched = await snapshot(page);
    expect(untouched?.text).toBe("HostInsideNext");
    expect(untouched?.projectionMatchesCanonical).toBe(true);
    expect(JSON.stringify(untouched?.document?.package.document.content)).toBe(
      content,
    );

    await clearRefusals(page);
    await select(page, positionIn(opened?.projectionJSON, "Next", 4));
    await page.keyboard.type("!");
    const typed = await snapshot(page);
    expect(typed?.text).toBe("HostInsideNext!");
    expect(typed?.projectionMatchesCanonical).toBe(true);
    await expectNoRefusals(page);
    const saved = await page.evaluate(async () => {
      const buffer = await globalThis.__folioCanonical?.save();
      return buffer ? [...new Uint8Array(buffer)] : null;
    });
    if (!saved) throw new TypeError("The canonical document did not save.");
    const xml =
      (await (
        await ZipArchive.read(Uint8Array.from(saved).buffer)
      )
        .entry("word/document.xml")
        ?.text()) ?? "";
    expect(xml).toContain(`<w:txbxContent>${BOX_TEXT}</w:txbxContent>`);
    expect(xml).toContain("Next!");
  }
});
