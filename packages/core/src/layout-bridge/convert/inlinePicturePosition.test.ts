import { expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { documentShape } from "../../__tests__/documentShapes";
import { parseDocx } from "../../docx/parser";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureParagraph } from "../../layout-engine/measure/measureParagraph";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { toFlowBlocks } from "./toFlowBlocks";

test("raised and lowered inline pictures change the ascent and descent of a mixed line", async () => {
  const source = await documentShape("image").build();
  const expectPictureLine = async (
    positionHalfPoints: number,
    expected: { ascent: number; descent?: number },
  ) => {
    const zip = await ZipArchive.read(source.slice().buffer);
    const xml = await zip.entry("word/document.xml")?.text();
    if (xml === undefined || !xml.includes("<w:r><w:drawing>")) {
      throw new Error("Expected a picture run in the fixture");
    }
    zip.set(
      "word/document.xml",
      xml.replace(
        "<w:r><w:drawing>",
        `<w:r><w:rPr><w:position w:val="${positionHalfPoints}"/></w:rPr><w:drawing>`,
      ),
    );
    const bytes = await zip.generate({ compression: "store" });
    const document = await parseDocx(bytes.buffer, { preloadFonts: false });
    const paragraph = toFlowBlocks(toProseDoc(document)).find(
      (block) =>
        block.kind === "paragraph" &&
        block.runs.some((run) => run.kind === "image"),
    );
    if (paragraph?.kind !== "paragraph") {
      throw new Error("Expected a paragraph containing the picture");
    }
    withFakeTextMeasure(
      () => {
        expect(measureParagraph(paragraph, 600).lines.at(0)).toMatchObject(
          expected,
        );
      },
      { charWidth: fixedCharWidth(10) },
    );
  };

  // The 96px picture's run moves by 6pt (8px) in either direction.
  await expectPictureLine(0, { ascent: 96 });
  await expectPictureLine(12, { ascent: 104 });
  await expectPictureLine(-12, { ascent: 88, descent: 8 });
});
