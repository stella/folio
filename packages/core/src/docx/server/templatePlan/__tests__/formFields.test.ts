import { beforeAll, describe, expect, test } from "bun:test";

import { buildTemplate, p, r, readPart } from "./fixtures";
import { applyFor, buildSourceWith, unwrap } from "./planning";

const begin = (formData: string): string =>
  `<w:r><w:fldChar w:fldCharType="begin"><w:ffData>${formData}</w:ffData></w:fldChar></w:r>`;
const instruction = (code: string): string =>
  `<w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r>`;
const separate = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

const CHECK_BOX =
  '<w:name w:val="Agree"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="1"/></w:checkBox>';
const DROP_DOWN =
  '<w:name w:val="Term"/><w:enabled/><w:calcOnExit w:val="0"/><w:ddList><w:result w:val="1"/><w:listEntry w:val="one year"/><w:listEntry w:val="two years"/></w:ddList>';
const TEXT_INPUT =
  '<w:name w:val="Party"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput><w:default w:val="[Name]"/><w:maxLength w:val="40"/></w:textInput>';

/** A source with a check box and a drop-down list that cache no text, and a
 *  text input whose result is its default. */
const sourceWithFormFields = async (): Promise<Uint8Array> =>
  await buildSourceWith({
    body: [
      p(
        `${begin(CHECK_BOX)}${instruction("FORMCHECKBOX")}${end}${r(" The parties agree.")}`,
      ),
      p(
        `${r("The term is ")}${begin(DROP_DOWN)}${instruction("FORMDROPDOWN")}${end}${r(".")}`,
      ),
      p(
        `${r("Signed by ")}${begin(TEXT_INPUT)}${instruction("FORMTEXT")}${separate}${r("[Name]")}${end}${r(" on the date below.")}`,
      ),
    ].join(""),
  });

let template: Uint8Array;

beforeAll(async () => {
  template = await buildTemplate();
});

describe("legacy form fields of the source", () => {
  test("stay fields with their box, list and default", async () => {
    const applied = unwrap(
      await applyFor({ source: await sourceWithFormFields(), template }),
    );
    const document = (await readPart(applied.docx, "word/document.xml")) ?? "";
    expect(document).toContain('<w:default w:val="1"/>');
    expect(document).toContain('<w:listEntry w:val="two years"/>');
    expect(document).toContain('<w:default w:val="[Name]"/>');
    expect(document.match(/FORMCHECKBOX|FORMDROPDOWN|FORMTEXT/gu)).toEqual([
      "FORMCHECKBOX",
      "FORMDROPDOWN",
      "FORMTEXT",
    ]);
    // The text input's result stays its text, inside the field.
    expect(document).toMatch(
      /FORMTEXT <\/w:instrText><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:fldChar w:fldCharType="separate"\/><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t xml:space="preserve">\[Name\]<\/w:t><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:fldChar w:fldCharType="end"\/>/u,
    );
    expect(applied.report.dropped.sourceFields).toEqual([]);
  });
});
