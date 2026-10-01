import { expect, test } from "bun:test";
import { panic } from "better-result";

import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { BlockContent } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";
import { parseRun } from "./runParser";
import { serializeNewEndnotesPart, serializeNewFootnotesPart } from "./serializer/noteSerializer";
import { serializeRun } from "./serializer/runSerializer";
import { getChildElements, parseXmlDocument } from "./xmlParser";

const TRANSITIONAL = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
for (const namespace of [TRANSITIONAL, "http://purl.oclc.org/ooxml/wordprocessingml/main"]) {
  for (const prefix of ["producer:", ""]) {
    for (const kind of ["footnote", "endnote"] as const) {
      test(`note marker ${kind} is typed and zero-width under ${namespace} (${prefix || "default namespace"})`, () => {
        const binding = prefix ? `xmlns:producer="${namespace}"` : `xmlns="${namespace}"`;
        const root =
          parseXmlDocument(
            `<${prefix}r ${binding}><${prefix}${kind}Ref/><${prefix}t>note</${prefix}t></${prefix}r>`,
          ) ?? panic("Expected run XML");
        const run = parseRun(root, null, null);
        expect(run.content).toEqual([
          { type: "noteMarker", kind },
          { type: "text", text: "note" },
        ]);
        const prose = toProseDoc({
          package: { document: { content: [{ type: "paragraph", content: [run] }] } },
        });
        expect(prose.textContent).toBe("note");
        expect(prose.content.size).toBe(6);
        const serialized = serializeRun(run);
        const wrapper =
          parseXmlDocument(`<root xmlns:w="${TRANSITIONAL}">${serialized}</root>`) ??
          panic("Expected serialized run");
        const savedRun = getChildElements(wrapper).at(0) ?? panic("Expected serialized child");
        expect(parseRun(savedRun, null, null).content).toEqual(run.content);
      });
    }
  }
}

for (const kind of ["footnote", "endnote"] as const) {
  test(`foreign ${kind}Ref markup is preserved and cannot suppress a real note marker`, () => {
    const root =
      parseXmlDocument(
        `<w:r xmlns:w="${TRANSITIONAL}" xmlns:foreign="urn:foreign"><foreign:${kind}Ref/></w:r>`,
      ) ?? panic("Expected foreign run");
    const run = parseRun(root, null, null);
    expect(run.content.at(0)?.type).toBe("preservedXml");
    const content = [{ type: "paragraph", content: [run] }] satisfies BlockContent[];
    const xml =
      kind === "footnote"
        ? serializeNewFootnotesPart([{ type: "footnote", id: 1, content }])
        : serializeNewEndnotesPart([{ type: "endnote", id: 1, content }]);
    expect(xml).toContain(`<w:${kind}Ref/>`);
    expect(xml).toContain(`<foreign:${kind}Ref`);
  });
}

for (const kind of ["footnote", "endnote"] as const) {
  for (const source of ["fresh", "existing note part"] as const) {
    test(`${kind} custom marks survive ${source} serialization without an automatic mark`, async () => {
      const initial = createEmptyDocument({ initialText: "Body" });
      const initialBody = initial.package.document.content.at(0);
      if (!initialBody || initialBody.type !== "paragraph") panic("Expected body");
      if (source === "existing note part") {
        initialBody.content.push({
          type: "run",
          content: [{ type: kind === "footnote" ? "footnoteRef" : "endnoteRef", id: 1 }],
        });
        const content = [
          {
            type: "paragraph",
            content: [{ type: "run", content: [{ type: "text", text: "Automatic note" }] }],
          },
        ] satisfies BlockContent[];
        if (kind === "footnote") initial.package.footnotes = [{ type: "footnote", id: 1, content }];
        else initial.package.endnotes = [{ type: "endnote", id: 1, content }];
      }
      const document =
        source === "fresh"
          ? initial
          : await parseDocx(await createDocx(initial), { preloadFonts: false });
      const body = document.package.document.content.at(0);
      if (!body || body.type !== "paragraph") panic("Expected loaded body");
      body.content.push({
        type: "run",
        content: [
          {
            type: kind === "footnote" ? "footnoteRef" : "endnoteRef",
            id: 2,
            customMarkFollows: true,
          },
          { type: "text", text: "†" },
        ],
      });
      const content = [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "Custom note" }] }],
        },
      ] satisfies BlockContent[];
      if (kind === "footnote")
        document.package.footnotes = [
          ...(document.package.footnotes ?? []),
          { type: "footnote", id: 2, content },
        ];
      else
        document.package.endnotes = [
          ...(document.package.endnotes ?? []),
          { type: "endnote", id: 2, content },
        ];
      const reopened = await parseDocx(await createDocx(document), { preloadFonts: false });
      const notes = kind === "footnote" ? reopened.package.footnotes : reopened.package.endnotes;
      const custom = notes?.find(({ id }) => id === 2) ?? panic("Expected custom note");
      const paragraph = custom.content.at(0);
      if (!paragraph || paragraph.type !== "paragraph") panic("Expected custom note paragraph");
      expect(paragraph.content).toMatchObject([
        { type: "run", content: [{ type: "text", text: "Custom note" }] },
      ]);
      expect(
        paragraph.content
          .flatMap((item) => (item.type === "run" ? item.content : []))
          .filter((item) => item.type === "noteMarker"),
      ).toEqual([]);
      const automatic = notes?.find(({ id }) => id === 1)?.content.at(0);
      if (source === "existing note part") {
        if (!automatic || automatic.type !== "paragraph") panic("Expected automatic note");
        expect(
          automatic.content
            .flatMap((item) => (item.type === "run" ? item.content : []))
            .filter((item) => item.type === "noteMarker"),
        ).toEqual([{ type: "noteMarker", kind }]);
      }
      const bodyParagraph = reopened.package.document.content.at(0);
      if (!bodyParagraph || bodyParagraph.type !== "paragraph") panic("Expected reopened body");
      expect(
        bodyParagraph.content
          .flatMap((item) => (item.type === "run" ? item.content : []))
          .find(
            (item) => (item.type === "footnoteRef" || item.type === "endnoteRef") && item.id === 2,
          ),
      ).toMatchObject({ customMarkFollows: true });
    });
  }
}
