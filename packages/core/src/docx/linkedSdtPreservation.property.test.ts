import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { createEmptyDocument } from "../utils/createDocument";
import { parseParagraph } from "./paragraphParser";
import { parseXmlDocument, WORDPROCESSINGML_NAMESPACE_URIS } from "./xmlParser";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";
import { EditorState } from "prosemirror-state";
import { resolveAllChangesInHeadlessState } from "../prosemirror/commands/comments";
import { UnrepresentableLinkedSdtRevisionError } from "./linkedSdtPreservation";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import type { Document, ParagraphContent } from "../types/document";

const sourceDocument = ({
  form,
  type,
  prefix,
  namespace,
}: {
  form: "simple" | "structural";
  type: "ins" | "del" | "moveFrom" | "moveTo";
  prefix: string;
  namespace: string;
}) => {
  const textTag = type === "del" || type === "moveFrom" ? "delText" : "t";
  const linked = `<${prefix}:hyperlink ${prefix}:anchor="target"><${prefix}:${type} ${prefix}:id="37" ${prefix}:author="Reviewer"><${prefix}:sdt><${prefix}:sdtPr><${prefix}:alias ${prefix}:val="source-control"/></${prefix}:sdtPr><${prefix}:sdtContent><${prefix}:r><${prefix}:${textTag}>linked</${prefix}:${textTag}></${prefix}:r></${prefix}:sdtContent></${prefix}:sdt></${prefix}:${type}></${prefix}:hyperlink>`;
  const body =
    form === "simple"
      ? `<${prefix}:fldSimple ${prefix}:instr="REF target">${linked}</${prefix}:fldSimple>`
      : `<${prefix}:r><${prefix}:fldChar ${prefix}:fldCharType="begin"/></${prefix}:r><${prefix}:r><${prefix}:instrText>REF target</${prefix}:instrText></${prefix}:r><${prefix}:r><${prefix}:fldChar ${prefix}:fldCharType="separate"/></${prefix}:r>${linked}<${prefix}:r><${prefix}:fldChar ${prefix}:fldCharType="end"/></${prefix}:r>`;
  const root = parseXmlDocument(
    `<${prefix}:p xmlns:${prefix}="${namespace}">${body}</${prefix}:p>`,
  );
  if (!root) throw new TypeError("Source field fixture did not parse.");
  const source = createEmptyDocument();
  source.package.document.content = [parseParagraph(root, null, null, null)];
  return source;
};

const sourceSpan = (source: Document) => {
  const paragraph = source.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") throw new TypeError("Source field paragraph disappeared.");
  const spans: string[] = [];
  const collect = (content: readonly ParagraphContent[]) => {
    for (const item of content) {
      if (item.type === "preservedInline" && item.xml.includes("source-control"))
        spans.push(item.xml);
      else if (item.type === "hyperlink") collect(item.children);
      else if (
        item.type === "inlineWrapper" ||
        item.type === "simpleField" ||
        item.type === "inlineSdt" ||
        item.type === "insertion" ||
        item.type === "deletion" ||
        item.type === "moveFrom" ||
        item.type === "moveTo"
      )
        collect(item.content);
    }
  };
  collect(paragraph.content);
  expect(spans).toHaveLength(1);
  const span = spans.at(0);
  if (!span) throw new TypeError("Opaque SDT span disappeared.");
  return span;
};

const assertTwoSaves = async (source: Document) => {
  const original = sourceSpan(source);
  expect(original).not.toMatch(/<(?:w|q|source):(ins|del|moveFrom|moveTo)\b/u);
  const initial = toProseDoc(source);
  let revisions = 0;
  initial.descendants((node) => {
    if (node.marks.some((mark) => mark.type.name === "insertion" || mark.type.name === "deletion"))
      revisions++;
  });
  expect(revisions).toBeGreaterThan(0);
  const projected = fromProseDoc(toProseDoc(source), source);
  expect(sourceSpan(projected)).toBe(original);
  const opened = await parseDocx(await createDocx(projected), { preloadFonts: false });
  expect(sourceSpan(opened)).toBe(original);
  const projectedAgain = fromProseDoc(toProseDoc(opened), opened);
  expect(sourceSpan(projectedAgain)).toBe(original);
  const reopened = await parseDocx(await createDocx(projectedAgain), { preloadFonts: false });
  expect(sourceSpan(reopened)).toBe(original);
  for (const mode of ["accept", "reject"] as const) {
    const resolved = resolveAllChangesInHeadlessState(
      EditorState.create({ doc: toProseDoc(reopened) }),
      mode,
    );
    const resolvedSource = fromProseDoc(resolved.doc, reopened);
    const savedResolved = await parseDocx(await createDocx(resolvedSource), {
      preloadFonts: false,
    });
    const paragraph = savedResolved.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") throw new TypeError("Resolved paragraph disappeared.");
    // A clean review result contains no pending wrapper in its serialized source.
    expect(serializeParagraph(paragraph)).not.toMatch(/<w:(ins|del|moveFrom|moveTo)\b/u);
    expect(serializeParagraph(paragraph)).not.toMatch(/<(?:w|q|source):delText\b/u);
  }
};

for (const type of ["ins", "del"] as const) {
  test(`linked SDT ${type} field opens and preserves its span through two saves`, async () => {
    for (const form of ["simple", "structural"] as const) {
      await assertTwoSaves(
        sourceDocument({
          form,
          type,
          prefix: "w",
          namespace: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        }),
      );
    }
  });
}

test(
  "generated linked SDT field spans survive source parsing, editor projection and two saves",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom("simple", "structural"),
        fc.constantFrom("ins", "del", "moveFrom", "moveTo"),
        fc.constantFrom("w", "q", "source"),
        fc.constantFrom(...WORDPROCESSINGML_NAMESPACE_URIS),
        async (form, type, prefix, namespace) => {
          await assertTwoSaves(sourceDocument({ form, type, prefix, namespace }));
        },
      ),
      { numRuns: 30 },
    );
  },
  propertyTestTimeout(20_000),
);

for (const type of ["ins", "del"] as const) {
  test(`a nested ${type} inside an opaque linked SDT is refused visibly`, () => {
    const root = parseXmlDocument(
      `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:hyperlink w:anchor="target"><w:sdt><w:sdtPr/><w:sdtContent><w:${type} w:id="38" w:author="Other"><w:r><w:t>pending</w:t></w:r></w:${type}></w:sdtContent></w:sdt></w:hyperlink></w:p>`,
    );
    if (!root) throw new TypeError("Fixture did not parse.");
    expect(() => parseParagraph(root, null, null, null)).toThrow(
      UnrepresentableLinkedSdtRevisionError,
    );
  });
}

for (const type of ["insertion", "deletion"] as const) {
  test(`${type} owns an SDT hyperlink without lifting through the control`, async () => {
    const source = createEmptyDocument();
    source.package.document.content = [
      {
        type: "paragraph",
        content: [
          {
            type,
            info: { id: 37, author: "Reviewer" },
            resolutionJoins: { before: 1, after: 1, remove: 1 },
            content: [
              {
                type: "inlineSdt",
                properties: { sdtType: "richText", alias: "boundary" },
                content: [
                  {
                    type: "hyperlink",
                    anchor: "target",
                    children: [{ type: "run", content: [{ type: "text", text: "linked" }] }],
                  },
                ],
              },
            ],
          },
        ],
      },
    ];
    const paragraph = source.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") throw new TypeError("Fixture paragraph missing.");
    const xml = serializeParagraph(paragraph);
    const tag = type === "insertion" ? "ins" : "del";
    expect(xml).toMatch(new RegExp(`<w:${tag}[^>]*>.*<w:sdt>.*<w:sdtContent><w:hyperlink`, "u"));
    expect(xml.match(/folio:resolutionJoins=/gu)).toHaveLength(1);
    const reopened = await parseDocx(await createDocx(fromProseDoc(toProseDoc(source), source)), {
      preloadFonts: false,
    });
    const reopenedParagraph = reopened.package.document.content.at(0);
    if (reopenedParagraph?.type !== "paragraph") throw new TypeError("Saved paragraph missing.");
    expect(serializeParagraph(reopenedParagraph)).toContain('<w:hyperlink w:anchor="target">');
    for (const mode of ["accept", "reject"] as const) {
      const resolved = resolveAllChangesInHeadlessState(
        EditorState.create({ doc: toProseDoc(reopened) }),
        mode,
      );
      const final = await parseDocx(await createDocx(fromProseDoc(resolved.doc, reopened)), {
        preloadFonts: false,
      });
      const expected = (mode === "accept") === (type === "insertion") ? "linked" : "";
      expect(toProseDoc(final).textContent).toBe(expected);
    }
  });
}
