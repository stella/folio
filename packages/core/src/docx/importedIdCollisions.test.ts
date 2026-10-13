import { expect, test } from "bun:test";
import { createParseWarningCollector } from "./parseContext";
import { PARA_ID_NAMESPACE_URIS } from "./paraIdAttribute";
import {
  diagnoseImportedIdCollisions,
  IMPORTED_ID_COLLISION_REGISTRY,
} from "./importedIdCollisions";

test("every registered definition site diagnoses repetitions and ignores part boundaries", () => {
  const exercised = new Set<string>();
  for (const [kind, rule] of Object.entries(IMPORTED_ID_COLLISION_REGISTRY)) {
    for (const site of rule.sites) {
      for (const element of site.elements) {
        const value = rule.encoding === "hex" ? "00abcdef" : "7";
        const namespace =
          rule.attributeNamespace === null ? "" : ` xmlns:y="${rule.attributeNamespace}"`;
        const attribute = rule.attributeNamespace === null ? rule.attribute : `y:${rule.attribute}`;
        const node = `<x:${element} xmlns:x="${site.namespace}"${namespace} ${attribute}="${value}"/>`;
        const collector = createParseWarningCollector();
        diagnoseImportedIdCollisions(
          new Map([
            ["one.xml", `<root>${node}${node}${node}</root>`],
            ["two.xml", `<root>${node}</root>`],
          ]),
          collector.context,
        );
        expect(collector.warnings()).toHaveLength(rule.scope === "package" ? 3 : 2);
        expect(
          collector
            .warnings()
            .every(
              ({ detail, location }) =>
                detail?.startsWith(`${kind} identity`) &&
                (location.part === "one.xml" || rule.scope === "package"),
            ),
        ).toBe(true);
        exercised.add(kind);
      }
    }
  }
  expect([...exercised].toSorted()).toEqual(Object.keys(IMPORTED_ID_COLLISION_REGISTRY).toSorted());
});

test("paragraph identity diagnostics share every namespace accepted by the paragraph reader", () => {
  for (const namespace of [
    ...PARA_ID_NAMESPACE_URIS,
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main",
  ]) {
    for (const attribute of ["paraId", "textId"]) {
      const collector = createParseWarningCollector();
      const node = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:id="${namespace}" id:${attribute}="15AB23CD"/>`;
      diagnoseImportedIdCollisions(
        new Map([["word/document.xml", `<root>${node}${node}</root>`]]),
        collector.context,
      );
      expect(collector.warnings()).toHaveLength(1);
      expect(collector.warnings().at(0)?.identityKind).toBe(attribute);
    }
  }
});

test("equivalent AlternateContent branches do not claim the same identity twice", () => {
  const collector = createParseWarningCollector();
  const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body><mc:AlternateContent><mc:Choice Requires="w"><w:p><w:bookmarkStart w:id="37" w:name="target"/><w:bookmarkEnd w:id="37"/></w:p></mc:Choice><mc:Fallback><w:p><w:bookmarkStart w:id="37" w:name="target"/><w:bookmarkEnd w:id="37"/></w:p></mc:Fallback></mc:AlternateContent></w:body></w:document>`;
  diagnoseImportedIdCollisions(new Map([["word/document.xml", xml]]), collector.context);
  expect(collector.warnings()).toEqual([]);
});

test("references and cNvPr do not claim definition identities", () => {
  const collector = createParseWarningCollector();
  const xml = `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:bookmarkStart w:id="7"/><w:bookmarkEnd w:id="7"/><w:footnoteReference w:id="7"/><wp:docPr id="7"/><pic:cNvPr id="7"/></root>`;
  diagnoseImportedIdCollisions(new Map([["word/document.xml", xml]]), collector.context);
  expect(collector.warnings()).toEqual([]);
});

test("comment metadata duplicate joins are diagnosed without counting wrapper and paragraph aliases twice", () => {
  const collector = createParseWarningCollector();
  const extended = `<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="15AB23CD" w15:done="0"/><w15:commentEx w15:paraId="15AB23CD" w15:done="1" w15:paraIdParent="5ABC6789"/></w15:commentsEx>`;
  const extensible = `<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"><w16cex:comment w16cex:paraId="15AB23CD" w16cex:dateUtc="2024-01-01T00:00:00Z"/><w16cex:comment w16cex:paraId="15AB23CD" w16cex:dateUtc="2024-01-02T00:00:00Z"/></w16cex:commentsExtensible>`;
  const comments = `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="23" w14:paraId="15AB23CD"><w:p w14:paraId="15AB23CD"><w:r><w:t>Comment</w:t></w:r></w:p></w:comment></w:comments>`;
  diagnoseImportedIdCollisions(
    new Map([
      ["word/commentsExtended.xml", extended],
      ["word/commentsExtensible.xml", extensible],
      ["word/comments.xml", comments],
    ]),
    collector.context,
  );
  const warnings = collector.warnings();
  expect(warnings).toHaveLength(2);
  expect(warnings.every(({ identityKind }) => identityKind === "paraId")).toBe(true);
  expect(warnings.map(({ location }) => location.part)).toEqual([
    "word/commentsExtended.xml",
    "word/commentsExtensible.xml",
  ]);
});

test("Strict namespace aliases and decimal spellings identify the same source definition", () => {
  const collector = createParseWarningCollector();
  const xml = `<root xmlns:a="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:b="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:foreign="urn:foreign"><a:num a:numId="007"/><b:num b:numId="7"/><b:num foreign:numId="7"/></root>`;
  diagnoseImportedIdCollisions(new Map([["word/numbering.xml", xml]]), collector.context);
  const warnings = collector.warnings();
  expect(warnings).toHaveLength(1);
  expect(warnings.at(0)?.location).toEqual({
    part: "word/numbering.xml",
    element: "b:num",
    at: `offset ${xml.indexOf('="7"') + 2}`,
  });
  expect(warnings.at(0)?.value).toBe("7");
  expect(warnings.at(0)?.detail).toBe(
    `numId identity repeats the definition at offset ${xml.indexOf('="007"') + 2} in word/numbering.xml.`,
  );
});
