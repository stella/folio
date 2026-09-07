/**
 * A save never introduces a namespace defect into a package part.
 *
 * The serializers derive each rebuilt part's `xmlns:*` from the markup they
 * assembled, so this gate checks the property that derivation exists for: every
 * prefix in a saved part resolves to a namespace URI, and `mc:Ignorable` names
 * only prefixes the part declares. Defects already present in a fixture's own
 * part are subtracted, so a malformed source stays the source's problem while
 * anything folio adds fails here.
 *
 * The whole corpus runs through `test.each`, so a fixture added later is
 * covered with no edit here. Fixture provenance and licensing: see
 * `__fixtures__/corpus/PROVENANCE.md`.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";

import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Document } from "../../types/document";
import { parseDocx } from "../parser";
import { repackDocx } from "../rezip";
import { serializeDocument } from "../serializer/documentSerializer";
import { OOXML_NAMESPACES } from "../serializer/partNamespaces";
import {
  getNamespacePrefix,
  parseXmlDocument,
  type XmlElement,
  type XmlNamespaceScope,
} from "../xmlParser";

const CORPUS_DIR = path.join(import.meta.dir, "__fixtures__", "corpus");
const VISUAL_FIXTURES_DIR = path.resolve(import.meta.dir, "../../../../../tests/visual/fixtures");

const CORPUS_FIXTURES: string[] = readdirSync(CORPUS_DIR)
  .filter((name) => name.endsWith(".docx"))
  .sort()
  .map((name) => path.join(CORPUS_DIR, name));

const VISUAL_FIXTURES = ["sample.docx", "docx-editor-demo.docx", "podily-bps.docx"].map((name) =>
  path.join(VISUAL_FIXTURES_DIR, name),
);

const FIXTURES = [...CORPUS_FIXTURES, ...VISUAL_FIXTURES];

const readFixture = (file: string): ArrayBuffer => {
  const bytes = readFileSync(file);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};

const resolvePrefix = (scope: XmlNamespaceScope | undefined, prefix: string): boolean => {
  for (let current = scope; current; current = current.parent) {
    if (current.bindings.has(prefix)) {
      return true;
    }
  }
  return false;
};

/**
 * Namespace defects in one part, as stable descriptions so a saved part's
 * defects can be subtracted from its source's.
 */
const namespaceDefects = (xml: string): Set<string> => {
  const defects = new Set<string>();
  const root = parseXmlDocument(xml);
  if (root === null) {
    defects.add("unparseable");
    return defects;
  }

  const visit = (element: XmlElement): void => {
    const { namespaceScope } = element;
    const elementPrefix = getNamespacePrefix(element.name ?? "");
    if (elementPrefix !== null && !resolvePrefix(namespaceScope, elementPrefix)) {
      defects.add(`unbound element prefix ${elementPrefix}`);
    }
    for (const attribute of Object.keys(element.attributes ?? {})) {
      const attributePrefix = getNamespacePrefix(attribute);
      if (
        attributePrefix !== null &&
        attributePrefix !== "xmlns" &&
        attributePrefix !== "xml" &&
        !resolvePrefix(namespaceScope, attributePrefix)
      ) {
        defects.add(`unbound attribute prefix ${attributePrefix}`);
      }
    }
    for (const child of element.elements ?? []) {
      if (child.type === "element") {
        visit(child);
      }
    }
  };
  visit(root);

  const ignorable = root.attributes?.["mc:Ignorable"];
  if (typeof ignorable === "string") {
    for (const prefix of ignorable.split(" ").filter(Boolean)) {
      if (!resolvePrefix(root.namespaceScope, prefix)) {
        defects.add(`mc:Ignorable names undeclared prefix ${prefix}`);
      }
    }
  }
  return defects;
};

type PartDefects = Map<string, Set<string>>;

const packageDefects = async (buffer: ArrayBuffer): Promise<PartDefects> => {
  const zip = await JSZip.loadAsync(buffer);
  const defects: PartDefects = new Map();
  for (const [name, file] of Object.entries(zip.files)) {
    if (file.dir || !name.toLowerCase().endsWith(".xml")) {
      continue;
    }
    defects.set(name, namespaceDefects(await file.async("text")));
  }
  return defects;
};

const roundTrip = async (original: Document): Promise<ArrayBuffer> =>
  repackDocx(fromProseDoc(toProseDoc(original), original), { updateModifiedDate: false });

describe("saved package namespace declarations", () => {
  test.each(FIXTURES.map((file) => [path.basename(file), file] as const))(
    "%s introduces no unbound prefix",
    async (_name, file) => {
      const original = readFixture(file);
      const saved = await roundTrip(await parseDocx(original));

      const before = await packageDefects(original);
      const after = await packageDefects(saved);

      const introduced = [...after].flatMap(([part, defects]) => {
        const known = before.get(part) ?? new Set<string>();
        return [...defects].filter((defect) => !known.has(defect)).map((d) => `${part}: ${d}`);
      });
      expect(introduced).toEqual([]);
    },
    30_000,
  );

  test("a text box replayed verbatim keeps its producer's prefix bound", () => {
    // The parser preserves an unmodeled drawing as source XML. A text box
    // written under ISO conformance reaches the serializer as `wne:txbxContent`
    // whose only binding was on the source document's root.
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [
                    {
                      type: "drawing",
                      image: {
                        type: "image",
                        rId: "rId1",
                        size: { width: 9525, height: 9525 },
                        wrap: { type: "inline" },
                      },
                      rawXml:
                        "<w:drawing><wp:inline><wp:txbx><wne:txbxContent><w:p/></wne:txbxContent></wp:txbx></wp:inline></w:drawing>",
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    };

    const xml = serializeDocument(document);
    expect(xml).toContain(`xmlns:wne="${OOXML_NAMESPACES.wne.uri}"`);
    expect([...namespaceDefects(xml)]).toEqual([]);
  });
});
