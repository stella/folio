import { describe, expect, test } from "bun:test";
import * as yProseMirror from "y-prosemirror";
import * as yjs from "yjs";

import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import {
  FOLIO_YJS_ATTR_SCHEMA_VERSION,
  FolioYjsAttrSchemaVersionError,
  proseDocumentParagraphSourceContract,
} from "../prosemirror/yjsDocumentMetadata";
import { createEmptyDocument } from "../utils/createDocument";
import { createHiddenEditorState } from "./hiddenEditorManager";

const METADATA_MAP_NAME = "folio:document-metadata";
const ATTR_SCHEMA_VERSION_KEY = "attrSchemaVersion";
const PARAGRAPH_SOURCE_CONTRACT_KEY = "paragraphSourceContract";
const PROSEMIRROR_FRAGMENT_NAME = "prosemirror";

const collaborationModules = { yProseMirror, yjs };

const parsedDocument = async () =>
  await parseDocx(await createDocx(createEmptyDocument({ initialText: "Room" })), {
    preloadFonts: false,
  });

describe("createHiddenEditorState with collaboration", () => {
  test("seeding stamps the attr-schema version", async () => {
    const document = await parsedDocument();
    const ydoc = new yjs.Doc();

    createHiddenEditorState({
      collaboration: {
        shouldSeed: true,
        yXmlFragment: ydoc.getXmlFragment(PROSEMIRROR_FRAGMENT_NAME),
      },
      collaborationModules,
      document,
    });

    expect(ydoc.getMap(METADATA_MAP_NAME).get(ATTR_SCHEMA_VERSION_KEY)).toBe(
      FOLIO_YJS_ATTR_SCHEMA_VERSION,
    );
    ydoc.destroy();
  });

  test("loads a snapshot written before the marker existed", async () => {
    const document = await parsedDocument();
    const ydoc = new yjs.Doc();
    const fragment = ydoc.getXmlFragment(PROSEMIRROR_FRAGMENT_NAME);
    const proseDocument = toProseDoc(document);
    yProseMirror.prosemirrorToYXmlFragment(proseDocument, fragment);
    // Exactly what a pre-marker build wrote: the source contract, no version.
    const contract = proseDocumentParagraphSourceContract(proseDocument);
    if (!contract) {
      throw new Error("Fixture must carry a paragraph source contract");
    }
    ydoc.getMap(METADATA_MAP_NAME).set(PARAGRAPH_SOURCE_CONTRACT_KEY, contract);
    expect(ydoc.getMap(METADATA_MAP_NAME).get(ATTR_SCHEMA_VERSION_KEY)).toBeUndefined();

    const state = createHiddenEditorState({
      collaboration: { shouldSeed: false, yXmlFragment: fragment },
      collaborationModules,
      document,
    });

    expect(state.doc.childCount).toBeGreaterThan(0);
    ydoc.destroy();
  });

  test("refuses a snapshot written by a newer attr schema", async () => {
    const document = await parsedDocument();
    const ydoc = new yjs.Doc();
    const fragment = ydoc.getXmlFragment(PROSEMIRROR_FRAGMENT_NAME);
    yProseMirror.prosemirrorToYXmlFragment(toProseDoc(document), fragment);
    ydoc.getMap(METADATA_MAP_NAME).set(ATTR_SCHEMA_VERSION_KEY, FOLIO_YJS_ATTR_SCHEMA_VERSION + 1);

    expect(() =>
      createHiddenEditorState({
        collaboration: { shouldSeed: false, yXmlFragment: fragment },
        collaborationModules,
        document,
      }),
    ).toThrow(FolioYjsAttrSchemaVersionError);
    ydoc.destroy();
  });
});
