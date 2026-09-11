import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type * as Y from "yjs";

import {
  ParagraphPropertySourceContract,
  ParagraphPropertySourceValidationError,
  type ParagraphPropertySourceAttribute,
  PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR,
  readProseDocumentParagraphPropertySourceContract,
} from "../docx/paragraphPropertySource";
import {
  readParagraphPropertyState,
  readPersistableParagraphPropertyState,
} from "./paragraphPropertyState";

const FOLIO_YJS_METADATA_MAP_NAME = "folio:document-metadata";
const PARAGRAPH_SOURCE_CONTRACT_KEY = "paragraphSourceContract";

export const proseDocumentParagraphSourceContract = (
  document: PMNode,
): ParagraphPropertySourceAttribute<ParagraphPropertySourceContract> =>
  readProseDocumentParagraphPropertySourceContract(document);

/** Reject internal-only paragraph state before it reaches collaboration storage. */
export const assertPersistableParagraphPropertyStates = (document: PMNode): void => {
  document.descendants((node) => {
    if (node.type.name !== "paragraph") {
      return true;
    }
    const persistable = readPersistableParagraphPropertyState(
      node.attrs["_paragraphPropertyState"],
    );
    if (persistable.status === "valid") {
      return false;
    }
    const internal = readParagraphPropertyState(node.attrs["_paragraphPropertyState"]);
    if (internal.status === "valid") {
      if (internal.value.type !== "transient-template") {
        panic("Persistable paragraph-property state failed persistence validation");
      }
      throw new ParagraphPropertySourceValidationError({
        code: "transient_state",
        message: "Transient paragraph-property templates cannot cross collaboration boundaries.",
      });
    }
    throw new ParagraphPropertySourceValidationError({
      code: "invalid_state",
      message: "Collaboration requires valid persistable paragraph-property state.",
      ...(internal.status === "invalid" ? { token: internal.raw } : {}),
    });
  });
};

export const writeYjsParagraphSourceContract = (ydoc: Y.Doc, document: PMNode): void => {
  assertPersistableParagraphPropertyStates(document);
  const contract = proseDocumentParagraphSourceContract(document);
  if (contract.status !== "valid") {
    panic("Cannot seed collaboration without a paragraph-property source contract");
  }
  ydoc
    .getMap(FOLIO_YJS_METADATA_MAP_NAME)
    .set(PARAGRAPH_SOURCE_CONTRACT_KEY, contract.value.serialized);
};

export const readYjsParagraphSourceContract = (
  ydoc: Y.Doc,
): ParagraphPropertySourceAttribute<ParagraphPropertySourceContract> => {
  const contract = ydoc.getMap(FOLIO_YJS_METADATA_MAP_NAME).get(PARAGRAPH_SOURCE_CONTRACT_KEY);
  return ParagraphPropertySourceContract.read(contract);
};

export const withParagraphSourceContract = (
  document: PMNode,
  contract: ParagraphPropertySourceContract,
): PMNode => {
  if (document.type.name !== "doc") {
    panic("A paragraph-property source contract can only attach to a document node");
  }
  assertPersistableParagraphPropertyStates(document);
  return document.type.create(
    { ...document.attrs, [PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR]: contract.serialized },
    document.content,
    document.marks,
  );
};
