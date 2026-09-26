import { paraIdAttribute, resolveParagraphIdentities } from "../src/docx/paraIdAttribute";
import {
  findWordprocessingChild,
  getChildElements,
  getLocalName,
  getNamespaceUri,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "../src/docx/xmlParser";

const bodyParaIds = (xml: string) =>
  getChildElements(findWordprocessingChild(parseXmlDocument(xml), "body"))
    .filter(
      (element) =>
        getLocalName(element.name) === "p" &&
        WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? ""),
    )
    .map(paraIdAttribute);

/** Resolve source/saved XML ordinals through the model's serialization, never PM node counts. */
export const censusParagraphOrdinals = (
  xml: string,
  serializedModel: string,
): Map<string, number> => {
  const sourceIds = bodyParaIds(xml);
  const modelIds = bodyParaIds(serializedModel);
  const { identities, ordinalsAligned } = resolveParagraphIdentities({
    sourceParaIds: sourceIds,
    serializedParaIds: modelIds,
  });
  const ordinals = new Map<string, number>();
  for (const identity of identities) {
    if (
      identity.type === "anonymous" ||
      modelIds.indexOf(identity.paraId) !== modelIds.lastIndexOf(identity.paraId)
    )
      continue;
    if (identity.type === "authored") {
      const ordinal = sourceIds.indexOf(identity.paraId);
      if (ordinal === sourceIds.lastIndexOf(identity.paraId))
        ordinals.set(identity.paraId, ordinal);
    } else if (ordinalsAligned && sourceIds[identity.ordinal] === undefined) {
      ordinals.set(identity.paraId, identity.ordinal);
    }
  }
  return ordinals;
};
