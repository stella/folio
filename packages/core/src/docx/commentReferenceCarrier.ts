import {
  findWordprocessingChild,
  getAttributeByNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

export const getCommentReferenceId = (
  runElement: XmlElement,
): number | null => {
  const commentReference = findWordprocessingChild(
    runElement,
    "commentReference",
  );
  if (!commentReference) return null;
  const id = Number.parseInt(
    getAttributeByNamespaceUri(
      commentReference,
      WORDPROCESSINGML_NAMESPACE_URIS,
      "id",
    ) ?? "",
    10,
  );
  return Number.isFinite(id) ? id : null;
};
