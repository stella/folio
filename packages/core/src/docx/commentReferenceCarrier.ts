import { findChild, getAttribute, type XmlElement } from "./xmlParser";

export const getCommentReferenceId = (
  runElement: XmlElement,
): number | null => {
  const commentReference = findChild(runElement, "w", "commentReference");
  if (!commentReference) return null;
  const id = Number.parseInt(
    getAttribute(commentReference, "w", "id") ?? "",
    10,
  );
  return Number.isFinite(id) ? id : null;
};
