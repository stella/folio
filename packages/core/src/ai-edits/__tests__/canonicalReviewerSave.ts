import { serializeCanonicalSave } from "../../docx/canonicalSave";
import type { CanonicalSaveSnapshot } from "../../types/canonicalSave";
import type { FolioDocxReviewer } from "../headless";

type SessionHolder = {
  session: { captureSaveSnapshot: () => CanonicalSaveSnapshot };
};

const holdsSession = (engine: unknown): engine is SessionHolder =>
  typeof engine === "object" &&
  engine !== null &&
  "session" in engine &&
  typeof engine.session === "object" &&
  engine.session !== null &&
  "captureSaveSnapshot" in engine.session &&
  typeof engine.session.captureSaveSnapshot === "function";

/**
 * The bytes a canonical reviewer's session saves, through the canonical save
 * the editor uses: one committed snapshot of the session's model.
 */
export const canonicalReviewerSave = async (
  reviewer: FolioDocxReviewer,
): Promise<ArrayBuffer> => {
  const engine: unknown = Reflect.get(reviewer, "engine");
  if (!holdsSession(engine))
    throw new Error("The reviewer does not edit through a canonical session.");
  const saved = await serializeCanonicalSave({
    snapshot: engine.session.captureSaveSnapshot(),
  });
  return saved.buffer;
};
