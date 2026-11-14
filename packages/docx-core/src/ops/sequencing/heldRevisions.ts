/** The tracked changes a piece of content holds, by revision id. */
import { IDENTITY_SPACES, identityKeysIn } from "../ids";

const REVISION_KEY_PREFIX = `${IDENTITY_SPACES.REVISION}:`;

/** The ids of the revisions held in `content`, as the revisions part keys them. */
export const heldRevisionIds = (content: unknown): string[] =>
  identityKeysIn(content).flatMap((key) =>
    key.startsWith(REVISION_KEY_PREFIX)
      ? [key.slice(REVISION_KEY_PREFIX.length)]
      : [],
  );
