import type { CollaborationModules } from "./hiddenEditorManager";

let collaborationModulesPromise: Promise<CollaborationModules> | null = null;

/** Lazily load and cache the optional collaboration runtime shared by adapters. */
export const loadCollaborationModules = (): Promise<CollaborationModules> => {
  collaborationModulesPromise ??= Promise.all([import("y-prosemirror"), import("yjs")])
    .then(([yProseMirror, yjs]) => ({ yProseMirror, yjs }))
    .catch((error: unknown) => {
      collaborationModulesPromise = null;
      throw error;
    });

  return collaborationModulesPromise;
};
