const UNSUPPORTED_PROJECTION_ERROR_NAME = "UnsupportedDocxToProseMirrorConversionError";

export const parseUnsupportedProjectionConsoleError = (
  type: string,
  text: string,
): string | undefined => {
  if (type !== "error") {
    return undefined;
  }
  const markerIndex = text.indexOf(UNSUPPORTED_PROJECTION_ERROR_NAME);
  if (markerIndex < 0) {
    return undefined;
  }
  return text.slice(markerIndex).split("\n", 1).at(0)?.trim();
};

type EditorReadinessProbeOptions = {
  playgroundErrorSelector: string;
  readySelector: string;
};

export const readEditorReadinessState = ({
  playgroundErrorSelector,
  readySelector,
}: EditorReadinessProbeOptions) => {
  const loadError = document.querySelector(".folio-editor-error p")?.textContent?.trim();
  if (loadError) {
    return { type: "error" as const, message: loadError };
  }
  const playgroundError = document.querySelector(playgroundErrorSelector)?.textContent?.trim();
  if (playgroundError) {
    return { type: "error" as const, message: playgroundError };
  }
  return document.querySelector(readySelector) ? { type: "ready" as const } : null;
};
