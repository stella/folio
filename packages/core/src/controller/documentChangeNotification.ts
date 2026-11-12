/**
 * How long a host waits after the last edit before it builds the document for
 * its change notification. Edits closer together than this share one build.
 */
export const DOCUMENT_CHANGE_NOTIFY_DELAY_MS = 250;
