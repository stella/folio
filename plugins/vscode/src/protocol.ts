/**
 * The messages the preview's extension side and its webview exchange. Both
 * sides check what they receive: a webview message crosses a process boundary,
 * so neither trusts the other's shape.
 */

/** Extension to webview. */
export type HostMessage =
  | { readonly type: "loading"; readonly fileName: string }
  | {
      readonly type: "document";
      readonly fileName: string;
      /** Standalone HTML of every page, fonts and images inlined. */
      readonly html: string;
      readonly pageCount: number;
    }
  | {
      readonly type: "error";
      readonly fileName: string;
      readonly message: string;
      readonly hint?: string;
    };

/** Webview to extension. */
export type WebviewMessage =
  /** The webview's script is listening; send the current state. */
  | { readonly type: "ready" }
  /** Render the file again, for the error state's retry button. */
  | { readonly type: "retry" };

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isString = (value: unknown): value is string => typeof value === "string";

export const isHostMessage = (value: unknown): value is HostMessage => {
  if (!isRecord(value) || !isString(value["fileName"])) return false;
  switch (value["type"]) {
    case "loading":
      return true;
    case "document":
      return (
        isString(value["html"]) &&
        typeof value["pageCount"] === "number" &&
        Number.isInteger(value["pageCount"]) &&
        value["pageCount"] >= 0
      );
    case "error":
      return isString(value["message"]) && (value["hint"] === undefined || isString(value["hint"]));
    default:
      return false;
  }
};

export const isWebviewMessage = (value: unknown): value is WebviewMessage =>
  isRecord(value) && (value["type"] === "ready" || value["type"] === "retry");

/** "1 page", "12 pages". */
export const pageCountLabel = (pageCount: number): string =>
  `${String(pageCount)} page${pageCount === 1 ? "" : "s"}`;
