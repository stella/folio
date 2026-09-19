// Forbid `btoa` in folio source: bytes become base64 through one owner,
// `packages/core/src/utils/base64.ts`.
//
// `btoa` takes a binary string, not bytes, so every call site first has to
// invent one, and the inventions are not equivalent. Decoding with
// `TextDecoder("latin1")` looks like the direct spelling but is not: the
// Encoding standard maps the `latin1` label to windows-1252, so bytes
// 0x80-0x9F come back as characters outside Latin-1 and `btoa` throws
// `InvalidCharacterError` on them. Image bytes contain those values, so that
// spelling fails on ordinary input, in browsers only, where `Buffer` is
// absent and no test runs. `String.fromCharCode(...chunk)` is correct but
// spreads a chunk of arguments per call and builds an intermediate string as
// large as the input.
//
// The owner asks none of this: it uses the runtime's own
// `Uint8Array.prototype.toBase64` when there is one and a table encoder when
// there is not, so it contains no `btoa` and needs no exemption from this rule.
//
// Flagged:
//   btoa(new TextDecoder("latin1").decode(bytes))
//   btoa(chunks.join(""))
//   globalThis.btoa(binary)
//
// Not flagged:
//   atob(encoded)     (decoding takes a base64 string; it has no binary-string trap)

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  report: (descriptor: { node: unknown; messageId: "handRolledBase64" }) => void;
};

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

// `btoa(...)`, `window.btoa(...)`, `globalThis.btoa(...)`: the callee is either
// the bare identifier or a member access whose property names it.
const callsBtoa = (callee: unknown): boolean => {
  if (!isAstNode(callee)) {
    return false;
  }
  if (callee.type === "Identifier") {
    return callee["name"] === "btoa";
  }
  if (callee.type !== "MemberExpression" || callee["computed"] === true) {
    return false;
  }
  const property = callee["property"];
  return isAstNode(property) && property.type === "Identifier" && property["name"] === "btoa";
};

export default {
  meta: { name: "folio-base64" },
  rules: {
    "no-hand-rolled-base64": {
      meta: {
        type: "problem",
        messages: {
          handRolledBase64:
            "`btoa` takes a binary string, and every way of building one from " +
            'bytes is a trap: `TextDecoder("latin1")` is the windows-1252 ' +
            "decoder by specification, so bytes 0x80-0x9F produce characters " +
            "`btoa` throws on, and `String.fromCharCode(...chunk)` spreads a " +
            "chunk per call for an intermediate string the size of the input. " +
            "Use `bytesToBase64` or `bytesToDataUrl` from " +
            "`packages/core/src/utils/base64.ts`, which encode bytes directly.",
        },
      },
      create(context: RuleContext) {
        return {
          CallExpression: (node: unknown) => {
            if (isAstNode(node) && callsBtoa(node.callee)) {
              context.report({ node, messageId: "handRolledBase64" });
            }
          },
        };
      },
    },
  },
};
