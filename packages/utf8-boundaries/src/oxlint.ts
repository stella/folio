type AstNode = { type: string } & Record<string, unknown>;
type Context = {
  filename?: string;
  getFilename?: () => string;
  options?: unknown[];
  report: (descriptor: { node: unknown; messageId: string }) => void;
};

type Rule = {
  meta: {
    type: "problem";
    messages: Record<string, string>;
    schema?: unknown[];
  };
  create: (context: Context) => Record<string, (node: unknown) => void>;
};

const isNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof value["type"] === "string";
const propertyName = (value: unknown): string | null => {
  if (!isNode(value)) return null;
  if (value.type === "Identifier")
    return typeof value["name"] === "string" ? value["name"] : null;
  if (value.type === "Literal")
    return typeof value["value"] === "string" ? value["value"] : null;
  return null;
};

const STREAM_METHODS = new Set(["on", "once", "addEventListener"]);
const STREAM_CONSTRUCTORS = new Set(["ReadableStream", "TransformStream"]);
const STREAM_HANDLER_PROPERTIES = new Set(["start", "pull", "transform"]);

const enablesStreaming = (options: unknown): boolean => {
  if (!isNode(options) || options.type !== "ObjectExpression") return false;
  const properties = options["properties"];
  if (!Array.isArray(properties)) return false;
  let stream = false;
  for (const entry of properties) {
    if (!isNode(entry)) return false;
    const key = entry["key"];
    if (
      entry.type === "SpreadElement" ||
      (entry["computed"] === true && (!isNode(key) || key.type !== "Literal"))
    ) {
      stream = false;
      continue;
    }
    if (propertyName(key) !== "stream") continue;
    const value = entry["value"];
    stream =
      isNode(value) && value.type === "Literal" && value["value"] === true;
  }
  return stream;
};

const noChunkedTextCodec: Rule = {
  meta: {
    type: "problem",
    messages: {
      decode:
        "TextDecoder.decode() inside a loop or stream handler must use { stream: true } to preserve split UTF-8 sequences.",
      decoder:
        "Create TextDecoder outside the loop so chunk decoding keeps one decoder across iterations.",
    },
  },
  create(context) {
    let loopDepth = 0;
    let streamDepth = 0;
    const streamCallbacks = new WeakSet<object>();
    const decoderBindings = new Set<string>();
    const isDecoderExpression = (value: unknown): boolean => {
      if (!isNode(value)) return false;
      if (value.type === "NewExpression") {
        const constructor = value["callee"];
        return (
          isNode(constructor) &&
          constructor.type === "Identifier" &&
          constructor["name"] === "TextDecoder"
        );
      }
      return (
        value.type === "Identifier" &&
        typeof value["name"] === "string" &&
        decoderBindings.has(value["name"])
      );
    };
    return {
      VariableDeclarator(raw) {
        if (!isNode(raw)) return;
        const id = raw["id"];
        const init = raw["init"];
        if (
          isNode(id) &&
          id.type === "Identifier" &&
          typeof id["name"] === "string" &&
          isDecoderExpression(init)
        )
          decoderBindings.add(id["name"]);
      },
      CallExpression(raw) {
        if (!isNode(raw)) return;
        const callee = raw["callee"];
        if (
          isNode(callee) &&
          callee.type === "MemberExpression" &&
          propertyName(callee["property"]) === "decode" &&
          isDecoderExpression(callee["object"])
        ) {
          if (loopDepth === 0 && streamDepth === 0) return;
          const args = raw["arguments"];
          const options = Array.isArray(args) ? args[1] : undefined;
          if (!enablesStreaming(options))
            context.report({ node: raw, messageId: "decode" });
        }
        if (
          isNode(callee) &&
          callee.type === "MemberExpression" &&
          STREAM_METHODS.has(propertyName(callee["property"]) ?? "")
        ) {
          const args = raw["arguments"];
          for (const argument of Array.isArray(args) ? args : []) {
            if (
              isNode(argument) &&
              ["ArrowFunctionExpression", "FunctionExpression"].includes(
                argument.type,
              )
            )
              streamCallbacks.add(argument);
          }
        }
      },
      NewExpression(raw) {
        if (!isNode(raw)) return;
        const callee = raw["callee"];
        const args = raw["arguments"];
        const streamOptions = Array.isArray(args) ? args[0] : undefined;
        const streamProperties = isNode(streamOptions)
          ? streamOptions["properties"]
          : undefined;
        if (
          isNode(callee) &&
          callee.type === "Identifier" &&
          STREAM_CONSTRUCTORS.has(
            typeof callee["name"] === "string" ? callee["name"] : "",
          ) &&
          isNode(streamOptions) &&
          streamOptions.type === "ObjectExpression" &&
          Array.isArray(streamProperties)
        ) {
          for (const property of streamProperties) {
            if (
              !isNode(property) ||
              !STREAM_HANDLER_PROPERTIES.has(
                propertyName(property["key"]) ?? "",
              )
            )
              continue;
            const handler = property["value"];
            if (isNode(handler) && handler.type === "FunctionExpression")
              streamCallbacks.add(handler);
          }
        }
        if (
          loopDepth > 0 &&
          isNode(callee) &&
          callee.type === "Identifier" &&
          callee["name"] === "TextDecoder"
        ) {
          context.report({ node: raw, messageId: "decoder" });
        }
      },
      "ForStatement:exit"() {
        loopDepth = Math.max(0, loopDepth - 1);
      },
      "ForInStatement:exit"() {
        loopDepth = Math.max(0, loopDepth - 1);
      },
      "ForOfStatement:exit"() {
        loopDepth = Math.max(0, loopDepth - 1);
      },
      "WhileStatement:exit"() {
        loopDepth = Math.max(0, loopDepth - 1);
      },
      "DoWhileStatement:exit"() {
        loopDepth = Math.max(0, loopDepth - 1);
      },
      ForStatement() {
        loopDepth++;
      },
      ForInStatement() {
        loopDepth++;
      },
      ForOfStatement() {
        loopDepth++;
      },
      WhileStatement() {
        loopDepth++;
      },
      DoWhileStatement() {
        loopDepth++;
      },
      FunctionExpression(node) {
        if (isNode(node) && streamCallbacks.has(node)) streamDepth++;
      },
      "FunctionExpression:exit"(node) {
        if (isNode(node) && streamCallbacks.has(node))
          streamDepth = Math.max(0, streamDepth - 1);
      },
      ArrowFunctionExpression(node) {
        if (isNode(node) && streamCallbacks.has(node)) streamDepth++;
      },
      "ArrowFunctionExpression:exit"(node) {
        if (isNode(node) && streamCallbacks.has(node))
          streamDepth = Math.max(0, streamDepth - 1);
      },
    };
  },
};

const plugin: {
  meta: {
    name: string;
  };
  rules: {
    "no-chunked-text-codec": Rule;
  };
} = {
  meta: { name: "utf8-boundaries" },
  rules: {
    "no-chunked-text-codec": noChunkedTextCodec,
  },
};

export default plugin;
