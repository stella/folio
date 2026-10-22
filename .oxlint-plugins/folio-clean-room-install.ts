type AstNode = Record<string, unknown> & { type: string };
type RuleContext = {
  getFilename: () => string;
  report: (descriptor: { node: unknown; messageId: "installOwner" }) => void;
};

const isNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof value.type === "string";

const isInstallArgument = (node: unknown) =>
  isNode(node) && (node["value"] === "install" || node["value"] === "add");

const isBunExecutable = (node: unknown) => {
  if (!isNode(node)) return false;
  if (typeof node["value"] === "string")
    return /(?:^|\/)bun$/u.test(node["value"]);
  // Scripts use process.execPath when spawning the current Bun runtime.
  const property = node["property"];
  return (
    node.type === "MemberExpression" &&
    isNode(property) &&
    (property["name"] === "execPath" || property["value"] === "execPath")
  );
};

const containsInstallCommand = (value: unknown) =>
  typeof value === "string" &&
  /(?:^|[\s;&|])(?:[^\s;&|]*\/)?bun\s+(?:--[^\s]+\s+)*(?:install|add)(?=$|[\s;`])/u.test(
    value,
  );

export default {
  meta: { name: "folio-clean-room-install" },
  rules: {
    "install-owner": {
      meta: {
        type: "problem",
        messages: {
          installOwner:
            "Use installCleanRoomDependencies from scripts/clean-room-install.ts; it writes the repository release-age policy before installing.",
        },
      },
      create(context: RuleContext) {
        if (
          context
            .getFilename()
            .replaceAll("\\", "/")
            .endsWith("/scripts/clean-room-install.ts")
        ) {
          return {};
        }
        const report = (node: unknown) =>
          context.report({ node, messageId: "installOwner" });
        return {
          Literal: (node: AstNode) => {
            const parent = node["parent"];
            if (!isNode(parent) || parent.type !== "CallExpression") return;
            const callee = parent["callee"];
            if (!isNode(callee)) return;
            const property = callee["property"];
            const name =
              callee.type === "Identifier"
                ? callee["name"]
                : isNode(property)
                  ? property["name"]
                  : undefined;
            if (
              typeof name === "string" &&
              /^(?:exec|execSync|execFile|execFileSync|spawn|spawnSync|\$)$/u.test(
                name,
              ) &&
              containsInstallCommand(node["value"])
            )
              report(node);
          },
          TemplateElement: (node: AstNode) => {
            const template = node["parent"];
            if (!isNode(template)) return;
            const parent = template["parent"];
            if (!isNode(parent) || parent.type !== "TaggedTemplateExpression")
              return;
            const value = node["value"];
            if (
              typeof value === "object" &&
              value !== null &&
              "raw" in value &&
              containsInstallCommand(value.raw)
            ) {
              report(node);
            }
          },
          ArrayExpression: (node: AstNode) => {
            const elements = node["elements"];
            if (
              Array.isArray(elements) &&
              isBunExecutable(elements.at(0)) &&
              elements.slice(1).some(isInstallArgument)
            ) {
              report(node);
            }
          },
          CallExpression: (node: AstNode) => {
            const args = node["arguments"];
            if (!Array.isArray(args) || !isBunExecutable(args.at(0))) return;
            const argumentsArray = args.at(1);
            if (
              !isNode(argumentsArray) ||
              argumentsArray.type !== "ArrayExpression"
            )
              return;
            const elements = argumentsArray["elements"];
            if (Array.isArray(elements) && elements.some(isInstallArgument))
              report(node);
          },
        };
      },
    },
  },
};
