import type { DocxConformanceClass } from "../../types/document";
import { parseXmlDocument } from "../xmlParser";
import { captureSourceProfileXml } from "../verbatimCapture";
import { serializePartElement } from "./partNamespaces";
import { panic } from "better-result";

/** A replacement carries its own scope, including when the source aliases `w`. */
export const sourceScopedBlocks = (
  xml: string,
  options: { conformance: DocxConformanceClass | undefined; bindings: ReadonlyMap<string, string> },
): string => {
  const wrapped = serializePartElement({
    partPath: "word/document.xml",
    rootName: "w:body",
    baselinePrefixes: ["w"],
    sourceBindings: options.bindings,
    body: xml,
  });
  const root = parseXmlDocument(wrapped);
  if (!root) panic("Generated source block markup has no root.");
  const namespaces = Object.entries(root.attributes ?? {}).filter(
    ([name]) => name.startsWith("xmlns:") || name === "mc:Ignorable",
  );
  const children: string[] = [];
  for (const child of root.elements ?? []) {
    if (child.type !== "element") continue;
    child.attributes = { ...Object.fromEntries(namespaces), ...child.attributes };
    children.push(captureSourceProfileXml(child, options.conformance));
  }
  return children.join("");
};
