import { descendantTextNodes } from "../../packages/core/src/layout-bridge/dom/textStreamDom";

const span = document.createElement("span");
const text = descendantTextNodes(span).at(0);
const property = "firstChild";
const unrelatedComputedProperty = span[property];

export { text, unrelatedComputedProperty };
