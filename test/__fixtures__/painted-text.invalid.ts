const span = document.createElement("span");
const text = span.firstChild;
const computedText = span["firstChild"];
const textNodeType = Node["TEXT_NODE"];
const showText = NodeFilter["SHOW_TEXT"];

export { computedText, showText, text, textNodeType };
