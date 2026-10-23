type TextAdvanceGap = { readonly offset: number; readonly advancePx: number };

export const applyTextAdvanceGaps = (
  host: HTMLElement,
  gaps: readonly TextAdvanceGap[],
): void => {
  host.style.textAutospace = "no-autospace";
  if (gaps.length === 0) return;
  const doc = host.ownerDocument;
  const walker = doc.createTreeWalker(host, 4);
  const nodes: Node[] = [];
  let node = walker.nextNode();
  while (node !== null) {
    nodes.push(node);
    node = walker.nextNode();
  }
  let sourceOffset = 0;
  let gapIndex = 0;
  for (const textNode of nodes) {
    const text = textNode.nodeValue ?? "";
    const end = sourceOffset + text.length;
    const fragment = doc.createDocumentFragment();
    let consumed = 0;
    while (gapIndex < gaps.length) {
      const gap = gaps[gapIndex];
      if (gap === undefined || gap.offset > end) break;
      gapIndex += 1;
      const boundary = gap.offset - sourceOffset;
      if (boundary <= consumed) continue;
      let start = boundary;
      while (start > consumed) {
        const trailing = text.charCodeAt(start - 1);
        const previous =
          trailing >= 0xdc00 && trailing <= 0xdfff ? start - 2 : start - 1;
        const character = text.slice(Math.max(consumed, previous), start);
        start = Math.max(consumed, previous);
        if (!/^\p{M}$/u.test(character)) break;
      }
      fragment.append(doc.createTextNode(text.slice(consumed, start)));
      const advance = doc.createElement("span");
      advance.style.letterSpacing = `${gap.advancePx}px`;
      advance.textContent = text.slice(start, boundary);
      fragment.append(advance);
      consumed = boundary;
    }
    if (consumed > 0) {
      fragment.append(doc.createTextNode(text.slice(consumed)));
      textNode.parentNode?.replaceChild(fragment, textNode);
    }
    sourceOffset = end;
  }
};
