/**
 * A minimal `Document` for painter tests.
 *
 * The painters build DOM through a handful of operations, so a stub covering
 * those is enough to assert structure, classes and geometry without a browser.
 * It lives here rather than in each spec because two renderers now paint pages
 * and both must be exercised against the same stub: a per-file copy is how the
 * two would come to be tested against subtly different DOM semantics.
 */

/**
 * A style object that also answers `setProperty`, which the painters use to
 * set CSS custom properties. A plain record silently lacks it, and the failure
 * surfaces as a `TypeError` deep inside a renderer rather than as a missing
 * stub, so it is defined here once.
 */
export type FakeStyle = Record<string, string> & {
  setProperty: (name: string, value: string) => void;
  removeProperty: (name: string) => void;
};

const createFakeStyle = (): FakeStyle => {
  const style = {} as FakeStyle;
  Object.defineProperty(style, "setProperty", {
    enumerable: false,
    value: (name: string, value: string) => {
      style[name] = value;
    },
  });
  Object.defineProperty(style, "removeProperty", {
    enumerable: false,
    value: (name: string) => {
      // oxlint-disable-next-line no-dynamic-delete -- mirrors CSSStyleDeclaration
      delete style[name];
    },
  });
  return style;
};

export class FakeElement {
  className = "";
  dataset: Record<string, string> = {};
  style: FakeStyle = createFakeStyle();
  children: FakeElement[] = [];
  parent: FakeElement | undefined;
  readonly classList = {
    add: (...classNames: string[]) => {
      const current = this.className.split(" ").filter(Boolean);
      for (const className of classNames) {
        if (!current.includes(className)) {
          current.push(className);
        }
      }
      this.className = current.join(" ");
    },
  };
  private ownText = "";
  readonly tagName: string;

  constructor(tagName: string) {
    this.tagName = tagName;
  }

  get textContent(): string {
    return this.ownText + this.children.map((child) => child.textContent).join("");
  }

  set textContent(value: string) {
    this.ownText = value;
    this.children = [];
  }

  append(...children: FakeElement[]): void {
    for (const child of children) {
      this.addChild(child);
    }
  }

  appendChild(child: FakeElement): FakeElement {
    this.addChild(child);
    return child;
  }

  prepend(...children: FakeElement[]): void {
    for (let index = children.length - 1; index >= 0; index -= 1) {
      const child = children[index];
      if (!child) {
        continue;
      }
      child.removeFromParent();
      child.parent = this;
      this.children.unshift(child);
    }
  }

  getContext() {
    return {
      font: "",
      measureText(text: string) {
        return {
          width: text.length * 5,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
        };
      },
    };
  }

  querySelectorAll<T = FakeElement>(selector: string): T[] {
    const selectors = selector.split(",").map((value) => value.trim());
    const matches: FakeElement[] = [];
    for (const child of this.children) {
      matches.push(...child.querySelectorAllMatching(selectors));
    }
    return matches as T[];
  }

  querySelector(selector: string): FakeElement | null {
    if (selector.startsWith(".")) {
      return findByClass(this, selector.slice(1)) ?? null;
    }
    if (this.tagName.toLowerCase() === selector.toLowerCase()) {
      return this;
    }
    for (const child of this.children) {
      const match = child.querySelector(selector);
      if (match) {
        return match;
      }
    }
    return null;
  }

  private querySelectorAllMatching(selectors: string[]): FakeElement[] {
    const matches = selectors.some((selector) => this.matches(selector)) ? [this] : [];
    for (const child of this.children) {
      matches.push(...child.querySelectorAllMatching(selectors));
    }
    return matches;
  }

  private matches(selector: string): boolean {
    if (selector === 'img[style*="z-index"]') {
      return (
        this.tagName.toLowerCase() === "img" &&
        this.style.zIndex !== undefined &&
        this.style.zIndex !== ""
      );
    }
    if (selector === '.layout-textbox[style*="z-index"]') {
      return (
        this.className.split(" ").includes("layout-textbox") &&
        this.style.zIndex !== undefined &&
        this.style.zIndex !== ""
      );
    }
    if (selector.startsWith(".")) {
      return this.className.split(" ").includes(selector.slice(1));
    }
    return this.tagName.toLowerCase() === selector.toLowerCase();
  }

  private removeFromParent(): void {
    if (!this.parent) {
      return;
    }
    const index = this.parent.children.indexOf(this);
    if (index !== -1) {
      this.parent.children.splice(index, 1);
    }
    this.parent = undefined;
  }

  private addChild(child: FakeElement): void {
    child.removeFromParent();
    child.parent = this;
    this.children.push(child);
  }
}

export const fakeDocument = {
  createElement(tagName: string): FakeElement {
    return new FakeElement(tagName);
  },
  createTextNode(text: string): FakeElement {
    const node = new FakeElement("#text");
    node.textContent = text;
    return node;
  },
} as unknown as Document;

export function collectPmAnchors(element: FakeElement): FakeElement[] {
  const ownAnchorKeys = Object.keys(element.dataset).filter((key) =>
    key.toLowerCase().includes("pm"),
  );
  const anchors = ownAnchorKeys.length > 0 ? [element] : [];

  for (const child of element.children) {
    anchors.push(...collectPmAnchors(child));
  }

  return anchors;
}

export function findByClass(element: FakeElement, className: string): FakeElement | undefined {
  if (element.className.split(" ").includes(className)) {
    return element;
  }

  for (const child of element.children) {
    const match = findByClass(child, className);
    if (match) {
      return match;
    }
  }

  return undefined;
}

export function collectByClass(element: FakeElement, className: string): FakeElement[] {
  const matches = element.className.split(" ").includes(className) ? [element] : [];

  for (const child of element.children) {
    matches.push(...collectByClass(child, className));
  }

  return matches;
}
