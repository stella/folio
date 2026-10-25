import { panic } from "better-result";

let nextIdentity = 0;

type ContentIdentityLimits = {
  maxBytes: number;
  maxEntries: number;
};

export class BoundedContentIdentityCache {
  readonly #limits: ContentIdentityLimits;
  readonly #identities = new Map<string, string>();
  #bytes = 0;

  constructor(limits: ContentIdentityLimits) {
    if (
      !Number.isSafeInteger(limits.maxBytes) ||
      limits.maxBytes <= 0 ||
      !Number.isSafeInteger(limits.maxEntries) ||
      limits.maxEntries <= 0
    ) {
      panic("Content identity cache limits must be positive safe integers.");
    }
    this.#limits = { ...limits };
  }

  intern(content: string): string {
    const existing = this.#identities.get(content);
    if (existing !== undefined) {
      this.#identities.delete(content);
      this.#identities.set(content, existing);
      return existing;
    }

    if (nextIdentity === Number.MAX_SAFE_INTEGER) {
      panic("Content identity space exhausted.");
    }
    nextIdentity += 1;
    const identity = `c${nextIdentity.toString(36)}`;
    const bytes = (content.length + identity.length) * 2;
    if (bytes > this.#limits.maxBytes) return identity;

    while (
      this.#identities.size >= this.#limits.maxEntries ||
      this.#bytes + bytes > this.#limits.maxBytes
    ) {
      const oldest = this.#identities.entries().next().value;
      if (oldest === undefined) {
        panic("Content identity cache accounting must match retained entries.");
      }
      const [oldContent, oldIdentity] = oldest;
      this.#identities.delete(oldContent);
      this.#bytes -= (oldContent.length + oldIdentity.length) * 2;
    }

    this.#identities.set(content, identity);
    this.#bytes += bytes;
    return identity;
  }
}

const identities = new BoundedContentIdentityCache({
  maxBytes: 1024 * 1024,
  maxEntries: 128,
});

export const boundedContentIdentity = (content: string): string =>
  identities.intern(content);
