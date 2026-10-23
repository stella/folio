import { expect, test } from "bun:test";

import { BoundedContentIdentityCache } from "./boundedContentIdentity";

test("equal retained content shares an identity without hash collisions", () => {
  const cache = new BoundedContentIdentityCache({
    maxBytes: 1000,
    maxEntries: 4,
  });
  const identity = cache.intern("font facts");
  expect(cache.intern("font facts")).toBe(identity);
  expect(cache.intern("different facts")).not.toBe(identity);
});

test("byte accounting evicts old content before reaching the entry limit", () => {
  const cache = new BoundedContentIdentityCache({
    maxBytes: 100,
    maxEntries: 10,
  });
  const first = cache.intern("first".repeat(5));
  cache.intern("second".repeat(5));
  expect(cache.intern("first".repeat(5))).not.toBe(first);
});

test("entry accounting retains recently reused content", () => {
  const cache = new BoundedContentIdentityCache({
    maxBytes: 1000,
    maxEntries: 2,
  });
  const hot = cache.intern("hot");
  const cold = cache.intern("cold");
  expect(cache.intern("hot")).toBe(hot);
  cache.intern("new");
  expect(cache.intern("hot")).toBe(hot);
  expect(cache.intern("cold")).not.toBe(cold);
});

test("oversized content has a short unique identity and cannot evict retained content", () => {
  const cache = new BoundedContentIdentityCache({
    maxBytes: 100,
    maxEntries: 2,
  });
  const retained = cache.intern("retained");
  const large = "a".repeat(10_000);
  const first = cache.intern(large);
  const second = cache.intern(large);
  expect(first.length).toBeLessThanOrEqual(12);
  expect(second.length).toBeLessThanOrEqual(12);
  expect(second).not.toBe(first);
  expect(cache.intern("retained")).toBe(retained);
});
