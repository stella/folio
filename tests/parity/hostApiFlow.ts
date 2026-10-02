import fc from "fast-check";
import {
  PAGED_SCROLL_NAVIGATION_CASES,
  SCROLL_NAVIGATION_CASES,
} from "../../packages/playground/src/scrollParityBridge";

export const HOST_NAVIGATION_CASES = [
  ...Object.values(SCROLL_NAVIGATION_CASES).map(({ method }) => ({
    type: "document" as const,
    method,
  })),
  ...Object.values(PAGED_SCROLL_NAVIGATION_CASES).map(({ method }) => ({
    type: "paged" as const,
    method,
  })),
];

// Every shrink still visits the full contract, in a seed-dependent order.
export const hostApiFlowArbitrary = fc.record({
  navigation: fc.shuffledSubarray(HOST_NAVIGATION_CASES, {
    minLength: HOST_NAVIGATION_CASES.length,
    maxLength: HOST_NAVIGATION_CASES.length,
  }),
  edits: fc.array(fc.constantFrom("café", "東京", "e\u0301", "مرحبا", "👩🏽‍⚖️", "<&>"), {
    minLength: HOST_NAVIGATION_CASES.length,
    maxLength: HOST_NAVIGATION_CASES.length,
  }),
  replacementAfter: fc.integer({ min: 1, max: HOST_NAVIGATION_CASES.length - 1 }),
});

export type ScrollObservation = {
  scrollTop: number;
  top: number;
  bottom: number;
  viewportTop: number;
  viewportBottom: number;
};

type CheckNavigationOptions = {
  before: ScrollObservation;
  after: ScrollObservation;
  outerBefore: number;
  outerAfter: number;
};

// Bounding boxes, rather than Playwright visibility, establish overflow visibility.
export const navigationWasEffective = ({
  before,
  after,
  outerBefore,
  outerAfter,
}: CheckNavigationOptions): boolean =>
  before.top > before.viewportBottom &&
  after.scrollTop > before.scrollTop &&
  after.top >= after.viewportTop &&
  after.top < after.viewportBottom &&
  after.bottom <= after.viewportBottom &&
  outerAfter === outerBefore;
