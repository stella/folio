import { describe, expect, test } from "bun:test";

import { imageLuminanceFilter } from "./imageLuminance";

const FILTER_FUNCTION = /(contrast|brightness)\(([-\d.e]+)\)/gu;

const applyFilter = (filter: string, channel: number): number => {
  let value = channel / 255;
  for (const match of filter.matchAll(FILTER_FUNCTION)) {
    const amount = Number(match[2]);
    value = match[1] === "contrast" ? amount * value + (1 - amount) / 2 : amount * value;
    value = Math.max(0, Math.min(1, value));
  }
  return Math.round(value * 255);
};

describe("imageLuminanceFilter", () => {
  test("keeps washout whites white", () => {
    const filter = imageLuminanceFilter({ brightness: 70.001, contrast: -70 });
    expect(filter).toBe("contrast(0.157067) brightness(1.910013)");
    expect(applyFilter(filter ?? "", 255)).toBe(255);
  });

  test("orders negative-intercept filters before contrast", () => {
    expect(imageLuminanceFilter({ brightness: -20, contrast: 50 })).toBe(
      "brightness(0.769231) contrast(2.6)",
    );
  });

  test("handles identity and percentage boundaries", () => {
    expect(imageLuminanceFilter({ brightness: 0, contrast: 0 })).toBeUndefined();
    expect(imageLuminanceFilter({ brightness: 100 })).toBe("contrast(0.333333) brightness(3)");
    expect(imageLuminanceFilter({ brightness: -100 })).toBe("brightness(0.333333) contrast(3)");
    expect(imageLuminanceFilter({ contrast: 100 })).toBe("brightness(1) contrast(1000)");
    expect(imageLuminanceFilter({ contrast: -100 })).toBe("contrast(0) brightness(1)");
  });
});
