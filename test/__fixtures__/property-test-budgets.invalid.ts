import { test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import { propertyConfig, propertyTestTimeout } from "../property-testing";

const property = fc.property(fc.integer(), (value) => Number.isFinite(value));
setDefaultTimeout(propertyTestTimeout());

test("missing", () => fc.assert(property, propertyConfig()), propertyTestTimeout());
test("text", () => fc.assert(property, propertyConfig()), propertyTestTimeout("30000"));
test("NaN", () => fc.assert(property, propertyConfig()), propertyTestTimeout(Number.NaN));
test("infinite", () => fc.assert(property, propertyConfig()), propertyTestTimeout(Infinity));
test("zero", () => fc.assert(property, propertyConfig()), propertyTestTimeout(0));
test("negative", () => fc.assert(property, propertyConfig()), propertyTestTimeout(-30_000));
test("spoof", () => fc.assert(property, propertyConfig()), "propertyTestTimeout(30_000)");
test("invalid default", () => fc.assert(property, propertyConfig()));
