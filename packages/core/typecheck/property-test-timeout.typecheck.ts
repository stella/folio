import { propertyTestTimeout } from "../../../../test/property-timeout";

propertyTestTimeout(15_000);

// @ts-expect-error A property timeout must state its base budget.
propertyTestTimeout();

// @ts-expect-error Budget values must be numeric.
propertyTestTimeout("15000");
