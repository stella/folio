// Every test that starts a child is bounded by its children.
import { beforeAll, setDefaultTimeout, test } from "bun:test";

import { propertyTestTimeout } from "../../property-timeout";
import { expectExited, runChild, subprocessTestTimeout } from "../../subprocess";

setDefaultTimeout(subprocessTestTimeout());

const version = () => expectExited(runChild(["bun", "--version"]));

beforeAll(() => version());

test("relies on the default", () => version());

test("states its children", () => [version(), version()], subprocessTestTimeout(2));

test("states its own work", () => version(), subprocessTestTimeout(1, 30_000));

test("options form", { timeout: subprocessTestTimeout() }, () => version());

test("no child", () => {}, propertyTestTimeout(5_000));
