// Six limits that are bare or not the subprocess budget.
import { setDefaultTimeout, test } from "bun:test";

import { scanTestTimeout } from "../../scan-timeout";
import { runChild, subprocessTestTimeout } from "../../subprocess";

const LIMIT_MS = 5 * 1000;

setDefaultTimeout(subprocessTestTimeout());

const version = () => runChild(["bun", "--version"]);

test("bare limit", () => version(), 5_000);

test("named limit", () => version(), LIMIT_MS);

test("options limit", { timeout: 5_000 }, () => version());

test("child limit", () => Bun.spawnSync(["bun", "--version"], { timeout: 5000 }));

test("hang guard", () => runChild(["bun", "--version"], { hangTimeoutMs: 5000 }));

test("smaller helper budget", () => version(), scanTestTimeout(5_000));
