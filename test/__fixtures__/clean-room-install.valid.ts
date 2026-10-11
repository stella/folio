import { $ } from "bun";
import { installCleanRoomDependencies } from "../../scripts/clean-room-install";

await installCleanRoomDependencies({ consumerDir: "/tmp/consumer", dependencies: ["react"] });
await $`bun pm pack`;
Bun.spawn(["bun", "run", "build"]);
