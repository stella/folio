import { GlobalRegistrator } from "@happy-dom/global-registrator";

// Vue captures its document at module initialization. Establish that prerequisite
// before any composable test imports it, then leave DOM lifecycles to each test.
GlobalRegistrator.register();
await import("vue");
await GlobalRegistrator.unregister();
