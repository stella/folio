import { expect, test } from "bun:test";
import manifest from "../packages/react/package.json";
import { REACT_PEER_FLOOR, reactPeerFloor, reactPeerInstallArgs } from "./packaged-consumer-lib";

test("clean-room consumers install the declared React and React DOM peer floor", () => {
  const dependencies = reactPeerInstallArgs();
  for (const peer of ["react", "react-dom"] as const) {
    const dependency = dependencies.find((argument) => argument.startsWith(`${peer}@`));
    expect(dependency).toBe(`${peer}@${manifest.peerDependencies[peer].slice(1)}`);
    expect(Bun.semver.satisfies(REACT_PEER_FLOOR.version, manifest.peerDependencies[peer])).toBe(
      true,
    );
  }
  for (const peer of ["@types/react", "@types/react-dom"]) {
    expect(dependencies).toContain(`${peer}@^${REACT_PEER_FLOOR.major}`);
  }
});

test("the consumer floor follows minor and major peer changes", () => {
  for (const version of ["19.2.0", "19.3.1", "20.1.0"]) {
    expect(reactPeerFloor({ react: `^${version}`, "react-dom": `^${version}` })).toEqual({
      version,
      major: version.split(".").at(0),
    });
  }
});

test("unsupported or divergent peer policies fail before staging consumers", () => {
  for (const range of ["^18 || ^19", ">=19.2.0", "*", "^19.2", "19.2.0"]) {
    expect(() => reactPeerFloor({ react: range, "react-dom": range })).toThrow();
  }
  expect(() => reactPeerFloor({ react: "^19.2.0", "react-dom": "^19.3.0" })).toThrow();
});
