FROM oven/bun:1.4.4-slim@sha256:1846adf22496bd711036aae8be147e57bdebbaea2e5264e3e0b5af1db187d9a4 AS bun
FROM node:26.10.0-bookworm-slim@sha256:3ffc19ea878019d9e9ae8971732ad4a03cda44f167107173174b60ed7c65bed3 AS node
FROM rust:1.99.0-bookworm@sha256:114c7a4425406451c2866b6aafe69fe29b1b298832db1277d411ac73c82d04d6 AS tools

COPY --from=bun /usr/local/bin/bun /usr/local/bin/bun
COPY --from=node /usr/local/bin/node /usr/local/bin/node
RUN ln -s /usr/local/bin/bun /usr/local/bin/bunx
RUN apt-get update && apt-get install -y --no-install-recommends \
      bash ca-certificates curl git python3 \
    && rm -rf /var/lib/apt/lists/*
RUN rustup component add clippy rustfmt && rustup target add wasm32-unknown-unknown

# Keep source, tests, toolchains and dependency caches for offline work.
# Resolve submodules from a full Git clone, then omit Git metadata from the image.
FROM tools AS source
COPY . /src
WORKDIR /src
RUN GIT_TERMINAL_PROMPT=0 git -c credential.helper= -c protocol.allow=never \
      -c protocol.https.allow=always submodule update --init --recursive \
    && find /src -name .git -prune -exec rm -rf {} +

FROM tools
COPY --from=source /src /src
WORKDIR /src
RUN BUN_INSTALL_GLOBAL_STORE=0 bun install --frozen-lockfile --ignore-scripts --network-concurrency=8
RUN cargo fetch --locked
# The CLI must match the wasm-bindgen runtime in the committed Cargo lockfile.
RUN cargo metadata --locked --format-version 1 --features wasm > /tmp/cargo-metadata.json \
    && version=$(python3 -c 'import json; packages=json.load(open("/tmp/cargo-metadata.json"))["packages"]; print(next(p["version"] for p in packages if p["name"] == "wasm-bindgen"))') \
    && cargo install wasm-bindgen-cli --version "$version" --locked \
    && rm /tmp/cargo-metadata.json
RUN cargo build --locked --workspace
RUN bun --filter @stll/docx-core wasm:generate && bun --filter @stll/folio-core wasm:generate
RUN bun run build
RUN bun x --no-install playwright install --with-deps chromium

CMD ["bash"]
