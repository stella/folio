# Build environment for the committed WebAssembly artifacts.
#
# The bytes are compared byte-for-byte against what CI builds, and CI builds on
# linux/amd64. A developer on another platform regenerates through this image so
# the artifact they commit is the one CI will reproduce.
#
# Versions come from the repository: the Rust toolchain from rust-toolchain.toml
# and wasm-bindgen from Cargo.lock, both passed in as build arguments so this
# file never becomes a second place they are pinned.
ARG RUST_VERSION
FROM rust:${RUST_VERSION}-bookworm

ARG WASM_BINDGEN_VERSION
ARG BUN_VERSION

RUN rustup target add wasm32-unknown-unknown
RUN cargo install wasm-bindgen-cli --version "${WASM_BINDGEN_VERSION}" --locked
RUN curl -fsSL https://bun.sh/install | bash -s "bun-v${BUN_VERSION}"
ENV PATH="/root/.bun/bin:${PATH}"

WORKDIR /workspace
