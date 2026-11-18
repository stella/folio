# Build environment for the committed WebAssembly artifacts.
#
# The bytes are compared byte-for-byte against what CI builds, and CI builds on
# linux/amd64. A developer on another platform regenerates through this image so
# the artifact they commit is the one CI will reproduce.
#
# Rust must match rust-toolchain.toml; the shared toolchain checker guards both.
# wasm-bindgen comes from Cargo.lock through a build argument.
FROM rust:1.100.0-bookworm@sha256:ade39a710ef85f3506cffcfec3e8a24ba090badef1dff2fb798f71d878833fea

ARG WASM_BINDGEN_VERSION
ARG BUN_VERSION

RUN rustup target add wasm32-unknown-unknown
RUN cargo install wasm-bindgen-cli --version "${WASM_BINDGEN_VERSION}" --locked
RUN curl -fsSL https://bun.sh/install | bash -s "bun-v${BUN_VERSION}"
ENV PATH="/root/.bun/bin:${PATH}"

WORKDIR /workspace
