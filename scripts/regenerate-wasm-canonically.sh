#!/usr/bin/env bash
# Regenerate a committed WebAssembly artifact the way CI builds it.
#
#   scripts/regenerate-wasm-canonically.sh @stll/docx-core   # the DOCX kernel
#   scripts/regenerate-wasm-canonically.sh @stll/folio-core  # the text shaper
#
# The committed .wasm is compared byte-for-byte against a fresh build, and CI
# builds on linux/amd64. On any other platform the local build is a legitimate
# difference rather than drift, so `wasm:check` skips the .wasm there and CI is
# the first place a stale artifact would surface. Running the generate step in
# this image produces an artifact for the target CI builds for, rather than for
# the developer's own machine.
#
# It is not a guarantee of the same bytes: the pinned Rust toolchain and
# wasm-bindgen fix the compiler, but the post-processing tools are built here
# and on the runner separately, and they have been seen to differ by a few
# hundred bytes. CI remains the arbiter; when its drift check fails it uploads
# what it built, and that artifact is what to commit.
#
# Requires Docker with linux/amd64 emulation available. The image is cached
# after the first run; the toolchain versions come from the repository, so a
# bump in rust-toolchain.toml, Cargo.lock or package.json rebuilds it.
set -euo pipefail

package="${1:-}"
if [[ -z "${package}" ]]; then
  echo "usage: ${0##*/} <@stll/docx-core|@stll/folio-core>" >&2
  exit 2
fi

root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
image="folio-canonical-wasm"

rust_version="$(sed -n 's/^channel = "\(.*\)"$/\1/p' "${root}/rust-toolchain.toml")"
wasm_bindgen_version="$(
  awk '/^name = "wasm-bindgen"$/ { getline; gsub(/version = |"/, ""); print; exit }' \
    "${root}/Cargo.lock"
)"
bun_version="$(sed -n 's/.*"packageManager": "bun@\(.*\)".*/\1/p' "${root}/package.json")"

docker build \
  --platform linux/amd64 \
  --build-arg "RUST_VERSION=${rust_version}" \
  --build-arg "WASM_BINDGEN_VERSION=${wasm_bindgen_version}" \
  --build-arg "BUN_VERSION=${bun_version}" \
  --file "${root}/scripts/canonical-wasm.Dockerfile" \
  --tag "${image}" \
  "${root}"

# Two named volumes, both kept between runs: the registry cache so a rerun does
# not refetch every crate, and the container's own `node_modules`.
#
# The install has to happen inside: bun resolves some packages through symlinks
# into a store outside the repository, which do not exist in the container. The
# volume masks the host's tree rather than replacing it, so the host keeps the
# binaries its own platform needs.
docker run --rm \
  --platform linux/amd64 \
  --volume "${root}:/workspace" \
  --volume "folio-canonical-wasm-cargo:/cargo" \
  --volume "folio-canonical-wasm-modules:/workspace/node_modules" \
  --env CARGO_HOME=/cargo \
  "${image}" \
  bash -c "bun install --frozen-lockfile && bun --filter '${package}' wasm:generate"
