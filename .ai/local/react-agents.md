## React Adapter

- Keep the package a thin renderer over folio-core. Shared editor behavior belongs in
  a framework-neutral core manager; hooks should only bind that manager to React.
- React Compiler runs during the package build, but imperative ProseMirror and ref
  paths can bail out. Preserve load-bearing manual memoization in bailout files unless
  profiling and `bun run check:react-compiler` verify its removal.
- Keep the React and Vue adapter contracts aligned; update parity checks with any
  intentional divergence.
