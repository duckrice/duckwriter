# DuckWriter

A free word processor for Windows that opens and saves Word documents (.docx).
Part of the duckware family, next to [Duckument](https://github.com/duckrice/duckument).

It's early: see [PLAN.md](PLAN.md) for what works and what's coming.

## Building

You need Node.js, Rust and the Visual Studio C++ build tools.

```bash
npm install
npm run tauri build -- --no-bundle
```

The app is then at `src-tauri/target/release/duckwriter.exe`.

## Tests

```bash
npm test
npm run samples
npm run smoke
```

`npm test` checks reading, editing and saving documents. `npm run smoke` drives
the built app (off-screen) and checks typing, the toolbar, opening and saving;
add `-- --screenshots` for pictures of each step.
