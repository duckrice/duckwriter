# DuckWriter

A free word processor for Windows that opens and saves Word documents (.docx),
made for scientific writing: equations, footnotes, captions and
cross-references, citations and a reference list, tracked changes and comments.
Part of the duckware family, next to [Duckument](https://github.com/duckrice/duckument).

See [PLAN.md](PLAN.md) for what it does and what's coming.

## Install

Download **DuckWriter-Setup.exe** from the
[latest release](https://github.com/duckrice/duckwriter/releases/latest) and run it.
Windows may say "Windows protected your PC", because the installer isn't
code-signed: click **More info**, then **Run anyway**. DuckWriter then keeps
itself up to date.

## Building

You need Node.js, Rust and the Visual Studio C++ build tools.

```bash
npm install
npm run tauri build -- --no-bundle
```

The app is then at `src-tauri/target/release/duckwriter.exe`. Without
`--no-bundle` the installer is made too (in `src-tauri/target/release/bundle/nsis`);
that needs the update signing key in `TAURI_SIGNING_PRIVATE_KEY`.

## Releasing

1. Put the new version in `src-tauri/tauri.conf.json`, `src-tauri/Cargo.toml` and `package.json`.
2. Commit, then tag and push: `git tag v1.0.1 && git push origin main v1.0.1`.

GitHub Actions ([release.yml](.github/workflows/release.yml)) builds the
installer, signs the update with the `TAURI_SIGNING_PRIVATE_KEY` secret, and
publishes the release with `DuckWriter-Setup.exe` and `latest.json`, which
installed copies check for updates.

## Tests

```bash
npm test
npm run samples
npm run smoke
```

`npm test` checks reading, editing and saving documents. `npm run smoke` drives
the built app (off-screen) and checks typing, the toolbar, opening and saving;
add `-- --screenshots` for pictures of each step.
