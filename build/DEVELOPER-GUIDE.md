# sdkjs webpack Build — Developer Guide

## Build Benchmarks

| Scenario | Tool | Wall clock |
|---|---|---|
| All 4 modules (grunt, sequential) | Closure Compiler (Java) | **399 s** |
| All 4 modules (webpack, parallel) | esbuild (Node.js) | **~18 s** (was ~50 s with Terser) |
| Single module cold start | webpack + esbuild | ~6 s (word; was ~26 s with Terser) |
| **Single module warm start (FS cache)** | webpack + esbuild | **~0.6 s** |
| Single module, dev mode (no minifier) | webpack | ~5.5 s |
| Watch: initial build | webpack dev mode | ~5.7 s |
| **Watch: incremental rebuild** | webpack dev mode | **~4 s** |

The filesystem cache (added to `webpack.sdk.factory.mjs`) is the biggest single DX win:
a production module restarts in **0.6 s** after the first build. The cache is stored in
`build/.webpack-cache/` and is automatically invalidated when source files, JSON configs,
or the factory file itself changes.

---

## Why HMR Is Not Possible

Hot Module Replacement requires three things this codebase cannot provide:

1. **Module granularity.** The `sdk-concat-loader` bundles all 400+ source files into a
   single webpack module — one concatenated blob. There is no smaller unit for webpack to
   replace independently.

2. **State survival.** The SDK maintains complex global state (undo history, cursor,
   selection, font cache, collaborative session) in `window.AscCommon`. Swapping the SDK
   bundle mid-session would corrupt all of it. The only safe recovery is a full page reload.

3. **No `module.hot.accept()` boundary.** The SDK is a monolithic state machine with no
   re-enterable initialisation path. There is nowhere to write an HMR accept handler.

---

## Three Developer Workflows (pick by task)

### Workflow 1 — Active JS Debugging (no build step) ★ recommended

Use the existing `develop/scripts.js` mode. The HTML pages load each source file
individually — no build required, changes appear on the next browser refresh.

```bash
# One-time setup (from sdkjs/build/):
npm run develop
# → writes develop/sdkjs/{word,cell,slide,visio}/scripts.js

# Switch nginx to source mode (from DocumentServer/develop/):
make front-dev           # static source mode
make front-dev-live      # source mode + livereload on port 35729
```

**How it works:** `develop/scripts.js` sets `window.sdk_scripts` to an array of URLs
pointing at every raw source file. The HTML page's inline script calls `document.write`
for each URL, loading them synchronously before RequireJS starts. nginx serves the source
tree directly at `/sdkjs/` and `/web-apps/`.

**Switch back to production bundles:**
```bash
make front-prod
```

### Workflow 2 — Bundle Verification / Integration Testing

Use `npm run watch:word` (or whichever module you're working on). webpack rebuilds the
bundle in development mode every time a source file changes.

```bash
# From sdkjs/build/:
npm run watch:word   # ~5.7 s initial, ~4 s per incremental rebuild
npm run watch:cell
npm run watch:slide
npm run watch:visio
```

- Development mode disables minification → **8× faster** per-rebuild vs production
- The concat loader already calls `addContextDependency()` for glob directories, so
  adding a new source file triggers a rebuild automatically without restarting watch
- Pair with `make front-dev` (nginx serving from the deploy directory) for browser
  auto-refresh via livereload

### Workflow 3 — Production Build (CI / release)

```bash
# All 4 modules in parallel (~50 s first run, ~2–3 s warm):
npm run build

# Individual modules:
npm run build:word
npm run build:cell
npm run build:slide
npm run build:visio
```

---

## Source Maps

Source maps are generated in development mode (`devtool: 'source-map'`).

**Current limitation:** The concat loader returns all source files as a single webpack
module attributed to `dummy.js`. The generated `.map` file contains only 3 sources
(webpack bootstrap + runtime + `dummy.js`). In DevTools you see one 28 MB file, not the
individual source files.

**Improvement (implemented):** The concat loader now runs asynchronously and generates a
proper source map with one source entry per input file using `source-map-js`. Each line
in the bundle maps back to its originating `.js` file and line number. This requires
`source-map-js` in `dependencies` (already added to `package.json`).

**Trade-off:** The `.map` files are large (34 MB for `sdk-all.js`). In watch mode this
is fine — the browser only fetches the map when DevTools is open. In production, source
maps are disabled (`devtool: false`).

**Best debugging experience:** Workflow 1 (develop/scripts.js mode). Individual files
load by their real names, DevTools shows the source tree as-is.

---

## Filesystem Cache

Added to `webpack.sdk.factory.mjs` via `cache: { type: 'filesystem' }`. The cache lives
at `build/.webpack-cache/` (gitignored). It is invalidated automatically by webpack when:

- Any source file registered via `addDependency()` changes
- Any glob-watched directory changes (new/deleted file)
- The module config JSON (`configs/word.json` etc.) changes
- `webpack.sdk.factory.mjs` itself changes (via `buildDependencies.config`)

On the first build after a clean checkout the cache is cold — full build time applies.
Every subsequent start (watch restart, CI warm cache, local re-run) hits the cache.

Measured speedups (production mode, word module):

| Run | Time |
|---|---|
| Cold | ~30–34 s |
| Warm | **0.6 s** |

---

## DefinePlugin — Why Only the Unprefixed Form

The Gruntfile uses `--define=window.AscCommon.g_cCompanyName='x'` (Closure Compiler
semantics). Closure Compiler understands that `--define` applies to both the declaration
and all usages.

webpack's DefinePlugin does a simpler AST-level text replacement. If you add
`'window.AscCommon.g_cCompanyName'` as a key, DefinePlugin replaces **every** occurrence
of that expression — including the LHS of the declaration in `common/commonDefines.js`:

```js
// commonDefines.js (in the min chunk):
window.AscCommon.g_cCompanyName = "onlyoffice";
// → DefinePlugin turns this into:
"onlyoffice" = "onlyoffice";   // ← invalid assignment, the minifier rejects
```

Only the unprefixed `AscCommon.g_cXxx` form is safe. This covers all call-site
comparisons (`if (AscCommon.g_cIsBeta === 'true')`) and enables esbuild dead-code
elimination. The `window.AscCommon.g_cXxx = "..."` declarations in `commonDefines.js`
stay as runtime defaults — they run once on load, then the constant-folded call-sites
take over.

---

## CJS/ESM Conflict — Fixed

`build/package.json` declares `"type": "module"`, which makes Node.js treat all `.js`
files in the `build/` directory as ESM. The three pipeline scripts use `require()` (CJS).

**Fix applied:** Scripts renamed from `.js` to `.cjs`:

| Before | After |
|---|---|
| `scripts/build-pipeline.js` | `scripts/build-pipeline.cjs` |
| `scripts/build-develop.js` | `scripts/build-develop.cjs` |
| `scripts/deploy-assets.js` | `scripts/deploy-assets.cjs` |

`package.json` `scripts` entries updated to match. `build-pipeline.cjs` internal
references to the other two scripts also updated.

---

## Concat Loader Improvements

`loaders/sdk-concat.cjs` was updated with two optimisations:

**1. Async parallel file reads**

The original loader used a synchronous `fs.readFileSync` loop (serial I/O). The loader
now uses `this.async()` and `Promise.all(files.map(f => fs.promises.readFile(f)))`,
reading all source files in parallel. Measured improvement: ~10–20% faster cold builds
on SSD.

**2. Per-file source maps**

The loader now generates a `SourceMapGenerator` mapping with one entry per source file
and one mapping per line. This is passed to webpack via `this.callback(null, code, map)`.
In development mode, DevTools can navigate to individual source files instead of showing
a 28 MB concatenated blob.

Requires `source-map-js` in `dependencies` (already added). If the package is absent the
loader falls back gracefully — build succeeds without source maps.

---

## Retired Grunt CLI Flags — Why `--level`, `--formatting`, `--src` Have No Equivalent

`build-pipeline.cjs` takes all configuration through env vars and rejects any CLI
argument outright — a stale caller still passing old Grunt flags fails loudly instead
of silently building an incomplete bundle. Most flags map directly to an env var
(`--addon` → `SDK_ADDONS`, `--desktop=true` → `SDK_PLATFORM=desktop`, etc. — see the
error message itself for the full mapping). Three don't:

- **`--level`** — Closure Compiler's `--compilation_level` (`ADVANCED` vs
  `WHITESPACE_ONLY`). There's no esbuild flag with the same meaning. The one distinction
  the old flag was actually used for in practice — a lighter pass for desktop/mobile —
  is preserved directly: `SDK_PLATFORM=desktop|mobile` sets esbuild's `minifySyntax: false`
  in `webpack.sdk.factory.mjs`. There's no general replacement beyond that one case.

- **`--formatting`** — Closure's pretty-print output flag (e.g. `PRETTY_PRINT`).
  The minifier has no equivalent output mode. To read unminified output, use development mode
  instead (`npm run watch:*`, or `NODE_ENV=development`), which skips minification entirely
  rather than asking it to format its output legibly.

- **`--src`** — overrode the source-tree root for a one-off build against a different
  checkout. The webpack config hardcodes `SRC_ROOT` relative to `build/` (`path.resolve(__dirname, '..')`);
  there's no env var to point it elsewhere. If you need this, it's a config change,
  not a flag.

---

## Resolved — ES5 Syntax Downleveling / webpack `target`

The old Gruntfile forced `--language_out=ECMASCRIPT5` on every Closure Compiler
build. esbuild (this pipeline's minifier) only minifies, with `target: 'es2022'` capping
the syntax it may introduce (e.g. `?.`, `??`), so `let`/`const`/arrow functions/classes/template literals now ship
as-authored. `webpack.sdk.factory.mjs` also sets no explicit `target`, and
`build/package.json` has no `browserslist` field, so webpack 5 falls back to
its own implicit target resolution.

**Resolved: ES6+ is safe.** Desktop pins CEF 107 (Chromium 107, October
2022 — ES2022-capable; see `Euro-Office/desktop-apps/win-linux/defaults.pri`).
Mobile floor is iOS 17/Safari 17 → ES2022 per `web-apps/build/browser-floor.mjs`
/ `browser-floor.manifest.mjs`, with Android WebView tracking current Chrome
regardless of OS version — `web-apps` commit `7ebda70613` ("centralise browser
target to ES2022, fixing i18next crash") traces a real mobile crash to
over-aggressive ES2015 downleveling (esbuild's `target:'es2015'` mis-compiling
i18next 25's class fields), the opposite direction of concern from this
change. No downlevel step needed.

---

## Quick Reference

```bash
# From sdkjs/build/

# Production build (all modules, parallel, warm ~2 s):
npm run build

# Watch a single module (dev mode, ~4 s incremental):
npm run watch:word

# Generate develop/scripts.js for source-file-per-request dev mode:
npm run develop

# Individual production module:
npm run build:word
```

```bash
# From DocumentServer/develop/

make front-dev         # nginx → source trees (no build needed after npm run develop)
make front-dev-live    # + livereload on port 35729
make front-prod        # nginx → compiled deploy bundles
```
