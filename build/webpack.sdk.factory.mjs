/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

/**
 * Shared webpack 5 config factory for all SDK modules (word, cell, slide, visio).
 *
 * sdkjs has no module system — all source files are global IIFE or bare-var scripts
 * that communicate via window.AscCommon / window.AscWord etc. The sdk-concat-loader
 * reads the ordered JSON configs and returns all files as ONE concatenated module,
 * preserving the shared scope that bare `var` declarations depend on.
 *
 * Each call to sdkConfig() returns TWO webpack compiler configs:
 *   [0]  sdk-all-min  — bootstrap files (device_scale, browser, skin, API defs …)
 *   [1]  sdk-all      — full feature set, wrapped in (function(window,undefined){…})(window)
 *
 * Environment variables (all optional, mirror the original Gruntfile.js):
 *   BUILD_ROOT        override deploy root; defaults to ../deploy/sdkjs
 *   SDK_PLATFORM      '' | 'desktop' | 'mobile'
 *   SDK_ADDONS        path.delimiter-separated list of addon directories
 *   COMPANY_NAME      default 'Euro-Office'
 *   PRODUCT_VERSION   default '0.0.0'
 *   BUILD_NUMBER      default '0'
 *   BETA              default 'false'
 *   APP_COPYRIGHT     default 'Copyright (C) Ascensio System SIA 2012-2025 …; Euro-Office contributors 2026 - …'
 *   PUBLISHER_URL     default 'https://github.com/Euro-Office/'
 *   NODE_ENV          'production' (default) | 'development'
 *   SDK_SOURCE_MAPS   '1' to emit source maps for a production build too
 *                      (development builds always get them regardless)
 *   WEBPACK_CACHE_DIR override filesystem cache location; defaults to build/.webpack-cache
 *   WATCH_AGGREGATE_TIMEOUT
 *                      override watch-mode debounce (ms); defaults to 500,
 *                      only takes effect under `--watch`
 */

import webpack    from 'webpack';
import TerserPlugin from 'terser-webpack-plugin';
import path       from 'path';
import os         from 'os';
import { fileURLToPath } from 'url';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const { parseAddonDirs, resolveBuildRoot, buildLicenseHeader, defaultAppCopyright, DEFAULT_PUBLISHER_URL } = require('./lib/env.cjs');

const __dirname     = path.dirname(fileURLToPath(import.meta.url));
const CONCAT_LOADER = path.join(__dirname, 'loaders', 'sdk-concat.cjs');
const DUMMY_ENTRY   = path.join(__dirname, 'dummy.js');

// webpack 5's own runtime bootstrap unconditionally emits a top-level
// `"use strict";` directive (see JavascriptModulesPlugin's renderMain), even
// though sdk-concat-loader concatenates files verbatim, with no per-file
// transform that would inject one — sdkjs's bare-var-across-files shared-scope model
// (see sdk-concat.cjs) predates strict mode and has never been validated
// against it (undeclared-assignment/delete/duplicate-parameter semantics all
// differ). Strip it to preserve the legacy script's non-strict semantics.
//
// The directive is replaced in place with a same-length no-op rather than
// removed, so no byte offset after it shifts — the accompanying source map
// (when devtool:'source-map' is on) stays valid without recomputing it, and
// this can never touch a legitimate `"use strict";` appearing verbatim
// inside a real source file's own string literals, since those come tens of
// KB into the bundle, well past PROLOGUE_SCAN_LIMIT.
export const STRICT_DIRECTIVE    = '"use strict";';
// Fallback only for callers (e.g. tests) that don't pass an explicit scanLimit.
// Production use always passes a limit derived from the actual banner length
// (see StripBundlePostprocessPlugin below), so a growing license.header can't
// silently push the directive past a hardcoded window.
export const PROLOGUE_SCAN_LIMIT = 2048;

// Pure and exported so it can be unit-tested (build/test/webpack-sdk-factory.test.cjs)
// without spinning up a full webpack build. Returns `source` unchanged if no
// directive is found within the prologue window.
export function stripBootstrapStrictDirective(source, scanLimit = PROLOGUE_SCAN_LIMIT) {
    const idx = source.slice(0, scanLimit).indexOf(STRICT_DIRECTIVE);
    if (idx === -1) return source;

    return source.slice(0, idx) +
        ';'.padEnd(STRICT_DIRECTIVE.length) +
        source.slice(idx + STRICT_DIRECTIVE.length);
}

// Exported (only) for the real-compilation integration test in
// build/test/webpack-sdk-factory.test.cjs: this plugin's correctness depends on
// webpack's own internal bootstrap-generation format and the minifier's output
// formatting, neither of which is a stable public API, so it needs an actual webpack
// run to validate, not just the pure-string-function tests above.
export class StripBundlePostprocessPlugin {
    constructor({ scanLimit = PROLOGUE_SCAN_LIMIT } = {}) {
        this.scanLimit = scanLimit;
    }

    apply(compiler) {
        compiler.hooks.compilation.tap('StripBundlePostprocessPlugin', (compilation) => {
            compilation.hooks.processAssets.tap(
                {
                    name: 'StripBundlePostprocessPlugin',
                    stage: webpack.Compilation.PROCESS_ASSETS_STAGE_REPORT,
                },
                (assets) => {
                    for (const name of Object.keys(assets)) {
                        if (!name.endsWith('.js')) continue;

                        const source = compilation.getAsset(name).source.source();
                        if (typeof source !== 'string') continue;

                        const patched = stripBootstrapStrictDirective(source, this.scanLimit);
                        if (patched === source) continue;

                        compilation.updateAsset(name, new webpack.sources.RawSource(patched));
                    }
                }
            );
        });
    }
}

// `Number(x) || 500` would silently treat WATCH_AGGREGATE_TIMEOUT=0 (a valid,
// if unusual, "no debounce" choice) the same as an unset/typo'd value, and a
// non-numeric typo would silently fall back to the default instead of
// surfacing the mistake. Exported for the unit test below.
export function resolveWatchAggregateTimeout(rawValue, fallback = 500) {
    if (rawValue === undefined || rawValue === '') return fallback;

    const parsed = Number(rawValue);
    if (!Number.isFinite(parsed) || parsed < 0) {
        throw new Error(`WATCH_AGGREGATE_TIMEOUT must be a non-negative number, got: ${JSON.stringify(rawValue)}`);
    }
    return parsed;
}

/**
 * @param {string} moduleName  'word' | 'cell' | 'slide' | 'visio'
 * @returns {object[]}  Two webpack compiler configs: [sdk-all-min, sdk-all]
 */
export function sdkConfig(moduleName) {
    // webpack's `mode` accepts only 'development' | 'production' | 'none'.
    // NODE_ENV is commonly set to other values (e.g. 'test' by Jest/Vitest/CI),
    // which would otherwise make webpack hard-fail before compiling anything.
    const env = process.env.NODE_ENV === 'development' ? 'development' : 'production';

    // Old grunt --map path could produce maps for the production/minified build too —
    // minification and source-map emission are independent concerns, so don't couple
    // "give me a map" to "give me a dev build" (which also disables minification).
    const emitSourceMaps = env === 'development' || process.env.SDK_SOURCE_MAPS === '1';

    const BUILD_ROOT = resolveBuildRoot(__dirname);

    const SRC_ROOT  = path.resolve(__dirname, '..');
    const OUT_DIR   = path.join(BUILD_ROOT, moduleName);
    const platform  = process.env.SDK_PLATFORM || '';
    const desktopOrMobile = platform === 'desktop' || platform === 'mobile';
    const addonDirs = parseAddonDirs();

    const companyName  = process.env.COMPANY_NAME     || 'Euro-Office';
    const version      = process.env.PRODUCT_VERSION  || '0.0.0';
    const buildNumber  = process.env.BUILD_NUMBER      || '0';
    const beta         = process.env.BETA              || 'false';

    // Matches the Euro-Office rebrand defaults main() picked up independently
    // (see build/Gruntfile.js history) after this migration branched off it.
    // Defaults come from env.cjs so this cache.version salt can't drift from
    // buildLicenseHeader()'s own copy of the same strings (used for the banner).
    const appCopyright = process.env.APP_COPYRIGHT || defaultAppCopyright();
    const publisherUrl = process.env.PUBLISHER_URL || DEFAULT_PUBLISHER_URL;

    // watchOptions is a documented no-op for one-shot builds (see chunkConfig()
    // below), but sdkConfig() has no way to know at this point whether the
    // caller is about to run under `--watch` or not — so a bad
    // WATCH_AGGREGATE_TIMEOUT must never hard-fail a plain `npm run build`
    // over a setting that build doesn't even use. Warn and fall back instead
    // of propagating resolveWatchAggregateTimeout()'s throw (which stays a
    // throw for its own unit tests, where the bad-input contract is exactly
    // what's being verified).
    let watchAggregateTimeout;
    try {
        watchAggregateTimeout = resolveWatchAggregateTimeout(process.env.WATCH_AGGREGATE_TIMEOUT);
    } catch (err) {
        console.warn(`sdkConfig: ${err.message} — falling back to the default (500ms)`);
        watchAggregateTimeout = 500;
    }

    const licenseText = buildLicenseHeader(__dirname);

    function chunkConfig(chunk, outName) {
        return {
            name: `${moduleName}:${chunk}`,
            mode: env,

            entry: {
                [outName]: DUMMY_ENTRY,
            },

            output: {
                path: OUT_DIR,
                filename: '[name].js',
                // publicPath intentionally unset: safe today because there's no code
                // splitting or dynamic import() in this config, so nothing needs to
                // resolve a chunk/asset URL at runtime. Set it (matching DocumentServer's
                // non-root deployment path) if splitChunks or import() is ever added here.
                // iife:false — we control wrapping via the loader:
                //   sdk-all-min: no wrapper
                //   sdk-all:     (function(window, undefined){…})(window)
                // Letting webpack add its own ()=>{} on top would still work
                // (code sets window.xxx), but iife:false gives a cleaner output.
                //
                // WARNING: ~287 places across sdkjs (GlobalSkin, etc.) read bare globals
                // with no `window.X = ...` assignment anywhere — they rely on webpack
                // inlining this single-module, no-import entry at true top level (no
                // function wrapper), so bare `var`s land on `window` the same way a
                // plain <script> tag would. That inlining is a webpack optimization that
                // only kicks in for a module with no imports/splitChunks/externals. If
                // this config ever gains import(), splitChunks, or an external, webpack
                // switches to its wrapped bootstrap form and all 287 bare-global reads
                // break silently at runtime — with a green build and no warning. Re-verify
                // this before adding any of those.
                iife: false,
                // Multiple chunk configs share OUT_DIR; do not wipe sibling output.
                clean: false,
            },

            module: {
                rules: [
                    {
                        // Match only our dummy entry, not real source files. webpack's
                        // string `test` does a startsWith() match, not equality, so an
                        // unanchored match would also catch e.g. a future dummy.json or
                        // dummy.js.bak, silently routing it through sdk-concat-loader
                        // and corrupting the bundle. Use exact equality instead.
                        test: (resourcePath) => resourcePath === DUMMY_ENTRY,
                        use: [
                            {
                                loader: CONCAT_LOADER,
                                options: {
                                    module:    moduleName,
                                    chunk,
                                    platform,
                                    srcRoot:   SRC_ROOT,
                                    addonDirs,
                                    // Passed through so the loader can patch the
                                    // window.AscCommon.g_c* runtime values after
                                    // commonDefines.js's own hardcoded assignments —
                                    // see the buildMeta comment in sdk-concat.cjs.
                                    buildMeta: { companyName, version, buildNumber, beta },
                                },
                            },
                        ],
                    },
                ],
            },

            plugins: [
                // Runs right after minification (OPTIMIZE_SIZE), because esbuild drops
                // every comment (legalComments:'none' below), banner included. Running
                // afterwards also means the ~400 per-file AGPL headers are already gone,
                // so the bundle carries exactly one license header.
                new webpack.BannerPlugin({
                    banner: licenseText,
                    raw: true,
                    entryOnly: true,
                    stage: webpack.Compilation.PROCESS_ASSETS_STAGE_OPTIMIZE_SIZE + 1,
                }),

                // scanLimit derives from the actual injected banner so a future
                // license.header edit that grows the banner can't silently push
                // webpack's "use strict"; directive past the scan window — see
                // the PROLOGUE_SCAN_LIMIT comment above.
                new StripBundlePostprocessPlugin({ scanLimit: licenseText.length + PROLOGUE_SCAN_LIMIT }),

                // Replaces Closure Compiler's --define= flags.
                // webpack DefinePlugin performs AST-level identifier replacement
                // so dead-code branches (if (g_cIsBeta === 'true') …) are
                // eliminated by the minifier (esbuild minifySyntax) in the same pass.
                //
                // Only the unprefixed form `AscCommon.g_cXxx` is listed here.
                // The `window.AscCommon.g_cXxx = "..."` declarations in
                // commonDefines.js must NOT be replaced — DefinePlugin would
                // turn the LHS into a string literal, making it an invalid
                // assignment. Those declarations stay as hardcoded placeholders
                // in source; sdk-concat-loader patches them to the real values
                // (buildMeta option above) right after commonDefines.js in the
                // 'min' chunk, so window.AscCommon.g_cXxx is correct too — not
                // just call-sites that get folded by this plugin.
                new webpack.DefinePlugin({
                    'AscCommon.g_cCompanyName':    JSON.stringify(companyName),
                    'AscCommon.g_cProductVersion': JSON.stringify(version),
                    'AscCommon.g_cBuildNumber':    JSON.stringify(buildNumber),
                    'AscCommon.g_cIsBeta':         JSON.stringify(beta),
                }),
            ],

            optimization: {
                minimize: env === 'production',
                minimizer: [
                    // esbuild via terser-webpack-plugin's adapter: ~4x faster and half the
                    // peak memory of Terser on these multi-MB concatenated bundles, at about
                    // the same output size.
                    new TerserPlugin({
                        minify: TerserPlugin.esbuildMinify,
                        extractComments: false,
                        // build-pipeline.cjs spawns 4 webpack-cli processes concurrently
                        // (word/cell/slide/visio), each running 2 chunk configs, so up to 8
                        // minifier worker pools at once. Left at the library default
                        // (parallel:true, sized to os.cpus().length), that's 8x
                        // over-subscription on small CI runners. Cap each pool to a
                        // quarter of the core count so total workers stay in the
                        // same ballpark as the host's core count.
                        parallel: Math.max(1, Math.floor(os.cpus().length / 4)),
                        // Passed to esbuild.transform() as-is (the adapter only adds
                        // sourcemap and sourcefile).
                        terserOptions: {
                            // minify:false plus the three flags below, NOT minify:true with
                            // minifyIdentifiers:false: with minify:true esbuild still renamed
                            // locals. Top-level names must never be renamed: sdkjs files
                            // communicate via window.AscCommon.xxx and bare top-level var
                            // declarations shared across the concatenated scope (the same
                            // reason the Terser build had mangle:false). esbuild still numbers
                            // a local that shadows an outer one (`window` -> `window2`, nested
                            // `function CShape` -> `CShape2`), which changes Function.name of
                            // such nested declarations; nothing in sdkjs relies on it beyond
                            // debug logging, and there is no option to turn it off.
                            minify: false,
                            minifyWhitespace: true,
                            minifyIdentifiers: false,
                            // Cap the syntax esbuild may introduce while minifying (e.g.
                            // `a?.b`, `??`, `catch{}`) at the documented ES2022 floor (see
                            // build/DEVELOPER-GUIDE.md). Source already at or below it ships as is.
                            target: 'es2022',
                            // Old build-desktop.bat/build-mobile.command ran Closure's
                            // WHITESPACE_ONLY (comments/whitespace stripped, no semantic
                            // transforms) instead of web's ADVANCED. Whitespace-only esbuild
                            // keeps that lighter tier on desktop/mobile.
                            minifySyntax: !desktopOrMobile,
                            // ASCII-only output (#80). doctrenderer/x2t run these bundles on a
                            // V8 built with v8_enable_i18n_support=false, whose reduced Unicode
                            // tables do not classify supplementary-plane ("astral") characters
                            // as ID_Start. The LaTeX symbol table in word/Math/NamesOfLiterals.js
                            // keys an object on such characters.
                            //
                            // The Terser build needed format.quote_keys on top of ascii_only,
                            // because Terser emitted those keys bare as `\u{1d552}:` which is
                            // pure ASCII yet still a bare astral identifier. esbuild has no
                            // quote_keys option: it unquotes plain keys, but with charset
                            // 'ascii' it keeps astral keys quoted as `"\u{1D552}":` (131 keys
                            // in word sdk-all.js), which is what #80 needs.
                            // build/scripts/check-bundle-ascii.cjs (run in check-build.yml) and
                            // build/test/webpack-sdk-terser-options.test.cjs guard both. Not yet
                            // verified end to end: `x2t -create-js-cache` on esbuild output.
                            //
                            // esbuild does not escape non-ASCII inside regex literals; such
                            // ranges must be written as \uXXXX escapes in source (see
                            // common/editorscommon.js), or the bundle scan fails.
                            charset: 'ascii',
                            // Strip every comment, including /*! and @license ones; the
                            // single banner is added after minification (BannerPlugin above).
                            legalComments: 'none',
                            // The legacy Closure Compiler build did not drop console
                            // calls, and sdkjs uses console.* for non-debug diagnostics
                            // (invalid-JS errors, clipboard permission warnings, custom
                            // function registration warnings, workbook diagnostics),
                            // silently discarding those in production removes real
                            // observability and can skip evaluation of their arguments.
                            // Opt in explicitly per-build instead of dropping by default.
                            // Like Terser's drop_console before, never on desktop/mobile.
                            drop: (!desktopOrMobile && process.env.DROP_CONSOLE === '1') ? ['console'] : [],
                        },
                    }),
                ],
            },

            // The SDK bundle intentionally exceeds webpack's 244 KiB default limit —
            // this is not a web-app chunk. Scoped to just this entry's output file
            // (rather than `performance: false` for the whole config) so a future
            // entry added to this config still gets the default size-regression hint.
            performance: {
                assetFilter: assetFilename => assetFilename !== `${outName}.js` && !assetFilename.endsWith('.map'),
            },

            // Persistent disk cache: cold restarts after the first build run in
            // ~0.6 s instead of ~34 s (production) or ~5.5 s (development).
            // Automatically invalidated when source files, configs, or this
            // factory file change via the registered addDependency() calls.
            cache: {
                type: 'filesystem',
                // Overridable for read-only source checkouts/mounts, where
                // build/.webpack-cache can't be created.
                cacheDirectory: process.env.WEBPACK_CACHE_DIR
                    ? path.resolve(process.env.WEBPACK_CACHE_DIR)
                    : path.join(__dirname, '.webpack-cache'),
                // DefinePlugin/BannerPlugin/devtool values below aren't tracked as
                // file dependencies, so a run that only changes an env var (e.g.
                // PRODUCT_VERSION or SDK_SOURCE_MAPS) must bump the cache version
                // itself or the previous run's bundle would be served unchanged.
                //
                // platform/addonDirs MUST be included too: they're passed as loader
                // *options*, not file dependencies, and WEBPACK_CACHE_DIR defaults to
                // the same fixed build/.webpack-cache path regardless of SDK_PLATFORM
                // or SDK_ADDONS. Verified empirically — without this, building once
                // with SDK_PLATFORM=desktop (or SDK_ADDONS=...) after a plain build
                // against the same cache dir silently reuses the plain build's cached
                // module and produces byte-identical output missing every
                // platform/addon file, with no error or warning.
                //
                // `mode`/NODE_ENV deliberately does NOT need to be in this key.
                // Verified empirically: production (SDK_SOURCE_MAPS=1) -> development
                // -> production again against the same cache dir produces byte-identical,
                // correctly-minified output on the second production run, served from
                // `[cached] modules`. The minifier only minifies the *output* of the cached
                // module (per optimization.minimize, gated on `mode`) as a separate
                // build-time step every run — it is never itself part of what gets
                // cached, so a mode switch can't serve a stale minified/unminified asset.
                version: `${companyName}-${version}-${buildNumber}-${beta}-${appCopyright}-${publisherUrl}-${emitSourceMaps}-${platform}-${addonDirs.join(',')}-${process.env.DROP_CONSOLE || ''}`,
                buildDependencies: {
                    config: [
                        fileURLToPath(import.meta.url),
                        CONCAT_LOADER,
                        path.join(__dirname, 'lib', 'sdk-configs.cjs'),
                        path.join(__dirname, 'lib', 'env.cjs'),
                        path.join(__dirname, 'license.header'),
                    ],
                },
            },

            devtool: emitSourceMaps ? 'source-map' : false,

            // Only takes effect under `--watch` (npm run watch:*); ignored for
            // one-shot builds. webpack 5's default aggregateTimeout is 20ms
            // (see Watching.js) — too short for an editor's save to fully land
            // on disk before the loader re-reads it. VS Code (and other
            // editors) can flush a single save as more than one MODIFY event,
            // especially over a 9p/WSL2 filesystem boundary; the first event
            // triggers a rebuild while sdk-concat-loader is re-concatenating
            // ~400+ source files, and it can catch this file's write
            // mid-flight, truncated. That produces the intermittent tiny
            // sdk-all.js / parse-error rebuild described in issue #78, always
            // followed by a second, correct rebuild once the write settles.
            //
            // This only debounces the *trigger* — sdk-concat.cjs still does a
            // plain readFile with no size/mtime stability check once the
            // timeout fires, so it reduces the race window (20ms -> 500ms of
            // required quiescence) rather than closing it structurally. A
            // write that's still landing after 500ms of silence (slow disk,
            // network mount) could still be read mid-flight. Overridable via
            // WATCH_AGGREGATE_TIMEOUT, matching this file's other env-tunable
            // knobs, in case 500ms proves insufficient on a given machine/CI
            // watcher.
            watchOptions: {
                aggregateTimeout: watchAggregateTimeout,
            },
        };
    }

    return [
        chunkConfig('min', 'sdk-all-min'),
        chunkConfig('all', 'sdk-all'),
    ];
}
