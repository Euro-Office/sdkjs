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

// @@license-banner@@ exists only so Terser's format.comments regex (below) can tell
// the single BannerPlugin-injected banner apart from the ~400 identical per-file AGPL
// headers it would otherwise also match. It has to survive in the bundle text through
// the Terser pass for that match to work, so it can't be stripped from licenseText
// before injection — instead, strip it from the asset after minification is done.
//
// Combines both the strict-mode-directive strip and the license-sentinel strip into a
// single processAssets pass: each is a separate transform on the same asset, and doing
// them as two independent plugins would materialize/scan/updateAsset the full bundle
// source twice per chunk (multi-MB for sdk-all, across 4 modules x 2 chunks per build).
//
// Exported (only) for the real-compilation integration test in
// build/test/webpack-sdk-factory.test.cjs — this plugin's correctness depends on
// webpack's own internal bootstrap-generation format and Terser's output formatting,
// neither of which is a stable public API, so it needs an actual webpack+Terser run to
// validate, not just the pure-string-function tests above. The constructor flags let
// that test exercise each transform in isolation without giving up the single-pass shape
// used in production (sdkConfig() below always enables both).
export class StripBundlePostprocessPlugin {
    constructor({ stripStrictMode = true, stripLicenseSentinel = true, scanLimit = PROLOGUE_SCAN_LIMIT } = {}) {
        this.stripStrictMode = stripStrictMode;
        this.stripLicenseSentinel = stripLicenseSentinel;
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

                        let patched = source;
                        if (this.stripStrictMode) {
                            patched = stripBootstrapStrictDirective(patched, this.scanLimit);
                        }
                        if (this.stripLicenseSentinel && patched.includes('@@license-banner@@')) {
                            patched = patched.replace(/\s?@@license-banner@@/, '');
                        }
                        if (patched === source) continue;

                        compilation.updateAsset(name, new webpack.sources.RawSource(patched));
                    }
                }
            );
        });
    }
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

    // Sentinel deliberately kept in (stripSentinel: false, the default) — Terser's
    // format.comments regex below matches on it, and StripBundlePostprocessPlugin
    // strips it from the asset only after that Terser pass runs.
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
                new webpack.BannerPlugin({
                    banner: licenseText,
                    raw: true,
                    entryOnly: true,
                }),

                // scanLimit derives from the actual injected banner so a future
                // license.header edit that grows the banner can't silently push
                // webpack's "use strict"; directive past the scan window — see
                // the PROLOGUE_SCAN_LIMIT comment above.
                new StripBundlePostprocessPlugin({ scanLimit: licenseText.length + PROLOGUE_SCAN_LIMIT }),

                // Replaces Closure Compiler's --define= flags.
                // webpack DefinePlugin performs AST-level identifier replacement
                // so dead-code branches (if (g_cIsBeta === 'true') …) are
                // eliminated by TerserPlugin in the same pass.
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
                    new TerserPlugin({
                        extractComments: false,
                        // build-pipeline.cjs spawns 4 webpack-cli processes concurrently
                        // (word/cell/slide/visio), each running 2 chunk configs — up to 8
                        // TerserPlugin worker pools at once. Left at the library default
                        // (parallel:true, sized to os.cpus().length), that's 8x
                        // over-subscription on small CI runners. Cap each pool to a
                        // quarter of the core count so total worker threads stay in the
                        // same ballpark as the host's core count.
                        parallel: Math.max(1, Math.floor(os.cpus().length / 4)),
                        terserOptions: {
                            format: {
                                // BannerPlugin (stage ADDITIONS) injects the license banner into
                                // the asset *before* Terser (stage OPTIMIZE_SIZE) runs, so the
                                // banner is just another comment to Terser at this point. Match
                                // only the sentinel below — matching on AGPL/Copyright/License
                                // text also matches the identical per-file header repeated in
                                // all ~400+ concatenated source files.
                                comments: /@@license-banner@@/,

                                // BOTH of the options below are required; neither is
                                // sufficient alone. doctrenderer/x2t run these bundles on a
                                // V8 built with v8_enable_i18n_support=false, whose reduced
                                // Unicode tables do not classify supplementary-plane
                                // ("astral") characters as ID_Start. The LaTeX symbol table
                                // in word/Math/NamesOfLiterals.js keys an object on such
                                // characters, and this pipeline emits them unquoted.
                                //
                                // ascii_only escapes non-ASCII *characters*, but an escaped
                                // bare key is still a bare astral identifier: with
                                // ascii_only alone this build emits 131 keys of the form
                                //     \u{1d552}:"\\doublea"
                                // which is pure ASCII text yet still resolves to U+1D552 for
                                // ID_Start classification, so that V8 rejects it exactly as
                                // it rejects the raw UTF-8 form. Measured: with ascii_only
                                // alone, `x2t -create-js-cache` still aborts and writes a
                                // 0-byte sdk-all.cache.
                                //
                                // quote_keys is therefore the load-bearing option -- it
                                // turns the key into a string literal, which ascii_only then
                                // escapes into `"\u{1d552}"`. ascii_only remains necessary
                                // in its own right so no raw non-ASCII byte reaches the
                                // bundle at all.
                                //
                                // quote_keys must live inside `format`; Terser rejects a
                                // top-level one with "`quote_keys` is not a supported
                                // option". Covered by
                                // build/test/webpack-sdk-terser-options.test.cjs and by the
                                // bundle scan in check-build.yml. See #80.
                                ascii_only: true,
                                quote_keys: true,
                            },
                            compress: (platform === 'desktop' || platform === 'mobile')
                                // Old build-desktop.bat/build-mobile.command ran Closure's
                                // WHITESPACE_ONLY (comments/whitespace stripped, no semantic
                                // transforms) instead of web's ADVANCED. `compress: false` is
                                // Terser's closest equivalent — it disables the whole compress
                                // pass (dead-code elim, inlining, etc.) and keeps output to
                                // whitespace/name-shortening-free minification, restoring that
                                // lighter tier instead of silently adopting web's ADVANCED-like
                                // pass on desktop/mobile.
                                ? false
                                : {
                                    // The legacy Closure Compiler build did not drop console
                                    // calls, and sdkjs uses console.* for non-debug diagnostics
                                    // (invalid-JS errors, clipboard permission warnings, custom
                                    // function registration warnings, workbook diagnostics) —
                                    // silently discarding those in production removes real
                                    // observability and can skip evaluation of their arguments.
                                    // Opt in explicitly per-build instead of dropping by default.
                                    drop_console: process.env.DROP_CONSOLE === '1',
                                },
                            // mangle:false is load-bearing — same reason as web-apps:
                            // sdkjs files communicate via window.AscCommon.xxx and bare
                            // top-level var declarations shared across concatenated scope.
                            // Mangling property names would silently corrupt those references.
                            mangle: false,
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
                // `[cached] modules`. Terser only minifies the *output* of the cached
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
        };
    }

    return [
        chunkConfig('min', 'sdk-all-min'),
        chunkConfig('all', 'sdk-all'),
    ];
}
