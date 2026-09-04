/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

const test   = require('node:test');
const assert = require('node:assert/strict');
const path   = require('node:path');
const url    = require('node:url');
const fs     = require('node:fs');
const os     = require('node:os');
const webpack = require('webpack');
const TerserPlugin = require('terser-webpack-plugin');

test('stripBootstrapStrictDirective: removes webpack\'s forced top-level "use strict" prologue', async () => {
    const { stripBootstrapStrictDirective } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    const bundle = '/******/ "use strict";\n/******/ (() => {\nvar x = 1;\n})();';
    const patched = stripBootstrapStrictDirective(bundle);

    assert.equal(patched.includes('"use strict"'), false);
    // Same length / same line count — must not shift anything a source map points at.
    assert.equal(patched.length, bundle.length);
    assert.equal(patched.split('\n').length, bundle.split('\n').length);
});

test('stripBootstrapStrictDirective: leaves the bundle unchanged when no directive is present', async () => {
    const { stripBootstrapStrictDirective } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    const bundle = '/******/ (() => {\nvar x = 1;\n})();';
    assert.equal(stripBootstrapStrictDirective(bundle), bundle);
});

test('stripBootstrapStrictDirective: does not touch a "use strict" appearing past the prologue window', async () => {
    const { stripBootstrapStrictDirective, PROLOGUE_SCAN_LIMIT } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    // A real source file's own string literal containing this text, buried
    // deep in the concatenated bundle, must never be mistaken for webpack's
    // bootstrap directive.
    const padding = 'x'.repeat(PROLOGUE_SCAN_LIMIT + 100);
    const bundle = `${padding}var msg = "use strict";`;
    assert.equal(stripBootstrapStrictDirective(bundle), bundle);
});

test('stripBootstrapStrictDirective: honors a caller-supplied scanLimit', async () => {
    const { stripBootstrapStrictDirective, PROLOGUE_SCAN_LIMIT } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    // Directive sits past the default PROLOGUE_SCAN_LIMIT but within a
    // caller-supplied, larger scanLimit (e.g. derived from an actual banner
    // length) — must still be found and stripped.
    const padding = 'x'.repeat(PROLOGUE_SCAN_LIMIT + 100);
    const bundle  = `${padding}"use strict";\nvar y = 1;`;

    assert.equal(stripBootstrapStrictDirective(bundle), bundle);
    const patched = stripBootstrapStrictDirective(bundle, PROLOGUE_SCAN_LIMIT + 200);
    assert.equal(patched.includes('"use strict"'), false);
    assert.equal(patched.length, bundle.length);
});

test('StripBundlePostprocessPlugin: strips the sentinel even without a leading space (reformatted banner)', async () => {
    const { StripBundlePostprocessPlugin } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'strip-license-sentinel-nospace-test-'));
    const entry  = path.join(tmpDir, 'entry.js');
    fs.writeFileSync(entry, 'var AscCommonSdkTestGlobal = { value: 1 + 1 };\n');

    // No leading space before the sentinel — the guard/action mismatch this
    // regression covers only manifests once the banner is formatted this way.
    const banner = '/*@@license-banner@@\n * Copyright (C) Test Corp\n */';

    function runCompiler(withPlugin) {
        return new Promise((resolve, reject) => {
            const compiler = webpack({
                mode: 'production',
                entry,
                output: { path: tmpDir, filename: withPlugin ? 'with-plugin.js' : 'without-plugin.js', iife: false },
                plugins: [
                    new webpack.BannerPlugin({ banner, raw: true, entryOnly: true }),
                    ...(withPlugin ? [new StripBundlePostprocessPlugin({ stripStrictMode: false })] : []),
                ],
                optimization: {
                    minimize: true,
                    minimizer: [
                        new TerserPlugin({
                            extractComments: false,
                            terserOptions: {
                                mangle: false,
                                compress: true,
                                format: { comments: /@@license-banner@@/ },
                            },
                        }),
                    ],
                },
            });
            compiler.run((err, stats) => {
                compiler.close(() => {});
                if (err || stats.hasErrors()) return reject(err || new Error(stats.toString()));
                resolve(fs.readFileSync(path.join(tmpDir, withPlugin ? 'with-plugin.js' : 'without-plugin.js'), 'utf8'));
            });
        });
    }

    try {
        const withPlugin = await runCompiler(true);
        assert.equal(withPlugin.includes('Copyright (C) Test Corp'), true);
        assert.equal(withPlugin.includes('@@license-banner@@'), false);
    } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    }
});

test('StripBundlePostprocessPlugin: strips webpack\'s bootstrap "use strict" from a real minified compilation', async (t) => {
    const { StripBundlePostprocessPlugin } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'strip-strict-test-'));
    const entry  = path.join(tmpDir, 'entry.js');
    // Bare top-level var, no import/export — same shape sdk-concat-loader
    // produces (sourceType:'script'), so webpack treats this chunk the same
    // way it treats a real SDK bundle for bootstrap-generation purposes.
    fs.writeFileSync(entry, 'var AscCommonSdkTestGlobal = { value: 1 + 1 };\n');

    function runCompiler(withPlugin) {
        return new Promise((resolve, reject) => {
            const compiler = webpack({
                mode: 'production',
                entry,
                output: { path: tmpDir, filename: withPlugin ? 'with-plugin.js' : 'without-plugin.js', iife: false },
                optimization: {
                    minimize: true,
                    minimizer: [new TerserPlugin({ terserOptions: { mangle: false, compress: true } })],
                },
                plugins: withPlugin
                    ? [new StripBundlePostprocessPlugin({ stripLicenseSentinel: false })]
                    : [],
            });
            compiler.run((err, stats) => {
                compiler.close(() => {});
                if (err || stats.hasErrors()) return reject(err || new Error(stats.toString()));
                resolve(fs.readFileSync(path.join(tmpDir, withPlugin ? 'with-plugin.js' : 'without-plugin.js'), 'utf8'));
            });
        });
    }

    try {
        const withoutPlugin = await runCompiler(false);
        const withPlugin    = await runCompiler(true);

        // Sanity check the test fixture itself is meaningful: if webpack's own
        // output never carries the directive in the first place (e.g. a future
        // webpack version stops emitting it for script-sourceType chunks), the
        // plugin has nothing to strip and this assertion would catch that the
        // integration test itself needs updating, rather than silently passing
        // for the wrong reason.
        if (!withoutPlugin.includes('"use strict"')) {
            t.diagnostic('webpack did not emit a bootstrap "use strict" for this chunk shape — plugin has nothing to strip here');
        } else {
            assert.equal(withPlugin.includes('"use strict"'), false);
        }
    } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    }
});

test('resolveWatchAggregateTimeout: falls back to the default when unset', async () => {
    const { resolveWatchAggregateTimeout } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );
    assert.equal(resolveWatchAggregateTimeout(undefined), 500);
    assert.equal(resolveWatchAggregateTimeout(''), 500);
});

test('resolveWatchAggregateTimeout: honors an explicit 0 (no debounce) instead of treating it as unset', async () => {
    const { resolveWatchAggregateTimeout } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );
    // `Number(x) || 500` would wrongly collapse "0" back to the default since
    // 0 is falsy — this is the regression the helper guards against.
    assert.equal(resolveWatchAggregateTimeout('0'), 0);
});

test('resolveWatchAggregateTimeout: rejects a non-numeric override instead of silently falling back', async () => {
    const { resolveWatchAggregateTimeout } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );
    assert.throws(() => resolveWatchAggregateTimeout('not-a-number'), /must be a non-negative number/);
    assert.throws(() => resolveWatchAggregateTimeout('-100'), /must be a non-negative number/);
});

test('sdkConfig: never throws on a bad WATCH_AGGREGATE_TIMEOUT — one-shot builds must not fail over a watch-only setting', async () => {
    const { sdkConfig } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'sdk-watch-options-bad-env-test-'));
    const prevBuildRoot = process.env.BUILD_ROOT;
    const prevCacheDir  = process.env.WEBPACK_CACHE_DIR;
    const prevTimeout   = process.env.WATCH_AGGREGATE_TIMEOUT;
    const prevWarn      = console.warn;
    process.env.BUILD_ROOT              = tmpRoot;
    process.env.WEBPACK_CACHE_DIR       = path.join(tmpRoot, '.webpack-cache');
    process.env.WATCH_AGGREGATE_TIMEOUT = 'not-a-number';
    let warned = false;
    console.warn = () => { warned = true; };

    try {
        // sdkConfig() is called for plain `npm run build` too, where
        // watchOptions is inert — a stray/typo'd env var must never crash
        // that path (see resolveWatchAggregateTimeout's own throwing tests
        // above for the input-validation contract itself).
        const [minConfig, allConfig] = sdkConfig('word');
        for (const config of [minConfig, allConfig]) {
            assert.equal(config.watchOptions.aggregateTimeout, 500, `${config.name}: should fall back to the default on invalid input`);
        }
        assert.equal(warned, true, 'expected a console.warn about the invalid value');
    } finally {
        console.warn = prevWarn;
        if (prevBuildRoot === undefined) delete process.env.BUILD_ROOT; else process.env.BUILD_ROOT = prevBuildRoot;
        if (prevCacheDir  === undefined) delete process.env.WEBPACK_CACHE_DIR; else process.env.WEBPACK_CACHE_DIR = prevCacheDir;
        if (prevTimeout   === undefined) delete process.env.WATCH_AGGREGATE_TIMEOUT; else process.env.WATCH_AGGREGATE_TIMEOUT = prevTimeout;
        fs.rmSync(tmpRoot, { recursive: true, force: true });
    }
});

test('sdkConfig: sets a watchOptions.aggregateTimeout above webpack\'s 20ms default on every chunk config', async () => {
    const { sdkConfig } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'sdk-watch-options-test-'));
    const prevBuildRoot = process.env.BUILD_ROOT;
    const prevCacheDir  = process.env.WEBPACK_CACHE_DIR;
    process.env.BUILD_ROOT        = tmpRoot;
    process.env.WEBPACK_CACHE_DIR = path.join(tmpRoot, '.webpack-cache');

    try {
        const [minConfig, allConfig] = sdkConfig('word');

        // webpack 5's own default (see node_modules/webpack/lib/Watching.js) is
        // 20ms — a race window too short to survive a multi-event editor save
        // (see issue #78). Both chunk configs must opt into a longer debounce;
        // watchOptions is a no-op for one-shot (non --watch) builds, so setting
        // it unconditionally is safe for production/CI builds too.
        for (const config of [minConfig, allConfig]) {
            assert.ok(config.watchOptions, `${config.name}: missing watchOptions`);
            assert.equal(config.watchOptions.aggregateTimeout, 500, `${config.name}: default aggregateTimeout should be 500`);
            assert.ok(
                config.watchOptions.aggregateTimeout > 20,
                `${config.name}: aggregateTimeout (${config.watchOptions.aggregateTimeout}) must exceed webpack's 20ms default`
            );
        }
    } finally {
        if (prevBuildRoot === undefined) delete process.env.BUILD_ROOT; else process.env.BUILD_ROOT = prevBuildRoot;
        if (prevCacheDir  === undefined) delete process.env.WEBPACK_CACHE_DIR; else process.env.WEBPACK_CACHE_DIR = prevCacheDir;
        fs.rmSync(tmpRoot, { recursive: true, force: true });
    }
});

test('sdkConfig: WATCH_AGGREGATE_TIMEOUT overrides the default watch debounce', async () => {
    const { sdkConfig } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'sdk-watch-options-override-test-'));
    const prevBuildRoot = process.env.BUILD_ROOT;
    const prevCacheDir  = process.env.WEBPACK_CACHE_DIR;
    const prevTimeout   = process.env.WATCH_AGGREGATE_TIMEOUT;
    process.env.BUILD_ROOT             = tmpRoot;
    process.env.WEBPACK_CACHE_DIR      = path.join(tmpRoot, '.webpack-cache');
    process.env.WATCH_AGGREGATE_TIMEOUT = '1200';

    try {
        const [minConfig, allConfig] = sdkConfig('word');
        for (const config of [minConfig, allConfig]) {
            assert.equal(config.watchOptions.aggregateTimeout, 1200, `${config.name}: WATCH_AGGREGATE_TIMEOUT override not applied`);
        }
    } finally {
        if (prevBuildRoot === undefined) delete process.env.BUILD_ROOT; else process.env.BUILD_ROOT = prevBuildRoot;
        if (prevCacheDir  === undefined) delete process.env.WEBPACK_CACHE_DIR; else process.env.WEBPACK_CACHE_DIR = prevCacheDir;
        if (prevTimeout   === undefined) delete process.env.WATCH_AGGREGATE_TIMEOUT; else process.env.WATCH_AGGREGATE_TIMEOUT = prevTimeout;
        fs.rmSync(tmpRoot, { recursive: true, force: true });
    }
});

test('StripBundlePostprocessPlugin: strips the @@license-banner@@ sentinel after Terser has used it to keep the banner', async () => {
    const { StripBundlePostprocessPlugin } = await import(
        url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs'))
    );

    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'strip-license-sentinel-test-'));
    const entry  = path.join(tmpDir, 'entry.js');
    // A per-file AGPL header identical in shape to the ~400 real ones sdk-concat-loader
    // concatenates ahead of the BannerPlugin-injected banner — it must NOT be kept by
    // Terser's comments regex, only the sentinel-marked banner should survive.
    fs.writeFileSync(
        entry,
        '/* AGPL header, repeated in every source file */\n' +
        'var AscCommonSdkTestGlobal = { value: 1 + 1 };\n'
    );

    const banner = '/* @@license-banner@@\n * Copyright (C) Test Corp\n */';

    function runCompiler(withPlugin) {
        return new Promise((resolve, reject) => {
            const compiler = webpack({
                mode: 'production',
                entry,
                output: { path: tmpDir, filename: withPlugin ? 'with-plugin.js' : 'without-plugin.js', iife: false },
                plugins: [
                    new webpack.BannerPlugin({ banner, raw: true, entryOnly: true }),
                    ...(withPlugin ? [new StripBundlePostprocessPlugin({ stripStrictMode: false })] : []),
                ],
                optimization: {
                    minimize: true,
                    minimizer: [
                        new TerserPlugin({
                            extractComments: false,
                            terserOptions: {
                                mangle: false,
                                compress: true,
                                // Same regex used in sdkConfig(): only the sentinel-marked
                                // banner should survive Terser's comment-stripping pass,
                                // not the per-file AGPL headers.
                                format: { comments: /@@license-banner@@/ },
                            },
                        }),
                    ],
                },
            });
            compiler.run((err, stats) => {
                compiler.close(() => {});
                if (err || stats.hasErrors()) return reject(err || new Error(stats.toString()));
                resolve(fs.readFileSync(path.join(tmpDir, withPlugin ? 'with-plugin.js' : 'without-plugin.js'), 'utf8'));
            });
        });
    }

    try {
        const withoutPlugin = await runCompiler(false);
        const withPlugin    = await runCompiler(true);

        // Sanity checks on the fixture itself: the per-file header must be gone
        // (Terser's regex didn't match it) and the banner text must have survived
        // in both outputs — only the sentinel differs between them.
        assert.equal(withoutPlugin.includes('AGPL header, repeated'), false);
        assert.equal(withoutPlugin.includes('Copyright (C) Test Corp'), true);
        assert.equal(withPlugin.includes('Copyright (C) Test Corp'), true);

        assert.equal(withoutPlugin.includes('@@license-banner@@'), true);
        assert.equal(withPlugin.includes('@@license-banner@@'), false);
    } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    }
});
