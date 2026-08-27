/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

/**
 * Snapshot/assertion coverage for sdkConfig()'s per-platform Terser options
 * (build/webpack.sdk.factory.mjs's chunkConfig()). These exact knobs were
 * what commit f787b365b8 fixed to solve real bundle-bloat and duplicate
 * license-header bugs, but had zero test coverage — a future refactor could
 * silently invert the desktop/mobile `compress` branch (re-enabling dead-code
 * elimination on platforms that intentionally disable it) or drop the
 * `mangle:false` invariant sdkjs's bare-global/shared-scope model depends on,
 * with nothing here to catch it.
 */

'use strict';

const test   = require('node:test');
const assert = require('node:assert/strict');
const path   = require('node:path');
const url    = require('node:url');

async function loadSdkConfig() {
    const mod = await import(url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs')));
    return mod.sdkConfig;
}

// sdkConfig() reads SDK_PLATFORM from process.env at call time (not at module
// load time), so tests can drive it directly as long as they restore it
// afterwards for any other test relying on the default ('').
function withPlatform(platform, fn) {
    const prev = process.env.SDK_PLATFORM;
    if (platform) {
        process.env.SDK_PLATFORM = platform;
    } else {
        delete process.env.SDK_PLATFORM;
    }
    try {
        return fn();
    } finally {
        if (prev === undefined) {
            delete process.env.SDK_PLATFORM;
        } else {
            process.env.SDK_PLATFORM = prev;
        }
    }
}

function terserOptionsOf(chunkConfigs) {
    // chunkConfigs is [minConfig, allConfig] from sdkConfig(); both chunks
    // share the same platform-derived terserOptions, so either is representative.
    return chunkConfigs[0].optimization.minimizer[0].options.minimizer.options;
}

test('sdkConfig: web platform (SDK_PLATFORM unset) enables Terser compress', async () => {
    const sdkConfig = await loadSdkConfig();
    const terserOptions = withPlatform('', () => terserOptionsOf(sdkConfig('word')));

    assert.notEqual(terserOptions.compress, false);
    assert.equal(typeof terserOptions.compress, 'object');
    assert.equal(terserOptions.mangle, false);
});

test('sdkConfig: desktop platform disables Terser compress (restores legacy WHITESPACE_ONLY-equivalent behavior)', async () => {
    const sdkConfig = await loadSdkConfig();
    const terserOptions = withPlatform('desktop', () => terserOptionsOf(sdkConfig('word')));

    assert.equal(terserOptions.compress, false);
    assert.equal(terserOptions.mangle, false);
});

test('sdkConfig: mobile platform disables Terser compress (same as desktop)', async () => {
    const sdkConfig = await loadSdkConfig();
    const terserOptions = withPlatform('mobile', () => terserOptionsOf(sdkConfig('word')));

    assert.equal(terserOptions.compress, false);
    assert.equal(terserOptions.mangle, false);
});

test('sdkConfig: comments format only preserves the license-banner sentinel, not every per-file AGPL header', async () => {
    const sdkConfig = await loadSdkConfig();
    const terserOptions = withPlatform('', () => terserOptionsOf(sdkConfig('word')));

    const commentsRegex = terserOptions.format.comments;
    assert.equal(commentsRegex.test('@@license-banner@@'), true);
    // A per-file AGPL header (repeated ~400+ times pre-minification) must NOT
    // match, or the duplicate-header bloat commit f787b365b8 fixed would regress.
    assert.equal(commentsRegex.test('This program is a free software product'), false);
});

test('sdkConfig: DROP_CONSOLE is opt-in, not the default, for non-desktop/mobile platforms', async () => {
    const sdkConfig = await loadSdkConfig();
    const prevDropConsole = process.env.DROP_CONSOLE;
    delete process.env.DROP_CONSOLE;
    try {
        const terserOptions = withPlatform('', () => terserOptionsOf(sdkConfig('word')));
        assert.equal(terserOptions.compress.drop_console, false);
    } finally {
        if (prevDropConsole === undefined) {
            delete process.env.DROP_CONSOLE;
        } else {
            process.env.DROP_CONSOLE = prevDropConsole;
        }
    }
});

// --- Regression coverage for issue #80 ---------------------------------------
// doctrenderer/x2t execute these bundles on a V8 built with
// v8_enable_i18n_support=false, whose reduced Unicode tables do not accept
// supplementary-plane ("astral") characters as identifiers. The webpack
// pipeline emitted the LaTeX symbol table's astral keys as bare identifiers
// (the reverse map in word/Math/NamesOfLiterals.js), and that V8 rejected the
// whole bundle with "SyntaxError: Invalid or unexpected token" -- aborting x2t
// on first start and making Euro-Office 9.3.4 unusable. ASCII-only output is
// the invariant the previous Closure Compiler build provided implicitly.

for (const platform of ['', 'desktop', 'mobile']) {
    test(`sdkConfig: Terser is configured for ASCII-only output on platform '${platform || 'web'}' (#80)`, async () => {
        const sdkConfig = await loadSdkConfig();
        const terserOptions = withPlatform(platform, () => terserOptionsOf(sdkConfig('word')));

        assert.equal(terserOptions.format.ascii_only, true);
        assert.equal(terserOptions.format.quote_keys, true);
    });
}

test('sdkConfig: Terser escapes astral object keys rather than emitting them bare (#80)', async () => {
    const terser     = require('terser');
    const sdkConfig  = await loadSdkConfig();
    const terserOptions = withPlatform('', () => terserOptionsOf(sdkConfig('word')));

    // Shape lifted from word/Math/NamesOfLiterals.js's reverse symbol table:
    // U+2219 (BMP) plus U+1D552 / U+1D538 (astral) used as object keys.
    const src = 'var Reverse={"∙":"\\\\bullet","𝕒":"\\\\doublea","𝔸":"\\\\doubleA"};';
    const { code } = await terser.minify(src, terserOptions);

    assert.ok(!/[^\x00-\x7F]/.test(code),
        `Terser emitted non-ASCII output, which doctrenderer's no-ICU V8 may reject: ${code}`);
    assert.equal(/[{,]\s*[\u{10000}-\u{10FFFF}]/u.test(code), false,
        `Terser emitted a bare astral object key, which doctrenderer's no-ICU V8 rejects: ${code}`);
});
