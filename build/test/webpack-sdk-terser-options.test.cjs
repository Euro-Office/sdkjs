/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

/**
 * Snapshot/assertion coverage for sdkConfig()'s per-platform minifier options
 * (build/webpack.sdk.factory.mjs's chunkConfig()). The minifier is esbuild,
 * run through terser-webpack-plugin's esbuildMinify adapter, so the options
 * below are esbuild TransformOptions. A future refactor could silently turn
 * on identifier renaming (breaking sdkjs's bare-global/shared-scope model),
 * invert the desktop/mobile minifySyntax branch, or lose ASCII-only output
 * (#80), with nothing here to catch it.
 */

const test   = require('node:test');
const assert = require('node:assert/strict');
const path   = require('node:path');
const url    = require('node:url');
const esbuild      = require('esbuild');
const TerserPlugin = require('terser-webpack-plugin');

// Share one definition of "bare astral key" with the CI bundle scan so the two
// cannot drift apart.
const { BARE_ASTRAL_RAW, BARE_ASTRAL_ESC } = require('../scripts/check-bundle-ascii.cjs');

async function loadSdkConfig() {
    const mod = await import(url.pathToFileURL(path.join(__dirname, '..', 'webpack.sdk.factory.mjs')));
    return mod.sdkConfig;
}

// sdkConfig() reads SDK_PLATFORM/DROP_CONSOLE from process.env at call time
// (not at module load time), so tests can drive it directly as long as they
// restore it afterwards.
function withEnv(name, value, fn) {
    const prev = process.env[name];
    if (value) {
        process.env[name] = value;
    } else {
        delete process.env[name];
    }
    try {
        return fn();
    } finally {
        if (prev === undefined) {
            delete process.env[name];
        } else {
            process.env[name] = prev;
        }
    }
}

function minimizerOf(chunkConfigs) {
    // chunkConfigs is [minConfig, allConfig] from sdkConfig(); both chunks
    // share the same platform-derived options, so either is representative.
    return chunkConfigs[0].optimization.minimizer[0].options.minimizer;
}

async function esbuildOptionsFor(platform) {
    const sdkConfig = await loadSdkConfig();
    return withEnv('SDK_PLATFORM', platform, () => minimizerOf(sdkConfig('word')).options);
}

test('sdkConfig: minifies with esbuild through terser-webpack-plugin', async () => {
    const sdkConfig = await loadSdkConfig();
    assert.equal(minimizerOf(sdkConfig('word')).implementation, TerserPlugin.esbuildMinify);
});

for (const platform of ['', 'desktop', 'mobile']) {
    test(`sdkConfig: esbuild never renames top-level or unshadowed identifiers on platform '${platform || 'web'}'`, async () => {
        const options = await esbuildOptionsFor(platform);

        // minify:true renames locals even with minifyIdentifiers:false, so both
        // must stay off. sdkjs's concatenated scripts share bare globals.
        assert.equal(options.minify, false);
        assert.equal(options.minifyIdentifiers, false);
        assert.equal(options.minifyWhitespace, true);
        assert.equal(options.legalComments, 'none');
        assert.equal(options.target, 'es2022');

        const { code } = await esbuild.transform(
            'function f() { var someLongLocalName = 1; return someLongLocalName; }', options);
        assert.match(code, /someLongLocalName/);
    });
}

test('sdkConfig: web platform (SDK_PLATFORM unset) enables esbuild minifySyntax', async () => {
    assert.equal((await esbuildOptionsFor('')).minifySyntax, true);
});

test('sdkConfig: desktop/mobile disable minifySyntax (legacy WHITESPACE_ONLY-equivalent tier)', async () => {
    assert.equal((await esbuildOptionsFor('desktop')).minifySyntax, false);
    assert.equal((await esbuildOptionsFor('mobile')).minifySyntax, false);
});

test('sdkConfig: dropping console is opt-in via DROP_CONSOLE=1, and never on desktop/mobile', async () => {
    const sdkConfig = await loadSdkConfig();
    const drop = (platform, dropConsole) => withEnv('DROP_CONSOLE', dropConsole,
        () => withEnv('SDK_PLATFORM', platform, () => minimizerOf(sdkConfig('word')).options.drop));

    assert.deepEqual(drop('', ''), []);
    assert.deepEqual(drop('', '1'), ['console']);
    assert.deepEqual(drop('desktop', '1'), []);
});

// --- Regression coverage for issue #80 ---------------------------------------
// doctrenderer/x2t execute these bundles on a V8 built with
// v8_enable_i18n_support=false, whose reduced Unicode tables do not accept
// supplementary-plane ("astral") characters as identifiers. The webpack
// pipeline once emitted the LaTeX symbol table's astral keys as bare
// identifiers (the reverse map in word/Math/NamesOfLiterals.js), and that V8
// rejected the whole bundle with "SyntaxError: Invalid or unexpected token".
// ASCII-only output with quoted astral keys is the invariant to keep.

// Shape lifted from word/Math/NamesOfLiterals.js's reverse symbol table:
// U+2219 (BMP) plus U+1D552 / U+1D538 (astral) used as object keys.
const SYMBOL_TABLE_SRC = 'var Reverse={"\u2219":"\\bullet","\u{1d552}":"\\doublea","\u{1d538}":"\\doubleA"};';

for (const platform of ['', 'desktop', 'mobile']) {
    test(`sdkConfig: esbuild emits ASCII-only output with quoted astral keys on platform '${platform || 'web'}' (#80)`, async () => {
        const options = await esbuildOptionsFor(platform);
        assert.equal(options.charset, 'ascii');

        const { code } = await esbuild.transform(SYMBOL_TABLE_SRC, options);

        assert.ok(!/[^\x00-\x7F]/.test(code),
            `esbuild emitted non-ASCII output, which doctrenderer's no-ICU V8 rejects: ${code}`);
        assert.equal(BARE_ASTRAL_RAW.test(code), false,
            `esbuild emitted a raw bare astral object key: ${code}`);
        assert.equal(BARE_ASTRAL_ESC.test(code), false,
            `esbuild emitted an escaped bare astral object key, which is still a bare astral identifier to a no-ICU V8: ${code}`);
    });
}

// Guards the guard: without charset 'ascii' the same input must produce raw
// non-ASCII. If this ever stops failing, the assertions above have gone blind.
test('sdkConfig: dropping charset ascii reintroduces raw non-ASCII output (#80)', async () => {
    const options = await esbuildOptionsFor('');
    const { code } = await esbuild.transform(SYMBOL_TABLE_SRC, { ...options, charset: 'utf8' });

    assert.equal(/[^\x00-\x7F]/.test(code), true,
        `Expected raw non-ASCII output with charset utf8, got: ${code}`);
});
