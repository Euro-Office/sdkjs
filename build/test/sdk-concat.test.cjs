/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

const test   = require('node:test');
const assert = require('node:assert/strict');
const fs     = require('node:fs');
const os     = require('node:os');
const path   = require('node:path');

const { validate } = require('schema-utils');

const sdkConcatLoader = require('../loaders/sdk-concat.cjs');

// Minimal on-disk fixture satisfying loadAllConfigs()'s expectations (a real
// configs/word.json + two small source files), so the loader can run
// end-to-end without a real sdkjs checkout.
function makeFixtureSrcRoot() {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'sdk-concat-test-'));
    fs.mkdirSync(path.join(root, 'configs'));
    fs.mkdirSync(path.join(root, 'src'));

    fs.writeFileSync(path.join(root, 'src', 'a.js'), 'let a = (x) => x + 1;\n');
    fs.writeFileSync(path.join(root, 'src', 'b.js'), 'const b = 2;\n');
    fs.writeFileSync(
        path.join(root, 'configs', 'word.json'),
        JSON.stringify({ sdk: { min: ['src/a.js', 'src/b.js'] } })
    );

    return root;
}

// Runs the loader with a minimal mocked webpack loader context, mirroring
// what webpack itself provides (this.async/getOptions/context/addDependency/
// sourceMap), and resolves with { content, map }.
function runLoader(srcRoot, opts) {
    return new Promise((resolve, reject) => {
        const context = {
            async: () => (err, content, map) => (err ? reject(err) : resolve({ content, map })),
            getOptions: () => opts,
            context: srcRoot,
            sourceMap: true,
            addDependency: () => {},
            addContextDependency: () => {},
        };
        sdkConcatLoader.call(context);
    });
}

test('loader: concatenates files verbatim, with no per-file transform', async () => {
    const srcRoot = makeFixtureSrcRoot();
    try {
        const { content } = await runLoader(srcRoot, { module: 'word', chunk: 'min', srcRoot });
        assert.match(content, /let a = \(x\) => x \+ 1;/);
        assert.match(content, /const b = 2;/);
    } finally {
        fs.rmSync(srcRoot, { recursive: true, force: true });
    }
});

test('loader: source map sourcesContent holds the original file content', async () => {
    const srcRoot = makeFixtureSrcRoot();
    try {
        const { map } = await runLoader(srcRoot, { module: 'word', chunk: 'min', srcRoot });

        assert.ok(map, 'expected a source map (this.sourceMap was true)');
        const aIndex = map.sources.findIndex(s => s.endsWith(path.join('src', 'a.js')));
        assert.notEqual(aIndex, -1);

        const originalA = fs.readFileSync(path.join(srcRoot, 'src', 'a.js'), 'utf8');
        assert.equal(map.sourcesContent[aIndex], originalA);
    } finally {
        fs.rmSync(srcRoot, { recursive: true, force: true });
    }
});

// The loader itself never calls schema-utils directly — this.getOptions(schema)
// does, inside real webpack. These tests exercise the exported schema the same
// way webpack does, so a bad option value (e.g. a typo'd SDK_PLATFORM) is
// actually asserted to fail, not just assumed to because the loader *passes*
// a schema.
test('schema: accepts a valid options object', () => {
    assert.doesNotThrow(() => validate(sdkConcatLoader.schema, {
        module: 'word',
        chunk: 'min',
        platform: 'desktop',
    }, { name: 'sdk-concat-loader' }));
});

test('schema: rejects an invalid platform value (e.g. a mistyped SDK_PLATFORM)', () => {
    assert.throws(() => validate(sdkConcatLoader.schema, {
        module: 'word',
        chunk: 'min',
        platform: 'Desktop',
    }, { name: 'sdk-concat-loader' }));
});

test('schema: rejects an invalid module value', () => {
    assert.throws(() => validate(sdkConcatLoader.schema, {
        module: 'bogus',
        chunk: 'min',
    }, { name: 'sdk-concat-loader' }));
});

test('schema: rejects unknown properties', () => {
    assert.throws(() => validate(sdkConcatLoader.schema, {
        module: 'word',
        chunk: 'min',
        typoedOption: true,
    }, { name: 'sdk-concat-loader' }));
});
