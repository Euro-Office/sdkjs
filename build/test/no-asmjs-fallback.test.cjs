/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// The engines are WebAssembly only: core-wasm builds no asm.js (*_ie.js)
// variants, so committed ones would be stale. Upstream syncs bring them back.

const test   = require('node:test');
const assert = require('node:assert/strict');
const path   = require('path');
const { execFileSync } = require('child_process');

const SRC_ROOT = path.resolve(__dirname, '..', '..');

function git(...args) {
    try {
        return execFileSync('git', ['-C', SRC_ROOT, ...args], { encoding: 'utf8' }).trim();
    } catch (err) {
        if (err.status === 1) return ''; // git grep: no match
        throw err;
    }
}

test('no asm.js engine files are tracked', () => {
    assert.equal(git('ls-files', '*_ie.js', '*.js.mem'), '');
});

// pdf/src/engine/viewer.js is legacy pdf/build/make.py output, only loaded by
// pdf/test/index_deploy.html
test('no source loads an asm.js engine', () => {
    assert.equal(git('grep', '-n', '-E', '_ie(\\.js)?["\']', '--', '*.js', '*.cjs',
        ':!build/test/', ':!pdf/src/engine/viewer.js'), '');
});
