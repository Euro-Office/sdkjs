/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// Polyfills for built-ins every supported engine has natively (ES2022 browsers,
// doctrenderer V8) and jQuery 3.2.1 were dropped. Upstream syncs re-add them:
// resolve the modify/delete conflict with `git rm`.

const test   = require('node:test');
const assert = require('node:assert/strict');
const fs     = require('node:fs');
const path   = require('node:path');

const SRC_ROOT = path.resolve(__dirname, '..', '..');
const REMOVED  = ['vendor/array.js', 'vendor/polyfill.js', 'vendor/jquery.min.js'];

test('removed vendor files stay deleted', () => {
    for (const file of REMOVED) {
        assert.equal(fs.existsSync(path.join(SRC_ROOT, file)), false, `${file} is back`);
    }
});

test('build configs and tooling do not reference removed vendor files', () => {
    const dirs = ['configs', 'build/loaders', 'build/scripts'];
    for (const dir of dirs) {
        for (const name of fs.readdirSync(path.join(SRC_ROOT, dir))) {
            const content = fs.readFileSync(path.join(SRC_ROOT, dir, name), 'utf8');
            for (const file of REMOVED) {
                const base = path.basename(file);
                assert.ok(!content.includes(base), `${dir}/${name} references ${base}`);
            }
        }
    }
});
