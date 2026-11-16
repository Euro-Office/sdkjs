/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// Stray files, Closure Compiler leftovers and unused sample documents were dropped.
// Upstream syncs re-add them: resolve the modify/delete conflict with `git rm`.

const test   = require('node:test');
const assert = require('node:assert/strict');
const fs     = require('node:fs');
const path   = require('node:path');

const SRC_ROOT = path.resolve(__dirname, '..', '..');
const REMOVED  = [
    'um',
    'jshintignore',
    '.codeclimate.yml',
    'sonar-project.properties',
    'word/Documents',
    'cell/documentation',
    'common/externs',
    'configs/externs.json',
    'pdf/build/make.py',
    'common/plugins/make.py',
    'common/embed/make.py',
    'common/embed/min.py',
    'build/build-develop-addons.py',
    'build/build-develop-addons-sdk-advanced.py',
    'build/build-develop-addons-sdk-whitespace.py',
    'build/build-mobile.command',
    'tools/minify_path_boolean.py',
    'slide/document',
    'word/document/editor.js',
];

test('removed files stay deleted', () => {
    for (const file of REMOVED) {
        assert.equal(fs.existsSync(path.join(SRC_ROOT, file)), false, `${file} is back`);
    }
});
