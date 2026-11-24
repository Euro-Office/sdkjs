/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// QUnit pages must load their libraries from tests/node_modules (pinned in
// tests/package.json) or the repo, never from a CDN: tests stay reproducible
// and run offline.

const test   = require('node:test');
const assert = require('node:assert/strict');
const fs     = require('node:fs');
const path   = require('node:path');

const TESTS_DIR = path.resolve(__dirname, '..', '..', 'tests');
const EXTERNAL  = /<(?:script|link)\b[^>]*\b(?:src|href)\s*=\s*["']?(?:https?:)?\/\//i;

test('test pages load no external scripts or stylesheets', () => {
    const pages = fs.readdirSync(TESTS_DIR, { recursive: true })
        .filter((f) => f.endsWith('.html') && !f.split(path.sep).includes('node_modules'));
    assert.ok(pages.length > 0, 'no test pages found');
    for (const page of pages) {
        const html = fs.readFileSync(path.join(TESTS_DIR, page), 'utf8');
        assert.ok(!EXTERNAL.test(html), `tests/${page} loads an external URL`);
    }
});
