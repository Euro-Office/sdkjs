/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// Every QUnit page must be listed in tests/runAll.js (allTests, optionally with
// a skippedTests reason) or in notSuites, so new pages cannot be forgotten.

const test   = require('node:test');
const assert = require('node:assert/strict');
const fs     = require('node:fs');
const path   = require('node:path');

const TESTS_DIR = path.resolve(__dirname, '..', '..', 'tests');
const { allTests, skippedTests, notSuites } = require(path.join(TESTS_DIR, 'runAll.js'));

const pages = fs.readdirSync(TESTS_DIR, { recursive: true })
    .filter((f) => f.endsWith('.html') && !f.split(path.sep).includes('node_modules'))
    .map((f) => f.split(path.sep).join('/'));

test('every test page is in allTests or notSuites', () => {
    const listed = new Set([...allTests, ...Object.keys(notSuites)]);
    assert.deepEqual(pages.filter((p) => !listed.has(p)), []);
});

test('listed pages exist, skipped ones are in allTests', () => {
    const existing = new Set(pages);
    assert.deepEqual([...allTests, ...Object.keys(notSuites)].filter((p) => !existing.has(p)), []);
    assert.deepEqual(Object.keys(skippedTests).filter((p) => !allTests.includes(p)), []);
    assert.deepEqual(allTests.filter((p) => p in notSuites), []);
});
