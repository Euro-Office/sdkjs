/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// Every supported engine (ES2022 browsers, doctrenderer V8/JavaScriptCore) has
// these built-ins, so their polyfills were removed. Upstream syncs re-add them.
// The unguarded Math.sinh/cosh/asinh/acosh/atanh/log10 overrides and the custom
// Math/String helpers are kept on purpose: formulas depend on them.

const test   = require('node:test');
const assert = require('node:assert/strict');
const fs     = require('node:fs');
const path   = require('node:path');
const vm     = require('node:vm');

const SRC_ROOT = path.resolve(__dirname, '..', '..');
const SKIP_DIR = new Set(['vendor', 'engine', 'externs', 'node_modules', 'build', 'tests', 'deploy', 'develop', '.git']);

const BUILTINS = [
	'Number.isInteger', 'Number.isFinite',
	'Math.tanh', 'Math.sign', 'Math.trunc', 'Math.log1p', 'Math.expm1',
	'Object.values', 'Object.create', 'String.fromCodePoint',
	...['codePointAt', 'replaceAll', 'startsWith', 'endsWith', 'trim', 'repeat', 'padStart', 'includes']
		.map((name) => 'String.prototype.' + name),
];

// `X.y =` or `X['y'] =`, but not `==`
const assignment = (name) => {
	const dot     = name.lastIndexOf('.');
	const owner   = name.slice(0, dot).replace(/\./g, '\\.');
	const member  = name.slice(dot + 1);
	return new RegExp(`\\b${owner}(\\.${member}|\\[['"]${member}['"]\\])\\s*=(?!=)`);
};
const PATTERNS = BUILTINS.map((name) => [name, assignment(name)]);

function sources(dir, out = []) {
	for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
		if (entry.isDirectory()) {
			if (!SKIP_DIR.has(entry.name)) sources(path.join(dir, entry.name), out);
		} else if (entry.name.endsWith('.js') && !entry.name.endsWith('_ie.js')) {
			out.push(path.join(dir, entry.name));
		}
	}
	return out;
}

test('sources do not polyfill ES2022 built-ins', () => {
	const files = [...sources(SRC_ROOT), path.join(SRC_ROOT, 'vendor/string.js')];
	const hits  = [];
	for (const file of files) {
		fs.readFileSync(file, 'utf8').split('\n').forEach((line, i) => {
			for (const [name, re] of PATTERNS) {
				if (re.test(line)) hits.push(`${path.relative(SRC_ROOT, file)}:${i + 1} ${name}`);
			}
		});
	}
	assert.deepEqual(hits, []);
});

test('vendor/string.js keeps its custom String methods', () => {
	const ctx = vm.createContext({});
	vm.runInContext(fs.readFileSync(path.join(SRC_ROOT, 'vendor/string.js'), 'utf8'), ctx);
	assert.equal(vm.runInContext('"a\\u{1F600}b".codePointsArray().join(",")', ctx), '97,128512,98');
	assert.equal(vm.runInContext('" hello".sentenceCase()', ctx), 'Hello');
	assert.equal(vm.runInContext('"abc".strongMatch(/[a-c]+/)', ctx), true);
});
