/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// The lint gate passes on the baseline state and fails on a new error, a
// function that got more complex and a new over-threshold function.

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { run } = require('../scripts/lint.cjs');

const REPO_ROOT = path.resolve(__dirname, '..', '..');

const complexFn = (name, branches) => `${name} = function (a) {\n` +
	Array.from({ length: branches }, (_, i) => `\tif (a === ${i}) { return ${i}; }\n`).join('') +
	'\treturn -1;\n};\n';

function fixture() {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'sdkjs-lint-'));
	fs.copyFileSync(path.join(REPO_ROOT, 'eslint.config.mjs'), path.join(dir, 'eslint.config.mjs'));
	fs.mkdirSync(path.join(dir, 'word'));
	return dir;
}

const BASE = 'function a() {\n\treturn 1;\n\tvar dead = 2;\n}\n' + complexFn('Foo.prototype.bar', 25);

async function gate(dir, source) {
	fs.writeFileSync(path.join(dir, 'word', 'a.js'), source);
	const out = [];
	const code = await run({ cwd: dir, patterns: ['word'], log: (line) => out.push(line) });
	return { code, out: out.join('\n') };
}

test('lint gate ratchets errors and metrics against the baseline', async (t) => {
	const dir = fixture();
	t.after(() => fs.rmSync(dir, { recursive: true, force: true }));

	fs.writeFileSync(path.join(dir, 'word', 'a.js'), BASE);
	assert.equal(await run({ cwd: dir, patterns: ['word'], update: true, log: () => {} }), 0);
	const baseline = JSON.parse(fs.readFileSync(path.join(dir, 'eslint-metrics-baseline.json'), 'utf8'));
	assert.deepEqual(baseline.complexity, { 'word/a.js | Foo.prototype.bar = function (a) {': 26 });

	assert.equal((await gate(dir, BASE)).code, 0, 'baseline state passes');
	assert.equal((await gate(dir, '// moved down\n\n' + BASE)).code, 0, 'line shifts pass');

	const unreachable = await gate(dir, BASE + 'function b() {\n\treturn;\n\tb();\n}\n');
	assert.equal(unreachable.code, 1);
	assert.match(unreachable.out, /no-unreachable/);

	const worse = await gate(dir, BASE.replace(complexFn('Foo.prototype.bar', 25), complexFn('Foo.prototype.bar', 27)));
	assert.equal(worse.code, 1);
	assert.match(worse.out, /Got worse, baseline 26/);

	const added = await gate(dir, BASE + complexFn('Foo.prototype.baz', 21));
	assert.equal(added.code, 1);
	assert.match(added.out, /complexity of 22\. Maximum allowed is 20\. New/);

	const better = await gate(dir, BASE.replace('\tvar dead = 2;\n', '').replace(complexFn('Foo.prototype.bar', 25), complexFn('Foo.prototype.bar', 22)));
	assert.equal(better.code, 0, 'improvements pass');
	assert.match(better.out, /2 baseline entries improved/);
});

test('line counts get 10% slack, complexity none', async () => {
	const { compare } = require('../scripts/lint.cjs');
	const entry = (rule, value) => ({ errors: {}, metrics: { [rule]: { k: { value, msg: { ruleId: rule, message: rule } } } } });
	const base = { 'max-lines': { k: 10000 }, 'max-lines-per-function': { k: 400 }, 'complexity': { k: 30 } };
	assert.equal(compare(entry('max-lines', 11000), {}, base).failures.length, 0);
	assert.equal(compare(entry('max-lines', 11001), {}, base).failures.length, 1);
	assert.equal(compare(entry('max-lines-per-function', 440), {}, base).failures.length, 0);
	assert.equal(compare(entry('max-lines-per-function', 441), {}, base).failures.length, 1);
	assert.equal(compare(entry('complexity', 31), {}, base).failures.length, 1);
});
