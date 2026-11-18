/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

/**
 * Keeps the public API documented without mass-editing upstream JSDoc.
 *
 * Runs stock `jsdoc -X` over the files the API reference is generated from and
 * flags public members (PascalCase or pluginMethod_ methods on Api* / Plugin
 * classes, and plugin events) that lack a description, have a param without a
 * type, or lack @typeofeditors. Known violations live in api-docs-baseline.json;
 * only new ones fail. jsdoc parse errors always fail.
 *
 *   node build/scripts/check-api-docs.cjs            check against the baseline
 *   node build/scripts/check-api-docs.cjs --update   rewrite the baseline
 */

'use strict';

const fs           = require('node:fs');
const path         = require('node:path');
const { spawnSync } = require('node:child_process');

const REPO_ROOT = path.resolve(__dirname, '..', '..');
const BASELINE  = path.join(REPO_ROOT, 'build', 'api-docs-baseline.json');

const API_FILES = [
	'word/apiBuilder.js',
	'cell/apiBuilder.js',
	'slide/apiBuilder.js',
	'pdf/apiBuilder.js',
	'common/apiBase_plugins.js',
	'common/plugins/plugin_base_api.js',
	'word/api_plugins.js',
	'cell/api_plugins.js',
	'slide/api_plugins.js',
	'pdf/api_plugins.js',
	'common/base-plugin-events.js',
	'word/plugin-events.js',
	'cell/plugin-events.js',
	'slide/plugin-events.js',
];

const PUBLIC_CLASS  = /^(Api\w*|Plugin)$/;
const PUBLIC_METHOD = /^([A-Z]|pluginMethod_)/;

function runJsdoc(files, root) {
	const res = spawnSync(process.execPath, [require.resolve('jsdoc/jsdoc.js'), '-X', ...files], {
		cwd: root,
		encoding: 'utf8',
		maxBuffer: 1 << 30,
	});
	const errors = (res.stderr || '').split('\n').filter((l) => l.trim());
	if (res.status !== 0 && errors.length === 0)
		errors.push(`jsdoc exited with status ${res.status}`);
	return { doclets: res.stdout ? JSON.parse(res.stdout) : [], errors };
}

// Returns "Class#Name" / "Class.Name" for public API doclets, null otherwise.
function publicName(d) {
	if (d.access === 'private' || d.scope === 'inner' || !d.memberof)
		return null;
	const cls = d.memberof.split('~').pop();
	if (d.kind === 'event')
		return d.undocumented ? null : d.longname;
	if (d.kind !== 'function' || !PUBLIC_CLASS.test(cls) || !PUBLIC_METHOD.test(d.name))
		return null;
	return cls + (d.scope === 'static' ? '.' : '#') + d.name;
}

function violationsOf(d) {
	if (d.undocumented || !d.description)
		return ['missing-description'];
	const out = [];
	if ((d.params || []).some((p) => !p.type))
		out.push('missing-param-type');
	if (!(d.tags || []).some((t) => t.title === 'typeofeditors'))
		out.push('missing-typeofeditors');
	return out;
}

function collect(files, root) {
	const { doclets, errors } = runJsdoc(files, root);
	const realRoot = fs.realpathSync(root);
	const violations = new Set();
	const documented = {};
	for (const d of doclets) {
		const name = publicName(d);
		if (!name)
			continue;
		const file = path.relative(realRoot, path.join(d.meta.path, d.meta.filename)).split(path.sep).join('/');
		const v = violationsOf(d);
		if (v.length === 0)
			documented[file] = (documented[file] || 0) + 1;
		for (const rule of v)
			violations.add(`${file}:${name}:${rule}`);
	}
	return { violations: [...violations].sort(), errors, documented };
}

function compare(current, baseline) {
	const base = new Set(baseline);
	const cur  = new Set(current);
	return {
		added: current.filter((k) => !base.has(k)),
		fixed: baseline.filter((k) => !cur.has(k)),
	};
}

function main(argv) {
	const start = Date.now();
	const { violations, errors, documented } = collect(API_FILES, REPO_ROOT);

	if (errors.length) {
		console.error('jsdoc reported errors:\n' + errors.join('\n'));
		return 1;
	}
	console.log('Documented public members per file:');
	for (const f of Object.keys(documented).sort())
		console.log(`  ${f}: ${documented[f]}`);

	if (argv.includes('--update')) {
		fs.writeFileSync(BASELINE, JSON.stringify(violations, null, '\t') + '\n');
		console.log(`Baseline written: ${violations.length} known violations.`);
		return 0;
	}

	const baseline = fs.existsSync(BASELINE) ? JSON.parse(fs.readFileSync(BASELINE, 'utf8')) : [];
	const { added, fixed } = compare(violations, baseline);
	console.log(`Baseline: ${baseline.length} known violations (${((Date.now() - start) / 1000).toFixed(1)}s).`);
	if (fixed.length) {
		console.log(`\n${fixed.length} fixed since baseline, trim it with --update:`);
		fixed.forEach((k) => console.log('  ' + k));
	}
	if (added.length) {
		console.error(`\n${added.length} new API doc violations:`);
		added.forEach((k) => console.error('  ' + k));
		console.error('\nDocument them (description, typed @param, @typeofeditors), or run with --update if intended.');
		return 1;
	}
	return 0;
}

module.exports = { collect, compare };

if (require.main === module)
	process.exitCode = main(process.argv.slice(2));
