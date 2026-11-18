/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// ESLint with two baselines, so CI fails only on regressions:
// - errors: ESLint's own bulk suppressions file (eslint-suppressions.json,
//   per file and rule counts), which plain `eslint` applies too.
// - complexity, max-lines-per-function, max-lines (warnings in the config):
//   eslint-metrics-baseline.json holds the value per function and file, so a
//   function that gets worse fails, not only a new one. Bulk suppressions
//   only count violations and cannot express that. Line counts get 10%
//   slack over their baseline so routine edits in the 10k-line files pass;
//   complexity has none.
// `node scripts/lint.cjs --update` rewrites both baselines.

const fs = require('node:fs');
const path = require('node:path');
const { ESLint } = require('eslint');

const REPO_ROOT = path.resolve(__dirname, '..', '..');
const DIRS = ['common', 'word', 'cell', 'slide', 'pdf', 'visio'];
const SUPPRESSIONS = 'eslint-suppressions.json';
const METRICS = 'eslint-metrics-baseline.json';
const METRIC_VALUE = {
	'complexity': /complexity of (\d+)/,
	'max-lines-per-function': /too many lines \((\d+)\)/,
	'max-lines': /too many lines \((\d+)\)/,
};

const SLACK = { 'max-lines-per-function': 1.1, 'max-lines': 1.1 };

const posix = (cwd, file) => path.relative(cwd, file).split(path.sep).join('/');

// Functions are keyed by the trimmed source line they start on (usually
// `Foo.prototype.bar = function (...)`), plus a counter for repeats, so the
// key survives code being added above them.
function metricKey(rel, msg, lines, seen) {
	if (msg.ruleId === 'max-lines') {
		return `${rel} | file`;
	}
	const base = `${rel} | ${(lines[msg.line - 1] || '').trim()}`;
	const n = seen[`${msg.ruleId} ${base}`] = (seen[`${msg.ruleId} ${base}`] || 0) + 1;
	return n > 1 ? `${base} #${n}` : base;
}

function collect(results, cwd) {
	const errors = {};
	const metrics = {};
	for (const result of results) {
		const rel = posix(cwd, result.filePath);
		const lines = result.source !== undefined ? result.source.split('\n') : fs.readFileSync(result.filePath, 'utf8').split('\n');
		const seen = {};
		for (const msg of result.messages) {
			if (msg.ruleId in METRIC_VALUE) {
				const value = Number(METRIC_VALUE[msg.ruleId].exec(msg.message)[1]);
				const rule = (metrics[msg.ruleId] ??= {});
				rule[metricKey(rel, msg, lines, seen)] = { value, msg: { ...msg, file: rel } };
			} else {
				const file = (errors[rel] ??= {});
				(file[msg.ruleId || 'fatal'] ??= []).push({ ...msg, file: rel });
			}
		}
	}
	return { errors, metrics };
}

function toBaselines({ errors, metrics }) {
	const suppressions = {};
	for (const [file, rules] of Object.entries(errors).sort()) {
		for (const [rule, msgs] of Object.entries(rules).sort()) {
			if (rule !== 'fatal') {
				(suppressions[file] ??= {})[rule] = { count: msgs.length };
			}
		}
	}
	const baseline = {};
	for (const [rule, entries] of Object.entries(metrics).sort()) {
		baseline[rule] = Object.fromEntries(Object.entries(entries).sort().map(([key, e]) => [key, e.value]));
	}
	return { suppressions, baseline };
}

// Returns the messages that fail the gate and how many baseline entries could shrink.
function compare(collected, suppressions, baseline) {
	const failures = [];
	let improved = 0;
	for (const [file, rules] of Object.entries(collected.errors)) {
		for (const [rule, msgs] of Object.entries(rules)) {
			const allowed = suppressions[file]?.[rule]?.count ?? 0;
			if (msgs.length > allowed) {
				failures.push(...msgs.map((m) => allowed ? { ...m, message: `${m.message} (${msgs.length} in file, baseline ${allowed})` } : m));
			} else if (msgs.length < allowed) {
				improved++;
			}
		}
	}
	for (const [file, rules] of Object.entries(suppressions)) {
		improved += Object.keys(rules).filter((rule) => !collected.errors[file]?.[rule]).length;
	}
	for (const [rule, entries] of Object.entries(collected.metrics)) {
		for (const [key, { value, msg }] of Object.entries(entries)) {
			const old = baseline[rule]?.[key];
			if (old === undefined) {
				failures.push({ ...msg, message: `${msg.message} New (or its first line changed, then run \`npm run lint:update\`).` });
			} else if (value > Math.floor(old * (SLACK[rule] || 1))) {
				failures.push({ ...msg, message: `${msg.message} Got worse, baseline ${old}.` });
			} else if (value < old) {
				improved++;
			}
		}
	}
	for (const [rule, entries] of Object.entries(baseline)) {
		improved += Object.keys(entries).filter((key) => !collected.metrics[rule]?.[key]).length;
	}
	return { failures, improved };
}

async function run({ cwd = REPO_ROOT, patterns = DIRS, update = false, cache = false, log = console.log } = {}) {
	const eslint = new ESLint({ cwd, cache, cacheLocation: path.join(cwd, 'build', '.eslintcache'), cacheStrategy: 'content' });
	const collected = collect(await eslint.lintFiles(patterns), cwd);
	const read = (name) => fs.existsSync(path.join(cwd, name)) ? JSON.parse(fs.readFileSync(path.join(cwd, name), 'utf8')) : {};
	const write = (name, data) => fs.writeFileSync(path.join(cwd, name), JSON.stringify(data, null, '\t') + '\n');
	if (update) {
		const { suppressions, baseline } = toBaselines(collected);
		write(SUPPRESSIONS, suppressions);
		write(METRICS, baseline);
		log(`Wrote ${SUPPRESSIONS} and ${METRICS}.`);
		const fatal = Object.values(collected.errors).flatMap((rules) => rules.fatal || []);
		fatal.forEach((m) => log(`${m.file}:${m.line}:${m.column} ${m.message}`));
		return fatal.length ? 1 : 0;
	}
	const { failures, improved } = compare(collected, read(SUPPRESSIONS), read(METRICS));
	for (const m of failures) {
		log(`${m.file}:${m.line}:${m.column} ${m.message} [${m.ruleId || 'fatal'}]`);
	}
	if (improved) {
		log(`${improved} baseline entries improved or went away, shrink the baseline with \`npm run lint:update\`.`);
	}
	log(failures.length ? `${failures.length} new lint problems.` : 'No new lint problems.');
	return failures.length ? 1 : 0;
}

module.exports = { run, collect, compare, toBaselines };

if (require.main === module) {
	run({ update: process.argv.includes('--update'), cache: process.argv.includes('--cache') })
		.then((code) => { process.exitCode = code; })
		.catch((err) => { console.error(err); process.exitCode = 2; });
}
