/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

// Run through `npm run lint` in build/: build/scripts/lint.cjs applies
// eslint-suppressions.json (existing errors) and eslint-metrics-baseline.json
// (complexity/size ratchet). The metric rules are warnings so the plain
// `eslint` CLI stays usable; the script turns regressions into failures.
export default [
	{
		ignores: [
			'**/*_ie.js',
			'**/engine/**',
			'common/hash/hash/engine.js',
			'vendor/',
			'common/externs/',
			'pdf/build/',
			'deploy/',
			'develop/',
			'build/',
			'tests/',
			'tools/',
		],
	},
	{
		files: ['**/*.js'],
		languageOptions: {
			ecmaVersion: 'latest',
			sourceType: 'script',
		},
		linterOptions: {
			reportUnusedDisableDirectives: 'off',
		},
		rules: {
			'for-direction': 'error',
			'no-debugger': 'error',
			'no-dupe-args': 'error',
			'no-dupe-else-if': 'error',
			'no-dupe-keys': 'error',
			'no-duplicate-case': 'error',
			'no-ex-assign': 'error',
			'no-func-assign': 'error',
			'no-invalid-regexp': 'error',
			'no-unreachable': 'error',
			'no-unsafe-negation': 'error',
			'complexity': ['warn', 20],
			'max-lines-per-function': ['warn', 300],
			'max-lines': ['warn', 10000],
		},
	},
	// common/Native stubs window/document/console for the embedded JS context of
	// DoctRenderer/x2t, which no test executes, so an undeclared identifier there
	// only fails in production. The hosts inject window and native
	// (doctrenderer.cpp) and CreateEmbedObject (v8_base.cpp); native.js defines
	// AscCommon before native_graphics.js is concatenated after it.
	{
		files: ['common/Native/native.js'],
		languageOptions: {globals: {window: 'readonly', native: 'readonly'}},
		rules: {'no-undef': 'error'},
	},
	{
		files: ['common/Native/native_graphics.js'],
		languageOptions: {globals: {window: 'readonly', CreateEmbedObject: 'readonly', AscCommon: 'readonly'}},
		rules: {'no-undef': 'error'},
	},
];
