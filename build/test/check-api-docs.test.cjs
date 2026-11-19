/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

const test   = require('node:test');
const assert = require('node:assert/strict');
const fs     = require('node:fs');
const os     = require('node:os');
const path   = require('node:path');

const { collect, compare } = require('../scripts/check-api-docs.cjs');

const FIXTURE = `
(function(window) {
	/**
	 * @class
	 */
	function ApiFoo() {}

	/**
	 * Documented method.
	 * @memberof ApiFoo
	 * @typeofeditors ["CDE"]
	 * @param {string} text - The text.
	 */
	ApiFoo.prototype.Documented = function(text) {};

	/**
	 * @memberof ApiFoo
	 * @param text - The text.
	 */
	ApiFoo.prototype.NoDescription = function(text) {};

	/**
	 * Has an untyped param.
	 * @memberof ApiFoo
	 * @typeofeditors ["CDE"]
	 * @param text - The text.
	 */
	ApiFoo.prototype.UntypedParam = function(text) {};

	ApiFoo.prototype.Undocumented = function() {};
	ApiFoo.prototype.private_Helper = function() {};
	ApiFoo.prototype.internalHelper = function() {};

	/**
	 * Hidden.
	 * @memberof ApiFoo
	 * @access private
	 */
	ApiFoo.prototype.Hidden = function() {};
})(window);
`;

function fixtureRoot(source) {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'check-api-docs-'));
	fs.writeFileSync(path.join(root, 'api.js'), source);
	return root;
}

test('flags undocumented public methods, not private ones', () => {
	const { violations, errors, documented } = collect(['api.js'], fixtureRoot(FIXTURE));
	assert.deepEqual(errors, []);
	assert.deepEqual(violations, [
		'api.js:ApiFoo#NoDescription:missing-description',
		'api.js:ApiFoo#Undocumented:missing-description',
		'api.js:ApiFoo#UntypedParam:missing-param-type',
	]);
	assert.deepEqual(documented, { 'api.js': 1 });
});

test('reports a missing @typeofeditors', () => {
	const src = FIXTURE.replace('@typeofeditors ["CDE"]\n\t * @param {string}', '@param {string}');
	const { violations } = collect(['api.js'], fixtureRoot(src));
	assert.ok(violations.includes('api.js:ApiFoo#Documented:missing-typeofeditors'));
});

test('reports jsdoc parse errors', () => {
	const { errors } = collect(['api.js'], fixtureRoot('function broken( {\n'));
	assert.match(errors.join('\n'), /Unable to parse/);
});

test('baseline suppresses known violations and catches new ones', () => {
	const { violations } = collect(['api.js'], fixtureRoot(FIXTURE));
	assert.deepEqual(compare(violations, violations), { added: [], fixed: [] });

	const known = violations.filter((k) => !k.includes('Undocumented'));
	const stale = 'api.js:ApiFoo#Gone:missing-description';
	assert.deepEqual(compare(violations, [...known, stale]), {
		added: ['api.js:ApiFoo#Undocumented:missing-description'],
		fixed: [stale],
	});
});
