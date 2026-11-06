/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const sources = [
	'common/stringserialize.js',
	'common/hash/hash/engine.js',
	'common/hash/hash/engine_ie.js'
];

function loadSerialization(source) {
	let code = fs.readFileSync(path.join(__dirname, '../..', source), 'utf8');
	if (source !== sources[0]) {
		// The workers embed serialization before their separate hash runtime.
		const end = code.indexOf('})(self);');
		assert.notEqual(end, -1, 'worker serialization prelude exists');
		code = code.slice(0, end + '})(self);'.length);
	}
	const context = vm.createContext({});
	vm.runInContext('window = self = globalThis;', context);
	vm.runInContext(code, context, { filename: source });
	return context;
}

for (const source of sources) {
	test(`${source}: hexadecimal encoding preserves all byte values and case`, () => {
		const context = loadSerialization(source);
		const encode = context.AscCommon.Hex.encode;
		const bytes = Uint8Array.from({ length: 256 }, (_, value) => value);

		for (const input of [bytes, Array.from(bytes), bytes.subarray(15, 33)]) {
			const lower = Buffer.from(input).toString('hex');
			assert.equal(encode(input), lower);
			assert.equal(encode(input, false), lower);
			assert.equal(encode(input, true), lower.toUpperCase());
		}
		assert.equal(encode([0, 1, 15, 16, 255]), '00010f10ff');
		assert.equal(encode(new Uint8Array()), '');
		assert.equal(encode([], true), '');
		assert.equal(context.AscCommon.Hex['encode'], encode);
	});

	test(`${source}: repeated encoding reuses byte representations`, () => {
		const context = loadSerialization(source);
		vm.runInContext(`
			globalThis.radixConversions = 0;
			const originalToString = Number.prototype.toString;
			Number.prototype.toString = function(radix) {
				if (radix === 16) ++globalThis.radixConversions;
				return originalToString.call(this, radix);
			};
		`, context);

		for (let index = 0; index < 100; index++) {
			assert.equal(context.AscCommon.Hex.encode([0, 15, 16, 255]), '000f10ff');
			assert.equal(context.AscCommon.Hex.encode([0, 15, 16, 255], true), '000F10FF');
		}
		assert.equal(context.radixConversions, 0,
			'encoding must not rebuild numeric byte-to-hex representations');
	});
}
