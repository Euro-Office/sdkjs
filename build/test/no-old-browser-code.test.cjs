/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// The browser floor is ES2022 (CEF 107, iOS 17): IE, legacy Edge and Presto
// Opera cannot load the bundles. Their code paths were removed; upstream syncs
// re-add them, so this fails when one comes back.

const test   = require('node:test');
const assert = require('node:assert/strict');
const fs     = require('node:fs');
const path   = require('node:path');

const SRC_ROOT = path.resolve(__dirname, '..', '..');
const SKIP_DIR = new Set(['vendor', 'engine', 'externs', 'node_modules', 'build', 'tests', 'deploy', 'develop', '.git']);

// common/plugins.js embeds a minified plugin bundle as a string; that copy is
// generated elsewhere and still carries the old detection.
const SKIP_LINE = /\/\*<code>\*\//;

const PATTERNS = [
	['AscBrowser.isIE / isIeEdge / isIE9-11', /AscBrowser\.(isIE|isIeEdge|isIE9|isIE10|isIE11)\b/],
	['AscBrowser.isOperaOld', /\bisOperaOld\b/],
	['window.opera', /window\.opera\b/],
	['msie user agent sniffing', /["']msie|["']trident/],
	['ActiveXObject', /\bActiveXObject\b/],
	['msSaveBlob', /\bmsSaveBlob\b/],
	['BlobBuilder', /\bBlobBuilder\b/],
	['VBArray', /\bVBArray\b/],
	['createTextRange', /\bcreateTextRange\b/],
	['DOM attachEvent', /\b(window|document)\.attachEvent\b|\.attachEvent\(\s*["']on[a-z]+["']/],
	['DOM detachEvent', /\b(window|document)\.detachEvent\b|\.detachEvent\(\s*["']on[a-z]+["']/],
	['prefixed requestAnimationFrame', /\b(webkit|moz|ms|o)(Request|Cancel|CancelRequest)AnimationFrame\b/],
	['cancelRequestAnimationFrame', /\bcancelRequestAnimationFrame\b/],
	['prefixed imageSmoothingEnabled', /\b(webkit|moz|ms|o)ImageSmoothingEnabled\b/],
];

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

test('no IE, legacy Edge or Presto Opera code paths in sources', () => {
	const hits = [];
	for (const file of sources(SRC_ROOT)) {
		const lines = fs.readFileSync(file, 'utf8').split('\n');
		lines.forEach((line, i) => {
			if (SKIP_LINE.test(line)) return;
			for (const [name, re] of PATTERNS) {
				if (re.test(line)) hits.push(`${path.relative(SRC_ROOT, file)}:${i + 1} ${name}`);
			}
		});
	}
	assert.deepEqual(hits, []);
});
