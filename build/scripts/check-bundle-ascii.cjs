/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

/**
 * Guards the two invariants the built sdkjs bundles must satisfy for
 * doctrenderer/x2t, which run them on a V8 built with
 * v8_enable_i18n_support=false (core: Common/3dParty/v8/tools/8.9/*).
 *
 *   1. No raw non-ASCII byte.
 *   2. No object key that is a bare supplementary-plane ("astral") identifier.
 *
 * (2) is the subtle one and the reason a plain "is it ASCII" grep is not
 * enough: with format.ascii_only but without format.quote_keys, Terser emits
 *
 *     \u{1d552}:"\\doublea"
 *
 * which contains no byte > 0x7F yet still parses as a bare identifier whose
 * code point is U+1D552. A no-ICU V8 rejects that exactly as it rejects the
 * raw UTF-8 spelling, so invariant (1) alone would let issue #80 regress
 * silently. See build/webpack.sdk.factory.mjs.
 *
 * Usage: node build/scripts/check-bundle-ascii.cjs [file...]
 * With no arguments, scans the built sdk-all bundles under the deploy root.
 */

'use strict';

const fs   = require('fs');
const path = require('path');
const { sync: globSync } = require('glob');
const { resolveBuildRoot } = require('../lib/env.cjs');

// An unquoted object key in `{`/`,` position that starts with either a raw
// astral character or a unicode escape. A quoted key cannot match: the quote
// sits between the delimiter and the escape.
const BARE_ASTRAL_RAW = /[{,]\s*[\u{10000}-\u{10FFFF}]/u;
const BARE_ASTRAL_ESC = /[{,]\s*\\u(?:\{[0-9a-fA-F]{5,6}\}|[dD][89abAB][0-9a-fA-F]{2})/;

function sampleAround(text, regex, limit = 3) {
    const global = new RegExp(regex.source, regex.flags.includes('g') ? regex.flags : regex.flags + 'g');
    const out = [];
    let m;
    while ((m = global.exec(text)) !== null && out.length < limit) {
        out.push(text.slice(m.index, Math.min(m.index + 40, text.length)));
    }
    return out;
}

function checkFile(file) {
    const buf  = fs.readFileSync(file);
    const problems = [];

    let nonAscii = 0;
    for (const byte of buf) if (byte > 0x7f) nonAscii++;
    if (nonAscii > 0) {
        problems.push(`${nonAscii} raw non-ASCII byte(s); Terser format.ascii_only must stay enabled`);
    }

    const text = buf.toString('utf8');
    for (const [label, re] of [['raw', BARE_ASTRAL_RAW], ['escaped', BARE_ASTRAL_ESC]]) {
        if (re.test(text)) {
            const samples = sampleAround(text, re).map((s) => JSON.stringify(s)).join(', ');
            problems.push(`bare astral object key (${label} form); Terser format.quote_keys must stay enabled -- e.g. ${samples}`);
        }
    }

    return problems;
}

function main(argv) {
    let files = argv.slice(2);
    if (files.length === 0) {
        const buildDir  = path.resolve(__dirname, '..');
        const deployRoot = resolveBuildRoot(buildDir);
        files = globSync('*/sdk-all{,-min}.js', { cwd: deployRoot, absolute: true }).sort();
        if (files.length === 0) {
            console.error(`check-bundle-ascii: no sdk-all bundles found under ${deployRoot} -- nothing was verified`);
            return 1;
        }
    }

    let failed = 0;
    for (const file of files) {
        const problems = checkFile(file);
        if (problems.length === 0) {
            console.log(`ok: ${file}`);
            continue;
        }
        failed++;
        for (const problem of problems) {
            console.error(`::error file=${file}::${problem} (see issue #80)`);
        }
    }

    console.log(`check-bundle-ascii: ${files.length} bundle(s) scanned, ${failed} failed`);
    return failed === 0 ? 0 : 1;
}

if (require.main === module) process.exit(main(process.argv));

module.exports = { checkFile, BARE_ASTRAL_RAW, BARE_ASTRAL_ESC };
