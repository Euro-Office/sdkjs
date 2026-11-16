/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// common/Drawings/Format/path-boolean.js is lazy-loaded for "merge shapes"
// (AscCommon.loadPathBoolean) and deployed by build/scripts/deploy-assets.cjs.
// fixtures/path-boolean-golden.json holds the results of the Closure-compiled
// path-boolean-min.js it replaced, so the readable source must reproduce them.

const test   = require('node:test');
const assert = require('node:assert/strict');
const fs     = require('node:fs');
const path   = require('node:path');
const os     = require('node:os');
const vm     = require('node:vm');
const { deployJsFile } = require('../scripts/deploy-assets.cjs');

const SRC_ROOT = path.resolve(__dirname, '..', '..');
const SOURCE   = 'common/Drawings/Format/path-boolean.js';
const golden   = require('./fixtures/path-boolean-golden.json');

function loadCompoundPath(code) {
    const ctx = vm.createContext({});
    vm.runInContext('var window = this;', ctx);
    vm.runInContext(code, ctx);
    return ctx.AscCommon.PathBoolean.CompoundPath;
}

function rect(CP, x, y, w, h) {
    const p = new CP();
    p.moveTo(x, y);
    p.lineTo(x + w, y);
    p.lineTo(x + w, y + h);
    p.lineTo(x, y + h);
    p.closePath();
    return p;
}

function circle(CP, cx, cy, r) {
    const k = 0.5523 * r;
    const p = new CP();
    p.moveTo(cx + r, cy);
    p.cubicCurveTo(cx + r, cy + k, cx + k, cy + r, cx, cy + r);
    p.cubicCurveTo(cx - k, cy + r, cx - r, cy + k, cx - r, cy);
    p.cubicCurveTo(cx - r, cy - k, cx - k, cy - r, cx, cy - r);
    p.cubicCurveTo(cx + k, cy - r, cx + r, cy - k, cx + r, cy);
    p.closePath();
    return p;
}

function triangle(CP, x, y, s) {
    const p = new CP();
    p.moveTo(x, y);
    p.lineTo(x + s, y + s);
    p.lineTo(x - s, y + s);
    p.closePath();
    return p;
}

const round = (n) => Math.round(n * 1e6) / 1e6;

// Reads results only through the exported (bracket-named) API, the surface
// common/Drawings/CommonController.js uses.
function dump(result) {
    return (Array.isArray(result) ? result : [result]).map((item) => {
        const children = item['getChildren'] ? item['getChildren']() : null;
        const paths    = Array.isArray(children) && children.length ? children : [item];
        const bounds   = item['getBounds']();
        return {
            bounds: [bounds['getLeft'](), bounds['getTop'](), bounds['getWidth'](), bounds['getHeight']()].map(round),
            subs  : paths.map((p) => ({
                closed: p['isClosed'](),
                segs  : p['getSegments']().map((s) => [s['getPoint'](), s['getHandleIn'](), s['getHandleOut']()]
                    .flatMap((pt) => [round(pt['getX']()), round(pt['getY']())])),
            })),
        };
    });
}

const PAIRS = {
    'rect/rect'    : (CP) => [rect(CP, 0, 0, 10, 10), rect(CP, 5, 5, 10, 10)],
    'circle/rect'  : (CP) => [circle(CP, 0, 0, 10), rect(CP, 0, -3, 20, 6)],
    'circle/circle': (CP) => [circle(CP, 0, 0, 10), circle(CP, 7, 3, 8)],
    'tri/circle'   : (CP) => [triangle(CP, 0, -5, 12), circle(CP, 0, 4, 6)],
    'nested'       : (CP) => [rect(CP, 0, 0, 20, 20), rect(CP, 5, 5, 5, 5)],
    'disjoint'     : (CP) => [rect(CP, 0, 0, 5, 5), circle(CP, 20, 20, 3)],
};

function checkGolden(code) {
    const CP = loadCompoundPath(code);
    for (const [name, make] of Object.entries(PAIRS)) {
        for (const op of ['unite', 'intersect', 'subtract', 'exclude', 'divide']) {
            const [a, b] = make(CP);
            const result = op === 'divide' ? CP.prototype['divide']([a, b]) : a[op](b);
            // JSON round-trip: arrays from the vm context have another realm's prototype.
            assert.deepEqual(JSON.parse(JSON.stringify(dump(result))), golden[`${name} ${op}`], `${name} ${op}`);
        }
    }
}

test('path-boolean.js reproduces the results of the Closure-built min file', () => {
    checkGolden(fs.readFileSync(path.join(SRC_ROOT, SOURCE), 'utf8'));
});

test('the deployed path-boolean.js keeps the paper.js notice and its results', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'path-boolean-'));
    try {
        const out = path.join(dir, 'path-boolean.js');
        await deployJsFile(path.join(SRC_ROOT, SOURCE), out);
        const code = fs.readFileSync(out, 'utf8');
        assert.match(code, /Copyright \(c\) 2011 - 2020, J.*rg Lehni & Jonathan Puckey/);
        assert.match(code, /The MIT License/);
        checkGolden(code);
    } finally {
        fs.rmSync(dir, { recursive: true, force: true });
    }
});

test('the lazy loader and the deploy step use path-boolean.js', () => {
    const loader = fs.readFileSync(path.join(SRC_ROOT, 'common/editorscommon.js'), 'utf8');
    const deploy = fs.readFileSync(path.join(SRC_ROOT, 'build/scripts/deploy-assets.cjs'), 'utf8');
    assert.match(loader, /loadScript\('\.\.\/\.\.\/\.\.\/\.\.\/sdkjs\/common\/Drawings\/Format\/path-boolean\.js'/);
    assert.match(deploy, /'Drawings\/Format\/path-boolean\.js'/);
    assert.equal(fs.existsSync(path.join(SRC_ROOT, 'common/Drawings/Format/path-boolean-min.js')), false);
});
