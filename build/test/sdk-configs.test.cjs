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

const { getFilesMin, getFilesAll, mergeConfigs, expandGlobs, globFiles } = require('../lib/sdk-configs.cjs');

test('getFilesMin: desktop min falls back to [] when desktop config omits "min"', () => {
    const sdkCfg = { min: ['a.js'], desktop: { common: ['b.js'] } };
    assert.deepEqual(getFilesMin(sdkCfg, 'desktop'), ['a.js']);
});

test('getFilesMin: mobile_banners min falls back to [] when omitted', () => {
    const sdkCfg = { min: ['a.js'], mobile_banners: { common: ['banner.js'] } };
    assert.deepEqual(getFilesMin(sdkCfg, 'mobile'), ['a.js']);
});

test('getFilesMin: mobile_banners min is prepended when present', () => {
    const sdkCfg = { min: ['a.js'], mobile_banners: { min: ['banner-min.js'] } };
    assert.deepEqual(getFilesMin(sdkCfg, 'mobile'), ['banner-min.js', 'a.js']);
});

test('getFilesAll: desktop common falls back to [] when omitted', () => {
    const sdkCfg = { common: ['a.js'], desktop: { min: ['b.js'] } };
    assert.deepEqual(getFilesAll(sdkCfg, 'desktop'), ['a.js']);
});

test('getFilesAll: mobile_banners common falls back to [] when omitted', () => {
    const sdkCfg = { common: ['a.js'], mobile_banners: { min: ['banner-min.js'] } };
    assert.deepEqual(getFilesAll(sdkCfg, 'mobile'), ['a.js']);
});

test('getFilesAll: mobile_banners common is prepended, then exclude_mobile filters, then mobile appends', () => {
    const sdkCfg = {
        common: ['a.js', 'skip.js'],
        mobile_banners: { common: ['banner.js'] },
        exclude_mobile: ['skip.js'],
        mobile: ['mobile-only.js'],
    };
    assert.deepEqual(
        getFilesAll(sdkCfg, 'mobile'),
        ['banner.js', 'a.js', 'mobile-only.js']
    );
});

test('mergeConfigs: array properties concatenate', () => {
    const base = { min: ['a.js'] };
    mergeConfigs(base, { min: ['b.js'] });
    assert.deepEqual(base.min, ['a.js', 'b.js']);
});

test('mergeConfigs: nested object properties merge recursively', () => {
    const base = { desktop: { min: ['a.js'] } };
    mergeConfigs(base, { desktop: { common: ['b.js'] } });
    assert.deepEqual(base, { desktop: { min: ['a.js'], common: ['b.js'] } });
});

test('expandGlobs: globs expand to sorted files only, literal entries pass through', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'sdk-configs-glob-'));
    try {
        for (const name of ['b.js', 'a.js', 'c.txt']) fs.writeFileSync(path.join(dir, name), '');
        fs.mkdirSync(path.join(dir, 'sub.js'));
        const watched = [];
        const result  = expandGlobs(['/literal.js', path.join(dir, '*.js')], (d) => watched.push(d));
        assert.deepEqual(result, ['/literal.js', path.join(dir, 'a.js'), path.join(dir, 'b.js')]);
        assert.deepEqual(watched, [dir]);
        assert.deepEqual(globFiles('{a,c}.*', { cwd: dir }).sort(), [path.join(dir, 'a.js'), path.join(dir, 'c.txt')]);
    } finally {
        fs.rmSync(dir, { recursive: true, force: true });
    }
});
