/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// doctrenderer (x2t, docbuilder, desktop converter) loads common/Native/native.js
// and XRegExp, then the editor bundle, with no jQuery. The bundle must load
// there: its top level may not reference a global jQuery or $.
// Sources are concatenated the way build/loaders/sdk-concat.cjs does.

const test   = require('node:test');
const assert = require('node:assert/strict');
const fs     = require('node:fs');
const path   = require('node:path');
const vm     = require('node:vm');
const { loadAllConfigs, getFilesMin, getFilesAll, expandGlobs } = require('../lib/sdk-configs.cjs');

const SRC_ROOT = path.resolve(__dirname, '..', '..');
const configs  = loadAllConfigs(SRC_ROOT, []);
const read     = (file) => fs.readFileSync(file, 'utf8');
const concat   = (files) => expandGlobs(files).map((file) => ';' + read(file)).join('\n');

function loadEditor(module, platform) {
    const cfg = configs[module].sdk;
    const ctx = vm.createContext({});
    vm.runInContext('var window = this, self = this, native;', ctx);
    vm.runInContext(read(path.join(SRC_ROOT, 'common/Native/native.js')), ctx);
    vm.runInContext(read(path.join(SRC_ROOT, 'vendor/xregexp-all-min.js')), ctx);
    vm.runInContext(concat(getFilesMin(cfg, platform)), ctx);
    vm.runInContext('(function(window, undefined) {\n' + concat(getFilesAll(cfg, platform)) + '\n})(window);', ctx);
    return ctx;
}

for (const module of ['word', 'cell', 'slide', 'visio']) {
    for (const platform of ['', 'desktop']) {
        test(`${module}${platform ? ' ' + platform : ''} loads in doctrenderer without jQuery`, () => {
            const ctx = loadEditor(module, platform);
            assert.equal(vm.runInContext('typeof jQuery + typeof $', ctx), 'undefinedundefined');
        });
    }
}

test('stacked bar chart data is computed without jQuery', () => {
    const ctx    = loadEditor('cell', '');
    const proto  = ctx.AscFormat.CChartsDrawer.prototype;
    const drawer = Object.create(proto);
    drawer.getChartGrouping = () => 'stacked';
    drawer._getChartType    = () => ctx.AscFormat.c_oChartTypes.Bar;

    const data = [[1, 2, 3], [-1, 4, -2]];
    proto._calculateStackedData2.call(drawer, data, {});
    assert.deepEqual(data, [[1, 3, 6], [-1, 4, -3]]);
});

test('word page reads its offset without jQuery', () => {
    const ctx = loadEditor('word', '');
    ctx.Asc.editor = { isPdfEditor: () => false };
    const page = new ctx.AscCommonWord.CEditorPage({});
    page.checkBodySize();
    assert.deepEqual([page.X, page.Y], [0, 0]);
});
