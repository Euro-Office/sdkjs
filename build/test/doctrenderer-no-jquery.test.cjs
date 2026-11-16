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

test('spreadsheet target position adds the document offset of the cell cursor', () => {
    const ctx    = loadEditor('cell', '');
    const cursor = {
        getClientRects       : () => [{}],
        getBoundingClientRect: () => ({ left: 10, top: 20 }),
        ownerDocument        : { defaultView: { pageXOffset: 5, pageYOffset: 7 } },
    };
    const api = Object.create(ctx.AscCommon.baseEditorsApi.prototype);
    api.editorId            = ctx.AscCommon.c_oEditorId.Spreadsheet;
    api.asc_getCellInfo     = () => ({ asc_getSelectionType: () => -1 });
    api.asc_getCellEditMode = () => true;
    api.wb = { cellEditor: { curLeft: 1, curTop: 2, curHeight: 3, cursor } };

    const ret = api.getTargetOnBodyCoords();
    assert.deepEqual([ret.X, ret.Y, ret.TargetH], [16, 29, 3]);

    cursor.getClientRects = () => [];
    const hidden = api.getTargetOnBodyCoords();
    assert.deepEqual([hidden.X, hidden.Y], [1, 2]);
});

test('editor sources do not use jQuery', () => {
    const usesJQuery = /\bjQuery\b|(?<![\w$.\\])\$\(|(?<![\w$\\])\$\.[A-Za-z]/;
    const isComment  = /^\s*(\/\/|\/\*|\*)/;
    const files      = new Set();
    for (const module of Object.keys(configs)) {
        for (const platform of ['', 'desktop', 'mobile']) {
            const cfg = configs[module].sdk;
            expandGlobs([...getFilesMin(cfg, platform), ...getFilesAll(cfg, platform)]).forEach((file) => files.add(file));
        }
    }
    const hits = [];
    for (const file of files) {
        if (file.includes(`${path.sep}vendor${path.sep}`)) {
            continue;
        }
        read(file).split('\n').forEach((line, i) => {
            if (usesJQuery.test(line) && !isComment.test(line)) {
                hits.push(`${path.relative(SRC_ROOT, file)}:${i + 1}`);
            }
        });
    }
    assert.deepEqual(hits, []);
});
