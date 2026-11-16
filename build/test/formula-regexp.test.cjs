/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// The spreadsheet formula parser recognises sheet references, defined names,
// functions, table structured references and pivot references with native
// RegExps (formerly XRegExp). The expected values below were recorded with the
// XRegExp 3.0.0 implementation, whose BMP-only Unicode classes are kept as explicit ranges.
// The cell bundle is loaded the doctrenderer way: no DOM, no XRegExp.

const test   = require('node:test');
const assert = require('node:assert/strict');
const fs     = require('node:fs');
const path   = require('node:path');
const vm     = require('node:vm');
const { loadAllConfigs, getFilesMin, getFilesAll, expandGlobs } = require('../lib/sdk-configs.cjs');

const SRC_ROOT = path.resolve(__dirname, '..', '..');
const read     = (file) => fs.readFileSync(file, 'utf8');
const concat   = (files) => expandGlobs(files).map((file) => ';' + read(file)).join('\n');

function loadCell() {
	const cfg = loadAllConfigs(SRC_ROOT, [])['cell'].sdk;
	const ctx = vm.createContext({});
	vm.runInContext('var window = this, self = this, native;', ctx);
	vm.runInContext(read(path.join(SRC_ROOT, 'common/Native/native.js')), ctx);
	vm.runInContext(concat(getFilesMin(cfg, '')), ctx);
	vm.runInContext('(function(window, undefined) {\n' + concat(getFilesAll(cfg, '')) + '\n})(window);', ctx);
	vm.runInContext('Asc.editor = {DocInfo: null};', ctx);
	return ctx;
}

const ctx = loadCell();
const P   = ctx.AscCommon.parserHelp;

const ASTRAL_LETTER = '\u{1D400}';
const ASTRAL_MARK   = '\u{1D167}';

test('bundle runs without XRegExp', () => {
	assert.equal(typeof ctx.XRegExp, 'undefined');
});

// The doctrenderer V8 is built without ICU (v8_enable_i18n_support=false), where \p{...} is a
// SyntaxError that breaks the whole bundle. Only IsLetter may use it, inside try/catch.
test('no Unicode property escapes in bundled sources', () => {
	const files = new Set();
	for (const editor of Object.values(loadAllConfigs(SRC_ROOT, []))) {
		for (const file of expandGlobs([...getFilesMin(editor.sdk, ''), ...getFilesAll(editor.sdk, '')])) {
			files.add(file);
		}
	}
	const hits = [];
	for (const file of files) {
		read(file).split('\n').forEach((line, i) => {
			if (/\\{1,2}[pP]\{/.test(line) && !line.trimStart().startsWith('//')) {
				hits.push(path.relative(SRC_ROOT, file) + ':' + (i + 1));
			}
		});
	}
	assert.deepEqual(hits.map((hit) => hit.replace(/:\d+$/, '')), ['common/editorscommon.js'], hits.join('\n'));
});

test('is3DRef: sheet names', () => {
	const cases = [
		// formula, support_digital_start, [found, sheet from, sheet to], consumed length
		['Sheet1!A1', false, [true, 'Sheet1', null], 7],
		['Sheet1:Sheet3!A1', false, [true, 'Sheet1', 'Sheet3'], 14],
		["'My Sheet'!A1", false, [true, 'My Sheet', null], 11],
		["'It''s'!A1", false, [true, "It's", null], 8],
		["'Sheet 1:Sheet 2'!A1:B2", false, [true, 'Sheet 1', 'Sheet 2'], 18],
		['Лист1:Лист2!B3', false, [true, 'Лист1', 'Лист2'], 12],
		['表1!A1', false, [true, '表1', null], 3],
		['é!A1', false, [true, 'é', null], 3],
		['a.b.c!A1', false, [true, 'a.b.c', null], 6],
		['1Sheet!A1', false, [false, null, null], 0],
		['1Sheet!A1', true, [true, '1Sheet', null], 7],
		['A1', false, [false, null, null], 0],
		// the Unicode classes are BMP only, as in XRegExp 3.0.0
		[ASTRAL_LETTER + '!A1', false, [false, null, null], 0],
		['x' + ASTRAL_MARK + '!A1', false, [false, null, null], 0],
	];
	for (const [formula, digitalStart, expected, length] of cases) {
		const r = P.is3DRef(formula, 0, digitalStart);
		assert.deepEqual([r[0], r[1], r[2]], expected, formula);
		assert.equal(r[0] ? P.pCurrPos : 0, length, formula + ' length');
	}
	assert.equal(P.is3DRef('[1]Sheet1!A1', 0)[3], '1', 'external link index');
	assert.equal(P.is3DRef('!Name', 0)[4].defname, 'Name', 'short external link');
});

test('isName and isFunc', () => {
	const names = [
		['Name+1', true, 'Name'],
		['my.name)', true, 'my.name'],
		['Имя ', true, 'Имя'],
		['_x', true, '_x'],
		['TRUE', false, 'TRUE'],
		['1abc', false, null],
		['x' + ASTRAL_MARK, false, null],
	];
	for (const [formula, ok, operand] of names) {
		assert.equal(P.isName(formula, 0), ok, formula);
		assert.equal(P.operand_str, operand, formula + ' operand');
	}
	const funcs = [
		['SUM(A1)', 'SUM'],
		['_xlfn.XLOOKUP(', '_xlfn.XLOOKUP'],
		['F.DIST(1', 'F.DIST'],
		['СУММ(A1)', 'СУММ'],
		['SUM (A1)', 'SUM '],
		['A1+B1', 'A1'],
		['"str"', null],
	];
	for (const [formula, operand] of funcs) {
		assert.equal(P.isFunc(formula, 0), operand !== null, formula);
		if (operand !== null) {
			assert.equal(P.operand_str, operand, formula + ' operand');
		}
	}
});

const TABLE_GROUPS = ['tableName', 'columnName1', 'hdtcc', 'hdt', 'hdtcstart', 'hdtcend', 'reservedColumn', 'columnRange', 'colStart', 'colEnd', 'oneColumn'];
function tableGroups(match) {
	const out = {};
	for (const name of TABLE_GROUPS) {
		if (match[name] !== undefined) {
			out[name] = match[name];
		}
	}
	return out;
}

test('isTable: structured references expose named groups on the match', () => {
	const cases = [
		['Table1[Col]', { tableName: 'Table1', columnName1: 'Col', oneColumn: 'Col' }],
		['Table1[]', { tableName: 'Table1' }],
		['Table1[#All]', { tableName: 'Table1', columnName1: '#All', reservedColumn: '#All' }],
		['table1[#all]', { tableName: 'table1', columnName1: '#all', reservedColumn: '#all' }],
		['Table1[@]', { tableName: 'Table1', columnName1: '@', hdtcc: '@', hdt: '@' }],
		['Table1[[#This Row],[Col]]', { tableName: 'Table1', columnName1: '[#This Row],[Col]', hdtcc: '[#This Row],[Col]', hdt: '[#This Row]', hdtcstart: 'Col' }],
		['Table1[[Column1]:[Column2]]', { tableName: 'Table1', columnName1: '[Column1]:[Column2]', columnRange: '[Column1]:[Column2]', colStart: 'Column1', colEnd: 'Column2' }],
		['Table1[[#Headers],[#Data],[Column1]:[Column2]]', { tableName: 'Table1', columnName1: '[#Headers],[#Data],[Column1]:[Column2]', hdtcc: '[#Headers],[#Data],[Column1]:[Column2]', hdt: '[#Headers],[#Data]', hdtcstart: 'Column1', hdtcend: 'Column2' }],
		['Table1[[#Data],[#Totals]]', { tableName: 'Table1', columnName1: '[#Data],[#Totals]', hdtcc: '[#Data],[#Totals]', hdt: '[#Data],[#Totals]' }],
		["Table1['[Col']", { tableName: 'Table1', columnName1: "'[Col'", oneColumn: "'[Col'" }],
		['Табл1[Кол]', { tableName: 'Табл1', columnName1: 'Кол', oneColumn: 'Кол' }],
	];
	for (const [formula, expected] of cases) {
		const match = P.isTable(formula, 0, false);
		assert.ok(match, formula);
		assert.deepEqual(tableGroups(match), expected, formula);
		assert.equal(P.operand_str, formula, formula + ' operand');
	}
	assert.equal(P.isTable('Sheet1!A1', 0, false), false);
});

test('isTable: localized names and ";" argument separator', () => {
	ctx.AscCommon.build_local_rx({ StructureTables: { h: 'Заголовки', d: 'Данные', a: 'Все', tr: 'Эта строка', t: 'Итоги' } });
	P.setDigitSeparator(',');
	try {
		let match = P.isTable('Таблица1[[#Эта строка];[Кол]]', 0, true, {});
		assert.deepEqual(tableGroups(match), { tableName: 'Таблица1', columnName1: '[#Эта строка];[Кол]', hdtcc: '[#Эта строка];[Кол]', hdt: '[#Эта строка]', hdtcstart: 'Кол' });
		match = P.isTable('Таблица1[[#Заголовки];[#Данные]]', 0, true, {});
		assert.equal(match.hdt, '[#Заголовки];[#Данные]');
		match = P.isTable('Таблица1[#Все]', 0, true, {});
		assert.equal(match.reservedColumn, '#Все');
		assert.equal(P.isTable('Table1[[#This Row],[Col]]', 0, false).hdt, '[#This Row]', 'non-local table regexp keeps the English names');
	} finally {
		P.setDigitSeparator('.');
		ctx.AscCommon.build_local_rx(null);
	}
});

test('pivot references', () => {
	assert.deepEqual(Array.from(P.isPivotRaw("'Field 1'[Item]", 0)), ["'Field 1'", 'Item']);
	assert.deepEqual(Array.from(P.isPivotRaw("Field['Item 1']", 0)), ['Field', "'Item 1'"]);
	assert.deepEqual(Array.from(P.isPivotRaw('Поле[Элемент]', 0)), ['Поле', 'Элемент']);
	assert.equal(P.isPivotRaw('1f[x]', 0), false);
	const names = [['Fi.eld', 'a(b'], ['It+em']];
	assert.deepEqual(Array.from(P.isPivot('Fi.eld[It+em]', 0, false, names)), ['Fi.eld', 'It+em'], 'names are matched literally');
	assert.equal(P.isPivot('Fixeld[Ittem]', 0, false, names), false, 'regexp metacharacters in names are escaped');
	const toFormula = ctx.Asc.CT_pivotTableDefinition.prototype.asc_convertNameToFormula;
	assert.equal(toFormula.call({}, 'Поле_1'), 'Поле_1');
	assert.equal(toFormula.call({}, 'Name 1'), "'Name 1'");
	assert.equal(toFormula.call({}, "It's"), "'It''s'");
});

test('complex numbers', () => {
	const parse = (s) => {
		const c = ctx.AscCommonExcel.Complex.prototype.parseComplexStr(s);
		return c instanceof ctx.AscCommonExcel.Complex ? [c.real, c.img, c.suffix] : 'error';
	};
	assert.deepEqual(parse('-1+2i'), [-1, 2, 'i']);
	assert.deepEqual(parse('3.5e2-4.1E-1j'), [350, -0.41, 'j']);
	assert.deepEqual(parse('4-3.5j'), [4, -3.5, 'j']);
	assert.deepEqual(parse('7'), [7, 0, 'i']);
	assert.deepEqual(parse('1+2k'), [1, 0, 'i'], 'only the leading number is read');
	assert.equal(parse('abc'), 'error');
});

test('cell editor word navigation', () => {
	const render = new ctx.AscCommonExcel.CellTextRender();
	render.getEndOfLine = () => -1;
	const words = (text) => {
		render.chars = Array.from(text, (ch) => ch.codePointAt(0));
		return [render.getNextWord(0), render.getPrevWord(text.length)];
	};
	assert.deepEqual(words('ab cd'), [3, 3]);
	assert.deepEqual(words("it's me"), [5, 5], 'apostrophe is part of a word');
	assert.deepEqual(words('x Пример'), [2, 2]);
	assert.deepEqual(words('x \u{1F600}ab'), [3, 3], 'emoji before a word');
	assert.deepEqual(words('abc'), [-1, 0]);
});
