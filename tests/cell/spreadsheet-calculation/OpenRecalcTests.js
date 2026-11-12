/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

// Which formulas opening a file recalculates. Each test builds a sheet, sets the cached values a file would
// carry, writes it to Editor.bin and opens it again through BinaryFileReader, like the editor opens a file.
$(function () {

	const CellValueType = AscCommon.CellValueType;

	Asc.spreadsheet_api.prototype._init = function () {
		this.isLoadFullApi = true;
	};

	const api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function () {
		}
	};
	const docInfo = new Asc.asc_CDocInfo();
	docInfo.asc_putTitle("TeSt.xlsx");
	api.DocInfo = docInfo;
	api.initCollaborativeEditing({});
	window["Asc"]["editor"] = api;
	AscCommon.g_oTableId.init();

	function openWorkbook(data) {
		// History still belongs to the previous workbook, do not record this one's creation there
		const wb = AscFormat.ExecuteNoHistory(function () {
			return new AscCommonExcel.Workbook(new AscCommonExcel.asc_CHandlersList(), api, true);
		}, this, [], true);
		api.wbModel = wb;
		AscCommon.History.init(wb);
		AscCommonExcel.g_oUndoRedoCell = new AscCommonExcel.UndoRedoCell(wb);
		AscCommonExcel.g_oUndoRedoWorksheet = new AscCommonExcel.UndoRedoWoorksheet(wb);
		AscCommonExcel.g_oUndoRedoWorkbook = new AscCommonExcel.UndoRedoWorkbook(wb);
		AscCommonExcel.g_oUndoRedoCol = new AscCommonExcel.UndoRedoRowCol(wb, false);
		AscCommonExcel.g_oUndoRedoRow = new AscCommonExcel.UndoRedoRowCol(wb, true);
		AscCommonExcel.g_oUndoRedoComment = new AscCommonExcel.UndoRedoComment(wb);
		AscCommonExcel.g_oUndoRedoAutoFilters = new AscCommonExcel.UndoRedoAutoFilters(wb);
		AscCommonExcel.g_DefNameWorksheet = new AscCommonExcel.Worksheet(wb, -1);
		AscCommon.g_oIdCounter.Set_Load(false);

		const oBinaryFileReader = new AscCommonExcel.BinaryFileReader();
		oBinaryFileReader.Read(data, wb);
		return wb;
	}

	openWorkbook(AscCommon.getEmpty());
	AscCommonExcel.getFormulasInfo();

	// B1:B4 = x, y, z, (empty); A1:A4 = {=B1:B4}; C1 = A4; C2 = A2; C3 = B1&"!"
	function fillSheet(ws, arrayFormula) {
		ws.getRange2("B1").setValue("x");
		ws.getRange2("B2").setValue("y");
		ws.getRange2("B3").setValue("z");
		const arrayRange = ws.getRange2("A1:A4");
		arrayRange.setValue(arrayFormula || "=B1:B4", null, null, arrayRange.bbox);
		ws.getRange2("C1").setValue("=A4");
		ws.getRange2("C2").setValue("=A2");
		ws.getRange2("C3").setValue("=B1&\"!\"");
	}

	// Replaces the value a cell carries into the file, null writes the cell without a value
	function setCachedValue(ws, ref, value) {
		const bbox = AscCommonExcel.g_oRangeCache.getAscRange(ref);
		ws._getCell(bbox.r1, bbox.c1, function (cell) {
			if (null === value) {
				cell.cleanText();
			} else {
				cell.setTypeInternal(CellValueType.String);
				cell.setValueTextInternal(value);
			}
		});
	}

	function saveAndOpen(wb) {
		return openWorkbook(new AscCommonExcel.BinaryFileWriter(wb).Write());
	}

	// Builds a workbook with prepare(wb, ws), writes it to Editor.bin and opens that
	function reopen(prepare) {
		const sourceWb = openWorkbook(AscCommon.getEmpty());
		prepare(sourceWb, sourceWb.getWorksheet(0));
		return saveAndOpen(sourceWb);
	}

	// Cells that opening marked for recalculation
	function getCellsToRecalc(wb) {
		const names = {};
		wb.dependencyFormulas._foreachChanged(function (cell) {
			if (cell) {
				names[cell.getName()] = true;
			}
		});
		return Object.keys(names).sort();
	}

	function getValues(ws, refs) {
		return refs.map(function (ref) {
			return ws.getRange2(ref).getValue();
		});
	}

	// The values stored in the cells, read without the getters that calculate a dirty cell
	function getStoredValues(ws, refs) {
		return refs.map(function (ref) {
			const bbox = AscCommonExcel.g_oRangeCache.getAscRange(ref);
			let res;
			ws._getCellNoEmpty(bbox.r1, bbox.c1, function (cell) {
				if (cell.getIsDirty()) {
					res = "dirty";
				} else {
					res = CellValueType.String === cell.type ? cell.text : cell.number;
				}
			});
			return res;
		});
	}

	QUnit.module("Recalculation on open");

	QUnit.test("Blank member of a calculated array keeps its cached blank, formulas read it as empty", function (assert) {
		// Google Sheets writes a blank array result, e.g. of UNIQUE over a column with an empty cell, as a cell
		// without a value, and saves no fullCalcOnLoad
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			ws.getRange2("E1").setValue("=A4+F1");
			ws.getRange2("F1").setValue("0");
			setCachedValue(ws, "A4", null);
			setCachedValue(ws, "C1", "cached"); // =A4
			setCachedValue(ws, "C3", "cached");
			setCachedValue(ws, "E1", "cached");
		});
		const ws = wb.getWorksheet(0);

		assert.deepEqual(getCellsToRecalc(wb), [], "Nothing is marked for recalculation");
		wb.dependencyFormulas.calcTree();
		assert.deepEqual(getStoredValues(ws, ["A1", "A2", "A3", "A4"]), ["x", "y", "z", null], "Array keeps its cached values, A4 stays blank");
		assert.deepEqual(getValues(ws, ["A4", "C1", "C3"]), ["", "cached", "cached"], "A4 and the dependents keep their cached values");

		ws.getRange2("F1").setValue("1");
		assert.strictEqual(ws.getRange2("E1").getValue(), "1", "Editing another input of a dependent: A4 is read as empty, not #NUM!");
		ws.getRange2("G1").setValue("=A4");
		ws.getRange2("G2").setValue("=COUNTA(A1:A4)");
		assert.deepEqual(getValues(ws, ["G1", "G2"]), ["0", "3"], "Formulas typed after open read A4 as empty");

		ws.getRange2("B4").setValue("w");
		assert.deepEqual(getValues(ws, ["A4", "C1", "G2"]), ["w", "w", "4"], "Editing an input recalculates the array and its dependents");
	});

	QUnit.test("Blank head of a calculated array is recalculated, blank members keep their cached blank", function (assert) {
		// A head without a value is marked like any formula without a value (setFormulaOpen)
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			setCachedValue(ws, "A1", null);
			setCachedValue(ws, "A2", "cached");
			setCachedValue(ws, "A4", null);
		});
		const ws = wb.getWorksheet(0);

		assert.deepEqual(getCellsToRecalc(wb), ["A1"], "Only the head is marked");
		wb.dependencyFormulas.calcTree();
		assert.deepEqual(getValues(ws, ["A1", "A2", "A4"]), ["x", "cached", ""], "Head is calculated, members keep their cached values");
	});

	QUnit.test("Array whose only value is its head is recalculated with its dependents", function (assert) {
		// A two-cell array with a blank second result looks like a file that saved only the head value: the open
		// rule cannot tell them apart and keeps the previous behaviour for both
		const wb = reopen(function (wb, ws) {
			ws.getRange2("B1").setValue("x");
			const arrayRange = ws.getRange2("A1:A2");
			arrayRange.setValue("=B1:B2", null, null, arrayRange.bbox);
			ws.getRange2("C1").setValue("=A2");
			setCachedValue(ws, "A2", null);
			setCachedValue(ws, "C1", "cached");
		});
		const ws = wb.getWorksheet(0);

		assert.deepEqual(getCellsToRecalc(wb), ["A2"], "The member without a value is marked");
		wb.dependencyFormulas.calcTree();
		assert.deepEqual(getValues(ws, ["A1", "A2", "C1"]), ["x", "0", "0"], "Member and dependent are calculated");
	});

	QUnit.test("Blank member takes the calculated value once its array is recalculated", function (assert) {
		// sdkjs calculates an empty array result as 0, like Excel; until then the cached blank is shown
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			setCachedValue(ws, "A4", null);
			setCachedValue(ws, "C1", "cached"); // =A4
		});
		const ws = wb.getWorksheet(0);
		wb.dependencyFormulas.calcTree();
		assert.deepEqual(getValues(ws, ["A4", "C1"]), ["", "cached"], "Cached values after open");

		ws.getRange2("B1").setValue("q");
		assert.deepEqual(getValues(ws, ["A1", "A4", "C1"]), ["q", "0", "0"], "Editing an input recalculates the whole array and its dependents");
	});

	QUnit.test("setFormulaOpen marks an array head without a value", function (assert) {
		// Readers that have no array step (the JSON import) rely on it
		const wb = openWorkbook(AscCommon.getEmpty());
		const ws = wb.getWorksheet(0);
		const manager = new AscCommonExcel.InitOpenManager(null, wb);
		const formula = new AscCommonExcel.OpenFormula();
		formula.t = Asc.ECellFormulaType.cellformulatypeArray;
		formula.ref = "A1:A4";
		formula.v = "B1:B4";
		const cell = new AscCommonExcel.Cell(ws);
		cell.setRowCol(0, 0);
		const tmp = {ws: ws, cell: cell, formula: formula, sharedFormulas: {}, prevFormulas: {}, siFormulas: {}, formulaArray: []};
		try {
			manager.setFormulaOpen(tmp);
			assert.deepEqual(getCellsToRecalc(wb), ["A1"], "Head is marked");
		} finally {
			wb.dependencyFormulas.changedCell = null;
			wb.dependencyFormulas.changedRange = null;
		}
	});

	QUnit.test("setArrayFormulasOpen marks a blank head of a calculated array on its own (the XML open path)", function (assert) {
		// The binary reader marks a blank head through setFormulaOpen, per cell. The XML reader (sdkjs-ooxml) has no
		// such step, so setArrayFormulasOpen must mark the head itself. Reopen a calculated array with a blank head,
		// clear what setFormulaOpen marked, and run the array loop alone: it must still mark the head, not the members.
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			setCachedValue(ws, "A1", null); // blank head
			setCachedValue(ws, "A2", "cached"); // a member with a value: the array was calculated
			setCachedValue(ws, "A4", null); // blank member
		});
		const ws = wb.getWorksheet(0);
		wb.dependencyFormulas.changedCell = null;
		wb.dependencyFormulas.changedRange = null;

		let parsed;
		ws._getCell(0, 0, function (cell) {
			parsed = cell.formulaParsed;
		});
		const manager = new AscCommonExcel.InitOpenManager(null, wb);
		manager.setArrayFormulasOpen({ws: ws, formulaArray: [parsed]});

		assert.deepEqual(getCellsToRecalc(wb), ["A1"], "The loop marks the blank head, not its blank members");
	});

	QUnit.test("Formulas without a value that read a blank array member are calculated from it", function (assert) {
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			setCachedValue(ws, "A4", null);
			setCachedValue(ws, "C1", null); // =A4
		});
		const ws = wb.getWorksheet(0);

		assert.deepEqual(getCellsToRecalc(wb), ["C1"], "Only the formula without a value is marked");
		wb.dependencyFormulas.calcTree();
		assert.deepEqual(getValues(ws, ["A4", "C1"]), ["", "0"], "It reads the blank member as empty, not #NUM!");
	});

	QUnit.test("Array formula that saved only its head value is recalculated with its dependents", function (assert) {
		// A writer may save the head's value only: the array may never have been calculated
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			["A2", "A3", "A4"].forEach(function (ref) {
				setCachedValue(ws, ref, null);
			});
			setCachedValue(ws, "C2", "cached"); // =A2
		});
		const ws = wb.getWorksheet(0);

		assert.deepEqual(getCellsToRecalc(wb), ["A2", "A3", "A4"], "Members without a value are marked");
		wb.dependencyFormulas.calcTree();
		assert.deepEqual(getValues(ws, ["A1", "A2", "A3", "A4"]), ["x", "y", "z", "0"], "Members are calculated");
		assert.strictEqual(ws.getRange2("C2").getValue(), "y", "Their dependents are recalculated");
	});

	QUnit.test("Blank members of a dynamic array (cm) are recalculated with their dependents", function (assert) {
		// Calculating a dynamic array may resize it, so it keeps the rule of an array that was never calculated
		const isSupportDynamicArrays = AscCommonExcel.bIsSupportDynamicArrays;
		AscCommonExcel.bIsSupportDynamicArrays = true;
		try {
			const wb = reopen(function (wb, ws) {
				fillSheet(ws);
				setCachedValue(ws, "A4", null);
				ws._getCell(0, 0, function (cell) {
					cell.getFormulaParsed().cm = 1;
				});
			});
			let cm;
			wb.getWorksheet(0)._getCell(0, 0, function (cell) {
				cm = cell.getFormulaParsed().cm;
			});
			assert.strictEqual(cm, 1, "cm is saved and read back");
			assert.deepEqual(getCellsToRecalc(wb), ["A4"], "The blank member is marked");
			wb.dependencyFormulas.calcTree();
			const ws = wb.getWorksheet(0);
			assert.deepEqual(getValues(ws, ["A1", "A2", "A3", "A4"]), ["x", "y", "z", "0"], "The blank member is calculated");
			assert.strictEqual(ws.getRange2("C1").getValue(), "0", "Its dependent is recalculated");
		} finally {
			AscCommonExcel.bIsSupportDynamicArrays = isSupportDynamicArrays;
		}
	});

	QUnit.test("Two-dimensional array: only its blank members keep their cached blank", function (assert) {
		const wb = reopen(function (wb, ws) {
			["x", "y", "z", "u", "v"].forEach(function (v, i) {
				ws.getRange2(["B1", "C1", "B2", "C2", "B3"][i]).setValue(v);
			});
			const arrayRange = ws.getRange2("E1:F3");
			arrayRange.setValue("=B1:C3", null, null, arrayRange.bbox);
			setCachedValue(ws, "F3", null);
		});
		const ws = wb.getWorksheet(0);

		assert.deepEqual(getCellsToRecalc(wb), [], "Nothing is marked for recalculation");
		wb.dependencyFormulas.calcTree();
		assert.deepEqual(getValues(ws, ["E1", "F1", "E2", "F2", "E3", "F3"]), ["x", "y", "z", "u", "v", ""], "Cached values are kept");
	});

	QUnit.test("Horizontal array on another sheet keeps its blank member", function (assert) {
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			const ws2 = wb.createWorksheet(1, "Second");
			["x", "y", "z"].forEach(function (v, i) {
				ws2.getRange2(["A1", "B1", "C1"][i]).setValue(v);
			});
			const arrayRange = ws2.getRange2("A2:D2");
			arrayRange.setValue("=A1:D1", null, null, arrayRange.bbox);
			ws2.getRange2("A3").setValue("=D2");
			// only its head has a value: recalculated, on this sheet
			ws2.getRange2("F1").setValue("v");
			const headOnlyRange = ws2.getRange2("F2:G2");
			headOnlyRange.setValue("=F1:G1", null, null, headOnlyRange.bbox);
			setCachedValue(ws2, "D2", null);
			setCachedValue(ws2, "A3", "cached");
			setCachedValue(ws2, "G2", null);
		});
		const ws2 = wb.getWorksheetByName("Second");

		assert.deepEqual(getCellsToRecalc(wb), ["G2"], "Only the member of the head-only array is marked");
		wb.dependencyFormulas.calcTree();
		assert.deepEqual(getValues(ws2, ["A2", "B2", "C2", "D2", "A3"]), ["x", "y", "z", "", "cached"], "Cached values are kept");
		assert.deepEqual(getValues(ws2, ["F2", "G2"]), ["v", "0"], "The marked member is calculated on its own sheet");
		ws2.getRange2("A4").setValue("=D2+1");
		assert.strictEqual(ws2.getRange2("A4").getValue(), "1", "Formulas read the blank member as empty");
	});

	QUnit.test("A formula without a value that is not an array member still reads as #NUM!", function (assert) {
		// Scope guard, not an endorsement of #NUM!: only blank array results read as empty, other formula cells without
		// values keep main's behaviour. If that behaviour is changed on purpose, update this test.
		const wb = openWorkbook(AscCommon.getEmpty());
		const ws = wb.getWorksheet(0);
		ws.getRange2("B1").setValue("=1+1");
		ws._getCell(0, 1, function (cell) {
			cell.cleanText();
		});
		ws._getCellNoEmpty(0, 1, function (cell) {
			assert.strictEqual(cell.isBlankArrayResult(), false, "Not a blank array result");
		});
		const parsed = new AscCommonExcel.parserFormula("B1+1", null, ws);
		parsed.parse();
		assert.strictEqual(parsed.calculate().getValue(), "#NUM!", "B1 is read as a number without a value");
	});

	QUnit.test("An empty string is a value: it marks the array as calculated", function (assert) {
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			["A2", "A3"].forEach(function (ref) {
				setCachedValue(ws, ref, "");
			});
			setCachedValue(ws, "A4", null);
		});

		assert.deepEqual(getCellsToRecalc(wb), [], "Nothing is marked for recalculation");
	});

	QUnit.test("A cell being calculated or waiting for its calculation is not a blank array result", function (assert) {
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			setCachedValue(ws, "A4", null);
		});
		const ws = wb.getWorksheet(0);
		wb.dependencyFormulas.calcTree();

		ws._getCellNoEmpty(3, 0, function (cell) {
			assert.strictEqual(cell.isBlankArrayResult(), true, "Blank member");
			cell.setIsCalc(true);
			assert.strictEqual(cell.isBlankArrayResult(), false, "While it is calculated");
			cell.setIsCalc(false);
			// a calculation skipped by the recursion limit leaves the cell dirty
			cell._checkDirty = function () {};
			try {
				cell.setIsDirty(true);
				assert.strictEqual(cell.isBlankArrayResult(), false, "While it waits for its calculation");
			} finally {
				delete cell._checkDirty;
				cell.setIsDirty(false);
				cell._hasChanged = false;
			}
		});
	});

	QUnit.test("Saving keeps a blank array member blank", function (assert) {
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			setCachedValue(ws, "A4", null);
		});
		wb.dependencyFormulas.calcTree();

		const savedWb = saveAndOpen(wb);
		assert.deepEqual(getStoredValues(savedWb.getWorksheet(0), ["A3", "A4"]), ["z", null], "A4 is saved without a value");
		assert.deepEqual(getCellsToRecalc(savedWb), [], "Opening the saved file marks nothing");
	});

	QUnit.test("Lookups over an array with a blank member on another sheet keep their cached values", function (assert) {
		// e.g. VLOOKUP(..., Sheet1!A:B, ...) over a UNIQUE result, the case that recalculated whole workbooks on open
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			const lookupWs = wb.createWorksheet(1, "Lookup");
			lookupWs.getRange2("A1").setValue("=VLOOKUP(\"y\",Sheet1!A:A,1,FALSE)");
			lookupWs.getRange2("A2").setValue("=COUNTA(Sheet1!A:A)");
			setCachedValue(ws, "A4", null);
			setCachedValue(lookupWs, "A1", "cached");
			setCachedValue(lookupWs, "A2", "cached");
		});
		const lookupWs = wb.getWorksheetByName("Lookup");

		assert.deepEqual(getCellsToRecalc(wb), [], "Nothing is marked for recalculation");
		wb.dependencyFormulas.calcTree();
		assert.deepEqual(getValues(lookupWs, ["A1", "A2"]), ["cached", "cached"], "Lookups keep their cached values");

		wb.getWorksheet(0).getRange2("B2").setValue("v");
		assert.deepEqual(getValues(lookupWs, ["A1", "A2"]), ["#N/A", "4"], "Editing an input recalculates the lookups");
	});

	QUnit.test("Approximate and binary lookups over a blank array member read it as empty", function (assert) {
		// these searches read the cells through a Range (ws.getCell3), not a Cell
		const wb = reopen(function (wb, ws) {
			["1", "2", "3"].forEach(function (v, i) {
				ws.getRange2("B" + (i + 1)).setValue(v);
			});
			const arrayRange = ws.getRange2("A1:A4");
			arrayRange.setValue("=B1:B4", null, null, arrayRange.bbox);
			setCachedValue(ws, "A4", null);
		});
		const ws = wb.getWorksheet(0);
		wb.dependencyFormulas.calcTree();

		ws.getRange2("D1").setValue("=XLOOKUP(2.5,A1:A4,A1:A4,\"none\",-1)");
		ws.getRange2("D2").setValue("=XMATCH(3,A1:A4,0,2)");
		ws.getRange2("D3").setValue("=XMATCH(2,A1:A4,0,-2)");
		assert.deepEqual(getValues(ws, ["D1", "D2", "D3"]), ["2", "3", "2"], "Lookups are calculated");
	});

	QUnit.test("Only a blank member of an array read from a file counts as a blank array result", function (assert) {
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			setCachedValue(ws, "A4", null);
		});
		const ws = wb.getWorksheet(0);
		wb.dependencyFormulas.calcTree();

		function isBlankArrayResult(ref) {
			let res;
			ws._getCellNoEmpty(AscCommonExcel.g_oRangeCache.getAscRange(ref).r1, 0, function (cell) {
				res = cell.isBlankArrayResult();
			});
			return res;
		}
		assert.deepEqual([isBlankArrayResult("A3"), isBlankArrayResult("A4")], [false, true], "Cell: only the blank member");
		assert.strictEqual(ws.getRange2("A4").isBlankArrayResult(), true, "Range proxies the cell");
		ws.getRange2("B1").setValue("q");
		assert.strictEqual(isBlankArrayResult("A4"), false, "Not any more once the array is recalculated");
	});

	QUnit.test("Array formula without cached values is calculated on open", function (assert) {
		// e.g. a file written by a library that does not calculate formulas
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			["A1", "A2", "A3", "A4"].forEach(function (ref) {
				setCachedValue(ws, ref, null);
			});
			setCachedValue(ws, "C1", "cached");
			setCachedValue(ws, "C2", "cached");
		});
		const ws = wb.getWorksheet(0);

		assert.deepEqual(getCellsToRecalc(wb), ["A1", "A2", "A3", "A4"], "The whole array is marked for recalculation");
		wb.dependencyFormulas.calcTree();
		assert.deepEqual(getValues(ws, ["A1", "A2", "A3", "A4"]), ["x", "y", "z", "0"], "All members are calculated");
		assert.deepEqual(getValues(ws, ["C1", "C2"]), ["0", "y"], "Dependents of the members are recalculated");
	});

	QUnit.test("Several arrays on a sheet are handled each on its own", function (assert) {
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			const secondRange = ws.getRange2("D1:D4");
			secondRange.setValue("=B1:B4", null, null, secondRange.bbox);
			setCachedValue(ws, "A4", null);
			["D1", "D2", "D3", "D4"].forEach(function (ref) {
				setCachedValue(ws, ref, null);
			});
		});
		const ws = wb.getWorksheet(0);

		assert.deepEqual(getCellsToRecalc(wb), ["D1", "D2", "D3", "D4"], "Only the array that was never calculated is marked");
		wb.dependencyFormulas.calcTree();
		assert.deepEqual(getValues(ws, ["A4", "D1", "D4"]), ["", "x", "0"], "The calculated array keeps its blank member");
	});

	QUnit.test("Volatile array formula is calculated on open", function (assert) {
		const wb = reopen(function (wb, ws) {
			fillSheet(ws, "=IF(RAND()>=0,B1:B4)");
			setCachedValue(ws, "A4", null);
		});
		const ws = wb.getWorksheet(0);

		let isVolatile;
		ws._getCell(0, 0, function (cell) {
			isVolatile = cell.getFormulaParsed().ca;
		});
		assert.strictEqual(isVolatile, true, "Formula is volatile");
		assert.deepEqual(getCellsToRecalc(wb), ["A1", "A2", "A3", "A4"], "The whole array is marked for recalculation");
		wb.dependencyFormulas.calcTree();
		assert.deepEqual(getValues(ws, ["A1", "A2", "A3", "A4"]), ["x", "y", "z", "0"], "All members are calculated");
	});

	QUnit.test("fullCalcOnLoad keeps recalculating array members without values", function (assert) {
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			setCachedValue(ws, "A4", null);
			setCachedValue(ws, "C3", "cached");
			wb.calcPr.fullCalcOnLoad = true;
		});
		const ws = wb.getWorksheet(0);

		assert.strictEqual(wb.calcPr.fullCalcOnLoad, true, "Flag is read");
		assert.deepEqual(getCellsToRecalc(wb), ["A4"], "Member without a value is marked, as before");
		wb.dependencyFormulas.calcTree();
		assert.deepEqual(getValues(ws, ["A4", "C3"]), ["0", "cached"], "Member is calculated, other cached values are kept");
		assert.strictEqual(saveAndOpen(wb).calcPr.fullCalcOnLoad, true, "Flag is kept on save");
	});

	QUnit.test("fullCalcOnLoad off keeps cached values", function (assert) {
		const wb = reopen(function (wb, ws) {
			fillSheet(ws);
			setCachedValue(ws, "C3", "cached");
			wb.calcPr.fullCalcOnLoad = false;
		});
		const ws = wb.getWorksheet(0);

		assert.strictEqual(wb.calcPr.fullCalcOnLoad, false, "Flag is read");
		assert.deepEqual(getCellsToRecalc(wb), [], "Nothing is marked for recalculation");
		wb.dependencyFormulas.calcTree();
		assert.strictEqual(ws.getRange2("C3").getValue(), "cached", "Cached value is kept");
	});
});
