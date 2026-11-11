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

$(function () {

	Asc.spreadsheet_api.prototype._init = function () {
		this._loadModules();
	};
	Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
		callback();
	};
	AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () {
	};
	AscCommonExcel.WorkbookView.prototype._init = function () {
	};
	AscCommonExcel.WorkbookView.prototype._isLockedUserProtectedRange = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {
	};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {
	};
	AscCommonExcel.WorkbookView.prototype.recalculateDrawingObjects = function () {
	};
	AscCommonExcel.WorkbookView.prototype.restoreFocus = function () {
	};
	AscCommonExcel.WorksheetView.prototype._init = function () {
	};
	AscCommonExcel.WorksheetView.prototype.updateRanges = function () {
	};
	AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function () {
	};
	AscCommonExcel.WorksheetView.prototype.cleanSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype._drawSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype._scrollToRange = function () {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
	};
	AscCommonExcel.WorksheetView.prototype._initCellsArea = function () {
	};
	AscCommonExcel.WorksheetView.prototype.getZoom = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareCellTextMetricsCache = function () {
	};
	AscCommonExcel.WorksheetView.prototype._isLockedCells = function (range, subType, callback) {
		callback(true);
		return true;
	};
	AscCommonExcel.WorksheetView.prototype._isLockedAll = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorksheetView.prototype._isLockedFrozenPane = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorksheetView.prototype._updateVisibleColsCount = function () {
	};
	AscCommonExcel.WorksheetView.prototype._calcActiveCellOffset = function () {
	};
	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
	};
	Asc.ReadDefTableStyles = function () {
	};

	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function () {
		}
	};
	window["Asc"]["editor"] = api;
	AscCommon.g_oTableId.init();
	api._onEndLoadSdk();
	api.isOpenOOXInBrowser = false;
	api.OpenDocumentFromBin(null, AscCommon.getEmpty());
	api.initCollaborativeEditing({});
	api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
		api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
	api.wbModel.handlers.add("getSelectionState", function () {
		return null;
	});

	var wsView = api.wb.getWorksheet(0);
	wsView.handlers = api.handlers;
	wsView.objectRender = new AscFormat.DrawingObjects();
	var ws = api.wbModel.aWorksheets[0];

	function select(ref) {
		wsView.setSelection(ws.getRange2(ref).getBBox0());
		api.wb._updateSelectionInfo();
	}

	function cellInfoCheckbox() {
		// _onWSSelectionChanged is stubbed, so refresh the cached selection info by hand
		api.wb._updateSelectionInfo();
		return api.asc_getCellInfo().asc_getCheckbox();
	}

	function cell(ref) {
		return ws.getRange2(ref);
	}

	function hasCheckbox(ref) {
		var res = false;
		ws.getRange2(ref)._foreachNoEmpty(function (c) {
			res = c.getCheckbox();
		});
		return res;
	}

	function reset() {
		ws.getRange2("A1:D10").cleanAll();
		AscCommon.History.Clear();
		api.isViewMode = false;
	}

	QUnit.module("Cell checkbox");

	QUnit.test("insert on a range sets the flag, fills blanks with FALSE, keeps values, one undo point", function (assert) {
		reset();
		cell("A2").setValue("x");
		cell("A3").setValue("5");
		AscCommon.History.Clear();

		select("A1:A4");
		api.asc_insertCheckbox();

		["A1", "A2", "A3", "A4"].forEach(function (ref) {
			assert.strictEqual(hasCheckbox(ref), true, ref + " has checkbox");
		});
		assert.strictEqual(cell("A1").getType(), AscCommon.CellValueType.Bool, "A1 blank -> bool");
		assert.strictEqual(cell("A1").getValue(), "FALSE", "A1 blank -> FALSE");
		assert.strictEqual(cell("A4").getValue(), "FALSE", "A4 blank -> FALSE");
		assert.strictEqual(cell("A2").getValue(), "x", "A2 text kept");
		assert.strictEqual(cell("A3").getValue(), "5", "A3 number kept");
		assert.strictEqual(cellInfoCheckbox(), true, "cell info reports checkbox");

		AscCommon.History.Undo();
		["A1", "A2", "A3", "A4"].forEach(function (ref) {
			assert.strictEqual(hasCheckbox(ref), false, ref + " checkbox undone");
		});
		assert.strictEqual(cell("A1").getValue(), "", "A1 blank again after one undo");
		assert.strictEqual(cell("A4").getValue(), "", "A4 blank again after one undo");
		assert.strictEqual(cell("A2").getValue(), "x", "A2 still kept after undo");
		assert.strictEqual(AscCommon.History.Can_Undo(), false, "insert was a single undo point");
	});

	QUnit.test("toggle flips TRUE/FALSE, not for formula cells or read-only", function (assert) {
		reset();
		cell("B2").setValue("=TRUE()");
		select("B1:B2");
		api.asc_insertCheckbox();
		assert.strictEqual(hasCheckbox("B2"), true, "formula cell has checkbox");

		select("B1");
		assert.strictEqual(wsView.toggleActiveCellCheckbox(), true, "Space toggles active cell");
		assert.strictEqual(cell("B1").getValue(), "TRUE", "FALSE -> TRUE");
		assert.strictEqual(wsView.toggleCheckbox(1, 0), true, "click toggles");
		assert.strictEqual(cell("B1").getValue(), "FALSE", "TRUE -> FALSE");

		assert.strictEqual(wsView.toggleCheckbox(1, 1), false, "formula cell not toggled");
		assert.strictEqual(cell("B2").getValue(), "TRUE", "formula value unchanged");

		api.isViewMode = true;
		assert.strictEqual(wsView.toggleCheckbox(1, 0), false, "read-only not toggled");
		assert.strictEqual(cell("B1").getValue(), "FALSE", "read-only value unchanged");
		api.isViewMode = false;

		cell("C1").setValue("TRUE");
		assert.strictEqual(wsView.toggleCheckbox(2, 0), false, "bool cell without checkbox not toggled");
	});

	QUnit.test("remove clears the flag and keeps the value", function (assert) {
		reset();
		select("C1:C2");
		api.asc_insertCheckbox();
		wsView.toggleCheckbox(2, 0);
		assert.strictEqual(cell("C1").getValue(), "TRUE", "C1 checked");

		select("C1:C2");
		api.asc_removeCheckbox();
		assert.strictEqual(hasCheckbox("C1"), false, "C1 flag cleared");
		assert.strictEqual(hasCheckbox("C2"), false, "C2 flag cleared");
		assert.strictEqual(cell("C1").getValue(), "TRUE", "C1 value kept");
		assert.strictEqual(cell("C2").getValue(), "FALSE", "C2 value kept");
		assert.strictEqual(cellInfoCheckbox(), false, "cell info reports no checkbox");
	});

	QUnit.test("checkbox flag survives binary write and read", function (assert) {
		reset();
		select("D1");
		api.asc_insertCheckbox();
		cell("D2").setValue("FALSE");

		var data = new AscCommonExcel.BinaryFileWriter(api.wbModel).Write();
		var wb = new AscCommonExcel.Workbook(undefined, undefined, false);
		wb.DrawingDocument = api.wbModel.DrawingDocument;
		AscFormat.ExecuteNoHistory(function () {
			new AscCommonExcel.BinaryFileReader().Read(data, wb);
		});
		var readWs = wb.aWorksheets[0];
		var getFlag = function (ref) {
			var res = null;
			readWs.getRange2(ref)._foreachNoEmpty(function (c) {
				res = c.getCheckbox();
			});
			return res;
		};
		assert.strictEqual(getFlag("D1"), true, "D1 checkbox read back");
		assert.strictEqual(getFlag("D2"), false, "D2 without checkbox read back");
	});

	QUnit.test("toggle is blocked on a protected sheet for locked cells", function (assert) {
		reset();
		select("A1");
		api.asc_insertCheckbox();
		var protection = new Asc.CSheetProtection();
		protection.setDefaultInterface();
		protection.sheet = true;
		ws.sheetProtection = protection;
		try {
			wsView.toggleCheckbox(0, 0);
			assert.strictEqual(cell("A1").getValue(), "FALSE", "click does not toggle a locked cell");
			cell("A1").setValue("FALSE");
			wsView.toggleActiveCellCheckbox();
			assert.strictEqual(cell("A1").getValue(), "FALSE", "Space does not toggle a locked cell");
		} finally {
			ws.sheetProtection = null;
		}
		wsView.toggleCheckbox(0, 0);
		assert.strictEqual(cell("A1").getValue(), "TRUE", "toggles again once unprotected");
	});

	QUnit.test("cell info reports the flag on non-boolean cells and remove resets it to default", function (assert) {
		reset();
		cell("A2").setValue("x");
		select("A2");
		api.asc_insertCheckbox();
		assert.strictEqual(cellInfoCheckbox(), true, "text cell with checkbox format reported");
		api.asc_insertCheckbox();
		assert.strictEqual(hasCheckbox("A2"), false, "second insert removes the flag");
		var xfs = null;
		ws.getRange2("A2")._foreachNoEmpty(function (c) {
			xfs = c.xfs;
		});
		assert.strictEqual(xfs ? xfs.checkbox : null, null, "removed flag leaves no checkbox property behind");
	});
});
