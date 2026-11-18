/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

// asc_nativePrint with json params (x2t conversions to PDF): the page options it hands to printing.
// Page layout is stubbed out, so these tests need no fonts.
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
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {
	};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {
	};
	AscCommonExcel.WorksheetView.prototype._init = function () {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
	};
	AscCommonExcel.WorksheetView.prototype._reinitializeScroll = function () {
	};
	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
	};
	Asc.ReadDefTableStyles = function () {
	};

	let api = new Asc.spreadsheet_api({
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

	// Runs asc_nativePrint without laying out or drawing pages; returns the page setup it would
	// print the first sheet with.
	function printedPageSetup(options) {
		let calcPagesPrint = api.wb.calcPagesPrint;
		let printSheets = api.wb.printSheets;
		let executeWithoutPreview = api.wb.executeWithoutPreview;
		let adjustPrint = null;
		api.wb.calcPagesPrint = function (adjust) {
			adjustPrint = adjust;
			return {arrPages: []};
		};
		api.wb.printSheets = function () {
			return {DocumentRenderer: {Memory: null}};
		};
		api.wb.executeWithoutPreview = function (runFunction) {
			runFunction();
		};
		try {
			api.asc_nativePrint(undefined, undefined, options);
		} finally {
			api.wb.calcPagesPrint = calcPagesPrint;
			api.wb.printSheets = printSheets;
			api.wb.executeWithoutPreview = executeWithoutPreview;
			api.wb.setPrintOptionsJson(null);
		}
		let pageOptions = adjustPrint && adjustPrint.asc_getPageOptionsMap();
		return pageOptions && pageOptions[0] ? pageOptions[0].asc_getPageSetup() : null;
	}

	QUnit.module("Print json params");

	QUnit.test("A sheet without fit to page keeps its scale", function (assert) {
		let ws = api.wbModel.getWorksheet(0);
		let pageSetup = ws.PagePrintOptions.pageSetup;
		// As read from a file without fitToPage: the fit values keep their default 1.
		ws.setFitToPage(false);
		pageSetup.fitToWidth = 1;
		pageSetup.fitToHeight = 1;
		pageSetup.scale = 80;

		let printed = printedPageSetup({"spreadsheetLayout": {"ignorePrintArea": false}, "adjustOptions": {"printType": 1}});
		assert.ok(printed, "the sheet has page options");
		assert.strictEqual(printed.asc_getFitToWidth(), 0, "not fitted to a page width");
		assert.strictEqual(printed.asc_getFitToHeight(), 0, "not fitted to a page height");
		assert.strictEqual(printed.asc_getScale(), 80, "the sheet's scale");
	});

	QUnit.test("A sheet with fit to page stays fitted", function (assert) {
		let ws = api.wbModel.getWorksheet(0);
		let pageSetup = ws.PagePrintOptions.pageSetup;
		ws.setFitToPage(true);
		pageSetup.fitToWidth = 1;
		pageSetup.fitToHeight = 0;

		let printed = printedPageSetup({"spreadsheetLayout": {"ignorePrintArea": false}, "adjustOptions": {"printType": 1}});
		assert.strictEqual(printed.asc_getFitToWidth(), 1, "fitted to one page width");
		assert.strictEqual(printed.asc_getFitToHeight(), 0, "any number of pages high");
	});

	QUnit.test("Fit values in the json params win", function (assert) {
		let ws = api.wbModel.getWorksheet(0);
		ws.setFitToPage(false);
		ws.PagePrintOptions.pageSetup.fitToWidth = 1;
		ws.PagePrintOptions.pageSetup.fitToHeight = 1;

		let printed = printedPageSetup({"spreadsheetLayout": {"fitToWidth": 2, "fitToHeight": 3}, "adjustOptions": {"printType": 1}});
		assert.strictEqual(printed.asc_getFitToWidth(), 2, "two pages wide");
		assert.strictEqual(printed.asc_getFitToHeight(), 3, "three pages high");

		printed = printedPageSetup({"spreadsheetLayout": {"fitToWidth": 0, "fitToHeight": 0, "scale": 100}, "adjustOptions": {"printType": 1}});
		assert.strictEqual(printed.asc_getFitToWidth(), 0, "scale in the json params: not fitted");
	});
});
