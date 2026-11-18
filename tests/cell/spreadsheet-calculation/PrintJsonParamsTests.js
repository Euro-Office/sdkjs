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

	QUnit.module("Print empty sheet");

	// Runs drawFunction(wbView) with a stub pdf printer; returns the header and footer draws.
	function drawnHeaderFooters(drawFunction) {
		let wbView = api.wb;
		let wsView = wbView.getWorksheet(0);
		let saved = {
			CPdfPrinter: AscCommonExcel.CPdfPrinter,
			fmgrGraphics: wbView.fmgrGraphics,
			font: wbView.m_oFont,
			stringRender: wsView.stringRender
		};
		let drawn = [];
		// This harness does not lay out the workbook view: give printSheets and drawForPrint what they read.
		AscCommonExcel.CPdfPrinter = function () {
			this.BeginPage = function () {};
			this.EndPage = function () {};
		};
		wbView.fmgrGraphics = saved.fmgrGraphics || [];
		wbView.m_oFont = saved.font || {clone: function () {}};
		wbView._executeWithoutZoom = function (runFunction) {
			runFunction();
		};
		wsView.stringRender = saved.stringRender || {};
		wsView.getZoom = function () {
			return 1;
		};
		wsView._drawHeaderFooter = function (drawingCtx, page, headerFooterData, index, count, bFooter) {
			drawn.push({page: page, index: index, count: count, footer: bFooter});
		};
		try {
			drawFunction(wbView);
		} finally {
			AscCommonExcel.CPdfPrinter = saved.CPdfPrinter;
			wbView.fmgrGraphics = saved.fmgrGraphics;
			wbView.m_oFont = saved.font;
			delete wbView._executeWithoutZoom;
			wsView.stringRender = saved.stringRender;
			delete wsView.getZoom;
			delete wsView._drawHeaderFooter;
		}
		return drawn;
	}

	// Header and footer draws of printSheets.
	function printedHeaderFooters(arrPages, adjustPrint) {
		return drawnHeaderFooters(function (wbView) {
			wbView.printSheets({arrPages: arrPages}, null, adjustPrint);
		});
	}

	// Header and footer draws of the print preview of page index, on a stub preview canvas.
	function previewHeaderFooters(arrPages, index) {
		return drawnHeaderFooters(function (wbView) {
			let previewState = wbView.printPreviewState;
			let saved = {ctx: previewState.getCtx(), pages: previewState.getPages()};
			let fillCtx = {
				fillRect: function () {}
			};
			previewState.setCtx({
				clear: function () {},
				setFillStyle: function () {
					return fillCtx;
				},
				getWidth: function () {
					return 0;
				},
				getHeight: function () {
					return 0;
				}
			});
			previewState.setPages({arrPages: arrPages});
			try {
				wbView.printSheetPrintPreview(index);
			} finally {
				previewState.setCtx(saved.ctx);
				previewState.setPages(saved.pages);
			}
		});
	}

	function withHeaderFooter(runFunction) {
		let ws = api.wbModel.getWorksheet(0);
		let headerFooter = ws.headerFooter;
		let pageSetup = ws.PagePrintOptions.pageSetup;
		let saved = {oddHeader: headerFooter.oddHeader, oddFooter: headerFooter.oddFooter, scale: pageSetup.scale};
		try {
			headerFooter.oddHeader = new Asc.CHeaderFooterData();
			headerFooter.oddHeader.setStr("&CHeader");
			headerFooter.oddFooter = new Asc.CHeaderFooterData();
			headerFooter.oddFooter.setStr("&CPage &P");
			pageSetup.scale = 70;
			runFunction();
		} finally {
			headerFooter.oddHeader = saved.oddHeader;
			headerFooter.oddFooter = saved.oddFooter;
			pageSetup.scale = saved.scale;
		}
	}

	QUnit.test("The blank page of an empty print gets the header and footer", function (assert) {
		withHeaderFooter(function () {
			// No sheet has a print page: printSheets prints one blank page.
			let drawn = printedHeaderFooters([]);

			assert.strictEqual(drawn.length, 2, "header and footer are drawn on the blank page");
			let defaults = AscCommon.c_oAscPrintDefaultSettings;
			for (let i = 0; i < drawn.length; i++) {
				let desc = i === 0 ? "header: " : "footer: ";
				let page = drawn[i].page || {};
				assert.strictEqual(drawn[i].footer, i === 1, desc + "drawn in order");
				assert.strictEqual(page.pageWidth, defaults.PageWidth, desc + "width of the blank page");
				assert.strictEqual(page.pageHeight, defaults.PageHeight, desc + "height of the blank page");
				assert.strictEqual(page.scale, 0.7, desc + "scale of the sheet");
				assert.strictEqual(drawn[i].index, 0, desc + "page 1");
				assert.strictEqual(drawn[i].count, 1, desc + "of 1 page");
			}
		});
	});

	QUnit.test("A page range past the last page prints a blank page without header and footer", function (assert) {
		withHeaderFooter(function () {
			let drawn = printedHeaderFooters([{indexWorksheet: 0}], {startPageIndex: 5, endPageIndex: 6});
			assert.strictEqual(drawn.length, 0, "nothing of the sheet is drawn on that page");
		});
	});

	QUnit.test("The print preview of an empty print shows the header and footer of the printout", function (assert) {
		withHeaderFooter(function () {
			let printed = printedHeaderFooters([]);
			let previewed = previewHeaderFooters([], 0);

			assert.strictEqual(previewed.length, 2, "header and footer are drawn in the preview");
			assert.deepEqual(previewed, printed, "same page size, scale and page number as the printout");
		});
	});

	QUnit.test("The print preview of a page past the last page draws no header and footer", function (assert) {
		withHeaderFooter(function () {
			let drawn = previewHeaderFooters([{indexWorksheet: 0}], 5);
			assert.strictEqual(drawn.length, 0, "nothing of the sheet is drawn on that page");
		});
	});
});
