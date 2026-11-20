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

	QUnit.module("Print pages of a workbook from an ODS file");

	QUnit.test("Empty pages are left out only for a workbook from an ODS file", function (assert) {
		let ws = api.wbModel.getWorksheet(0);
		let wsView = api.wb.getWorksheet(0);
		let oldSourceFileType = api.wbModel.sourceFileType;
		// Five pages of 50 rows in columns 50-55, away from what the other tests use
		let c1 = 50, c2 = 55, r1 = 1000, rows = 50, count = 5;
		let firstRow = function (index) {
			return r1 + index * rows;
		};
		let pages = function () {
			let arrPages = [];
			for (let i = 0; i < count; ++i) {
				let page = new AscCommonExcel.CPagePrint();
				page.pageRange = new Asc.Range(c1, firstRow(i), c2, firstRow(i) + rows - 1);
				arrPages.push(page);
			}
			return arrPages;
		};
		let index = function (page) {
			return (page.pageRange.r1 - r1) / rows;
		};
		let lastState = null;
		// The pages the print path keeps, the way _calcPagesPrint adds them one by one
		let printedPages = function () {
			let arrPages = [];
			lastState = api.wb.getSkipEmptyPrintPages();
			pages().forEach(function (page) {
				wsView._addPrintPage(arrPages, page, lastState);
			});
			return arrPages.map(index);
		};
		let cell = function (row, col) {
			return ws.getRange3(row, col, row, col);
		};
		let comment = null, sparklineGroup = null;
		try {
			cell(firstRow(4) + 49, c2).setFillColor(new AscCommonExcel.RgbColor(0xFF0000));
			api.wbModel.sourceFileType = Asc.c_oAscFileType.ODS;
			assert.deepEqual(printedPages(), [], "ODS, no content at all: no page is added");
			assert.strictEqual(index(lastState.firstEmptyPage), 0, "ODS, no content at all: the first page is kept for a job without pages");
			assert.strictEqual(lastState.skippedCount, 5, "ODS, no content at all: five pages left out");
			assert.strictEqual(lastState.limitReached, false, "ODS, no content at all: below the page limit");

			let maxPrintPages = AscCommonExcel.c_kMaxPrintPages;
			try {
				AscCommonExcel.c_kMaxPrintPages = 3;
				printedPages();
				assert.strictEqual(lastState.limitReached, true, "ODS: left out pages stop at the page limit");
			} finally {
				AscCommonExcel.c_kMaxPrintPages = maxPrintPages;
			}

			cell(firstRow(0), c1).setValue("first");
			cell(firstRow(4) + 49, c2).setValue("last");

			api.wbModel.sourceFileType = null;
			assert.deepEqual(printedPages(), [0, 1, 2, 3, 4], "unknown source: every page is printed");
			assert.strictEqual(lastState, null, "unknown source: the print path does not leave out pages");
			api.wbModel.sourceFileType = Asc.c_oAscFileType.XLSX;
			assert.deepEqual(printedPages(), [0, 1, 2, 3, 4], "XLSX: every page is printed");

			api.wbModel.sourceFileType = Asc.c_oAscFileType.ODS;
			assert.deepEqual(printedPages(), [0, 4], "ODS: only the pages with values");
			api.wbModel.sourceFileType = Asc.c_oAscFileType.FODS;
			assert.deepEqual(printedPages(), [0, 4], "flat ODS: only the pages with values");
			api.wbModel.sourceFileType = Asc.c_oAscFileType.OTS;
			assert.deepEqual(printedPages(), [0, 4], "ODS template: only the pages with values");
			api.wbModel.sourceFileType = Asc.c_oAscFileType.ODS;

			let diagonal = new AscCommonExcel.Border();
			diagonal.initDefault();
			diagonal.d.setStyle(Asc.c_oAscBorderStyles.Thin);
			diagonal.dd = true;
			cell(firstRow(1) + 5, c1 + 1).setBorderSrc(diagonal);
			assert.deepEqual(printedPages(), [0, 4], "ODS: a diagonal line alone does not count, as in LibreOffice");

			let border = new AscCommonExcel.Border();
			border.initDefault();
			border.b.setStyle(Asc.c_oAscBorderStyles.Thin);
			cell(firstRow(1) + 10, c1 + 2).setBorderSrc(border);
			assert.deepEqual(printedPages(), [0, 1, 4], "ODS: a page with a border is printed");

			cell(firstRow(2) + 20, c1 + 3).setValue("=\"\"");
			assert.deepEqual(printedPages(), [0, 1, 2, 4], "ODS: a page with a formula is printed, also with an empty result");

			cell(firstRow(3) + 30, c1 + 4).setFillColor(new AscCommonExcel.RgbColor(0x00FF00));
			assert.deepEqual(printedPages(), [0, 1, 2, 4], "ODS: a page with a fill only is not printed");

			sparklineGroup = {arrSparklines: [{sqRef: new Asc.Range(c2, firstRow(3) + 40, c2, firstRow(3) + 40)}]};
			ws.aSparklineGroups.push(sparklineGroup);
			assert.deepEqual(printedPages(), [0, 1, 2, 3, 4], "ODS: a page with a sparkline is printed");
			ws.aSparklineGroups.splice(ws.aSparklineGroups.indexOf(sparklineGroup), 1);
			sparklineGroup = null;
			assert.deepEqual(printedPages(), [0, 1, 2, 4], "ODS: without the sparkline the page is left out again");

			comment = new Asc.asc_CCommentData();
			comment.nRow = firstRow(3) + 5;
			comment.nCol = c1;
			ws.aComments.push(comment);
			assert.deepEqual(printedPages(), [0, 1, 2, 3, 4], "ODS: a page with a comment is printed");
		} finally {
			if (sparklineGroup) {
				ws.aSparklineGroups.splice(ws.aSparklineGroups.indexOf(sparklineGroup), 1);
			}
			if (comment) {
				ws.aComments.splice(ws.aComments.indexOf(comment), 1);
			}
			api.wbModel.sourceFileType = oldSourceFileType;
			ws.getRange3(r1, c1, firstRow(count) - 1, c2).cleanAll();
		}
	});

	QUnit.test("Text running in from the left keeps a page and is measured once per layout pass", function (assert) {
		let ws = api.wbModel.getWorksheet(0);
		let wsView = api.wb.getWorksheet(0);
		let oldSourceFileType = api.wbModel.sourceFileType;
		let addCellTextToCache = wsView._addCellTextToCache;
		let getRowHeight = wsView._getRowHeight;
		// One band of 50 rows, the text cell in column 70, pages to its right. The test view cannot measure
		// text (no drawing context), so the measurement is replaced: the text runs 10 columns to the right.
		let col = 70, r1 = 2000, r2 = 2049, row = 2010;
		let page = function (c1, c2) {
			let printPage = new AscCommonExcel.CPagePrint();
			printPage.pageRange = new Asc.Range(c1, r1, c2, r2);
			return printPage;
		};
		let measured = 0;
		try {
			api.wbModel.sourceFileType = Asc.c_oAscFileType.ODS;
			ws.getRange3(row, col, row, col).setValue("text that runs to the right");
			wsView._getRowHeight = function () {
				return 15;
			};
			wsView._addCellTextToCache = function (c, r) {
				if (c === col && r === row) {
					++measured;
				}
				return {flags: {isMerged: function () {
					return false;
				}, wrapText: false}, sideR: 10};
			};
			let arrPages = [];
			let state = api.wb.getSkipEmptyPrintPages();
			[page(71, 73), page(74, 76), page(81, 83), page(1000, 1002)].forEach(function (printPage) {
				wsView._addPrintPage(arrPages, printPage, state);
			});
			assert.deepEqual(arrPages.map(function (printPage) {
				return printPage.pageRange.c1;
			}), [71, 74], "ODS: the pages the text runs into are printed, the pages after its end are not");
			assert.strictEqual(measured, 1, "the text is measured once for all pages to its right");
			assert.strictEqual(wsView._isPrintPageEmpty(new Asc.Range(74, r1, 76, r2)), false, "without a print state the check still works");
		} finally {
			wsView._addCellTextToCache = addCellTextToCache;
			wsView._getRowHeight = getRowHeight;
			api.wbModel.sourceFileType = oldSourceFileType;
			ws.getRange3(r1, col, r2, col).cleanAll();
		}
	});

	QUnit.test("Page numbers of a workbook from an ODS file run across the sheets", function (assert) {
		let wsView = api.wb.getWorksheet(0);
		let wb = api.wbModel;
		let oldSourceFileType = wb.sourceFileType;
		let getWorksheet = wb.getWorksheet;
		// sheets: the first page number of each sheet's page style (null = continue); pages: the sheet of each printed page
		let numbering = function (sheets, pages) {
			wb.getWorksheet = function (index) {
				let number = sheets[index];
				return index < sheets.length ? {PagePrintOptions: {pageSetup: {useFirstPageNumber: number !== null, firstPageNumber: number}}} : null;
			};
			try {
				let printPages = pages.map(function (sheet) {
					return {indexWorksheet: sheet};
				});
				return printPages.map(function (page, index) {
					let res = wsView._getOdfPrintPageNumbering(printPages, index);
					return res && [res.index, res.firstPageNumber];
				});
			} finally {
				wb.getWorksheet = getWorksheet;
			}
		};
		let pageNumberText = function (indexPrintPage, opt_firstPageNumber) {
			let parser = new AscCommonExcel.HeaderFooterParser();
			parser.parse("&P");
			parser.calculateTokens(wsView, indexPrintPage, 10, true, opt_firstPageNumber);
			let text = null;
			parser.tokens.forEach(function (portion) {
				portion && portion.forEach(function (token) {
					if (token && token.getField && token.getField() === Asc.c_oAscHeaderFooterField.pageNumber) {
						text = token._calculatedText;
					}
				});
			});
			return text;
		};
		try {
			wb.sourceFileType = Asc.c_oAscFileType.XLSX;
			assert.deepEqual(numbering([null, 5], [0, 1]), [null, null], "XLSX: every sheet starts again (Excel), no ODF numbering");
			wb.sourceFileType = null;
			assert.deepEqual(numbering([null, 5], [0, 1]), [null, null], "unknown source: no ODF numbering");

			wb.sourceFileType = Asc.c_oAscFileType.ODS;
			assert.deepEqual(numbering([null, null, null], [0, 0, 1, 2]), [[0, 1], [1, 1], [2, 1], [3, 1]],
				"ODS: the pages are numbered across the sheets");
			assert.deepEqual(numbering([null, 5, null], [0, 0, 1, 1, 2]), [[0, 1], [1, 1], [0, 5], [1, 5], [2, 5]],
				"ODS: a sheet with a first page number starts again, the sheets after it continue");
			assert.deepEqual(numbering([3, 3], [0, 0, 1]), [[0, 3], [1, 3], [2, 3]],
				"ODS: the same first page number on the next sheet (same page style) continues");
			assert.deepEqual(numbering([3, 7], [0, 1]), [[0, 3], [0, 7]], "ODS: another first page number starts again");
			assert.deepEqual(numbering([null, 4, null], [0, 2]), [[0, 1], [0, 4]],
				"ODS: a sheet without printed pages still starts the numbers again");
			assert.deepEqual(numbering([null, 0], [0, 1]), [[0, 1], [1, 1]], "ODS: first page number 0 continues");
			wb.sourceFileType = Asc.c_oAscFileType.FODS;
			assert.deepEqual(numbering([null, null], [0, 1]), [[0, 1], [1, 1]], "flat ODS: numbered across the sheets");
			wb.sourceFileType = Asc.c_oAscFileType.OTS;
			assert.deepEqual(numbering([null, null], [0, 1]), [[0, 1], [1, 1]], "ODS template: numbered across the sheets");

			assert.strictEqual(pageNumberText(2, 5), "7", "page number field: index 2 of a run that starts at 5");
			assert.strictEqual(pageNumberText(2), "3", "page number field: without a start number the sheet's own setting (1)");
		} finally {
			wb.getWorksheet = getWorksheet;
			wb.sourceFileType = oldSourceFileType;
		}
	});
});
