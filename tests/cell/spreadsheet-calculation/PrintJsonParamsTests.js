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
		let sheetsWithNumbers = function (sheets) {
			return function (index) {
				let number = sheets[index];
				return index < sheets.length ? {PagePrintOptions: {pageSetup: {useFirstPageNumber: number !== null, firstPageNumber: number}}} : null;
			};
		};
		// the printed pages as calcPagesPrint returns them in arrPages: the sheet of each page
		let printJob = function (pages) {
			return pages.map(function (sheet) {
				return {indexWorksheet: sheet};
			});
		};
		// sheets: the first page number of each sheet's page style (null = continue); pages: the sheet of each printed page;
		// jobSheets: the printed sheets of the job as calcPagesPrint records them in CPrintPagesData.jobSheets
		let numbering = function (sheets, pages, jobSheets) {
			wb.getWorksheet = sheetsWithNumbers(sheets);
			try {
				let printPages = printJob(pages);
				return printPages.map(function (page, index) {
					let res = wsView._getOdfPrintPageNumbering(printPages, index, jobSheets);
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
			assert.deepEqual(numbering([null, 5], [0, 1], [0, 1]), [null, null], "XLSX: every sheet starts again (Excel), no ODF numbering");
			wb.sourceFileType = null;
			assert.deepEqual(numbering([null, 5], [0, 1], [0, 1]), [null, null], "unknown source: no ODF numbering");

			wb.sourceFileType = Asc.c_oAscFileType.ODS;
			assert.deepEqual(numbering([null, null, null], [0, 0, 1, 2], [0, 1, 2]), [[0, 1], [1, 1], [2, 1], [3, 1]],
				"ODS: the pages are numbered across the sheets");
			assert.deepEqual(numbering([null, 5, null], [0, 0, 1, 1, 2], [0, 1, 2]), [[0, 1], [1, 1], [0, 5], [1, 5], [2, 5]],
				"ODS: a sheet with a first page number starts again, the sheets after it continue");
			assert.deepEqual(numbering([3, 3], [0, 0, 1], [0, 1]), [[0, 3], [1, 3], [0, 3]],
				"ODS: every sheet with its own first page number starts again (page style names are not kept)");
			assert.deepEqual(numbering([3, 7], [0, 1], [0, 1]), [[0, 3], [0, 7]], "ODS: another first page number starts again");
			assert.deepEqual(numbering([null, 4, null], [0, 2], [0, 1, 2]), [[0, 1], [0, 4]],
				"ODS: a sheet without printed pages still starts the numbers again");
			assert.deepEqual(numbering([null, 0], [0, 1], [0, 1]), [[0, 1], [1, 1]], "ODS: first page number 0 continues");
			assert.deepEqual(numbering([null, 5, null], [1, 2], [1, 2]), [[0, 5], [1, 5]],
				"ODS, the job starts at a sheet with a first page number: the next sheet continues from it");
			assert.deepEqual(numbering([5, null], [1, 1], [0, 1]), [[0, 5], [1, 5]],
				"ODS, whole workbook: a first sheet without printed pages still passes on its number");
			assert.deepEqual(numbering([5, null], [1, 1], [1]), [[0, 1], [1, 1]],
				"ODS, one sheet printed: no number is taken over from a sheet that is not printed");
			assert.deepEqual(numbering([null, 5], [0, 1]), [null, null],
				"ODS, pages without the sheets of the job (not from calcPagesPrint): no ODF numbering");
			assert.deepEqual(numbering([null, 5, null], [0, 0, 2, 2], [0, 2]), [[0, 1], [1, 1], [2, 1], [3, 1]],
				"ODS, sheets 1 and 3 printed: the numbered sheet between them is not in the job and passes nothing on");
			assert.deepEqual(numbering([null, 5, null], [0, 2], [0, 2]), [[0, 1], [1, 1]],
				"ODS, whole workbook with a hidden numbered sheet 2 (calcPagesPrint leaves it out of jobSheets): sheet 3 continues from sheet 1");
			assert.deepEqual(numbering([5, null, null], [2, 2, 0, 0], [2, 0]), [[0, 1], [1, 1], [0, 5], [1, 5]],
				"ODS, selected tabs in the order 3, 1: numbered in print order, sheet 3 takes no number from sheet 1 printed after it");
			wb.sourceFileType = Asc.c_oAscFileType.FODS;
			assert.deepEqual(numbering([null, null], [0, 1], [0, 1]), [[0, 1], [1, 1]], "flat ODS: numbered across the sheets");
			wb.sourceFileType = Asc.c_oAscFileType.OTS;
			assert.deepEqual(numbering([null, null], [0, 1], [0, 1]), [[0, 1], [1, 1]], "ODS template: numbered across the sheets");

			assert.strictEqual(pageNumberText(2, 5), "7", "page number field: index 2 of a run that starts at 5");

			// drawForPrint hands the index and the run's first number to the header/footer (blank page path, no drawing)
			let calls = [];
			let hadStringRender = wsView.hasOwnProperty("stringRender");
			let stringRender = wsView.stringRender;
			wsView.drawHeaderFooter = function (drawingCtx, printPagesData, indexPrintPage, countPrintPages, opt_firstPageNumber) {
				calls.push([indexPrintPage, opt_firstPageNumber]);
			};
			wsView.getZoom = function () {
				return 1;
			};
			if (!wsView.stringRender) {
				wsView.stringRender = {};
			}
			let sheets = [null, 5];
			wb.getWorksheet = function (index) {
				return sheetsWithNumbers(sheets)(index);
			};
			try {
				let printPages = printJob([0, 0, 1, 1]);
				for (let i = 0; i < printPages.length; i++) {
					wsView.drawForPrint({}, null, i, printPages, [0, 1]);
				}
				assert.deepEqual(calls, [[0, 1], [1, 1], [0, 5], [1, 5]], "ODS, drawForPrint: index since the restart and the run's first number");
				calls = [];
				sheets = [null, null];
				for (let i = 0; i < printPages.length; i++) {
					wsView.drawForPrint({}, null, i, printPages, [0, 1]);
				}
				assert.deepEqual(calls, [[0, 1], [1, 1], [2, 1], [3, 1]], "ODS, drawForPrint: the index runs on across the sheets");
				sheets = [null, 5];
				calls = [];
				wb.sourceFileType = Asc.c_oAscFileType.XLSX;
				for (let i = 0; i < printPages.length; i++) {
					wsView.drawForPrint({}, null, i, printPages, [0, 1]);
				}
				assert.deepEqual(calls, [[0, null], [1, null], [0, null], [1, null]], "XLSX, drawForPrint: per sheet index, no run number");
				wb.sourceFileType = Asc.c_oAscFileType.ODS;
				calls = [];
				for (let i = 0; i < printPages.length; i++) {
					wsView.drawForPrint({}, null, i, printPages);
				}
				assert.deepEqual(calls, [[0, null], [1, null], [0, null], [1, null]],
					"ODS, drawForPrint without the sheets of the job: per sheet index, no run number");
			} finally {
				wb.getWorksheet = getWorksheet;
				delete wsView.drawHeaderFooter;
				delete wsView.getZoom;
				if (hadStringRender) {
					wsView.stringRender = stringRender;
				} else {
					delete wsView.stringRender;
				}
			}

			// drawForPrint with a printed page, as printing and the print preview call it: the cells, the drawings and the
			// fonts are stubbed, the page number reaches the header/footer the same way
			calls = [];
			let objectRender = wsView.objectRender;
			let page = function (sheet) {
				return {
					indexWorksheet: sheet, pageRange: new Asc.Range(0, 0, 0, 0), pageWidth: 210, pageHeight: 297, scale: 1,
					pageClipRectLeft: 0, pageClipRectTop: 0, pageClipRectWidth: 100, pageClipRectHeight: 100,
					leftFieldInPx: 0, topFieldInPx: 0, startOffsetPx: 0, titleWidth: 0, titleHeight: 0
				};
			};
			let drawnPages = [];
			wsView.drawHeaderFooter = function (drawingCtx, printPagesData, indexPrintPage, countPrintPages, opt_firstPageNumber) {
				calls.push([indexPrintPage, opt_firstPageNumber]);
			};
			wsView.getZoom = function () {
				return 1;
			};
			wsView.stringRender = {
				resetTransform: function () {
				}
			};
			wsView.objectRender = {
				print: function () {
				},
				updateDrawingsTransform: function () {
				}
			};
			wsView._setDefaultFont = function () {
			};
			wsView._drawCellsAndBorders = function (drawingCtx, range) {
				drawnPages.push(range.r1);
			};
			wsView._RemoveClipRect = function () {
			};
			// no drawing context of the sheet here
			wsView._getPPIX = wsView._getPPIY = function () {
				return 96;
			};
			wb.getWorksheet = sheetsWithNumbers([null, 5]);
			try {
				let printPages = [page(0), page(0), page(1), page(1)];
				for (let i = 0; i < printPages.length; i++) {
					wsView.drawForPrint({DocumentRenderer: {}}, printPages[i], i, printPages, [0, 1]);
				}
				assert.strictEqual(drawnPages.length, 4, "ODS, drawForPrint with printed pages: the cells of every page are drawn (stubbed)");
				assert.deepEqual(calls, [[0, 1], [1, 1], [0, 5], [1, 5]],
					"ODS, drawForPrint with printed pages: index since the restart and the run's first number");
			} finally {
				wb.getWorksheet = getWorksheet;
				delete wsView.drawHeaderFooter;
				delete wsView.getZoom;
				delete wsView._setDefaultFont;
				delete wsView._drawCellsAndBorders;
				delete wsView._RemoveClipRect;
				delete wsView._getPPIX;
				delete wsView._getPPIY;
				wsView.objectRender = objectRender;
				if (hadStringRender) {
					wsView.stringRender = stringRender;
				} else {
					delete wsView.stringRender;
				}
			}
			assert.strictEqual(pageNumberText(2), "3", "page number field: without a start number the sheet's own setting (1)");

			// the real drawHeaderFooter (only the drawing stubbed): the page number text and the first/odd/even choice
			// follow the run, as drawForPrint hands them over
			let drawn = [];
			let getPrintHeaderFooterFromJson = api.wb.getPrintHeaderFooterFromJson;
			let headerFooter = new Asc.CHeaderFooter(null);
			headerFooter.setDifferentFirst(true);
			headerFooter.setDifferentOddEven(true);
			headerFooter.setFirstHeader("F&P");
			headerFooter.setOddHeader("O&P");
			headerFooter.setEvenHeader("E&P");
			headerFooter.setFirstFooter("f&P");
			headerFooter.setOddFooter("o&P");
			headerFooter.setEvenFooter("e&P");
			api.wb.getPrintHeaderFooterFromJson = function () {
				return headerFooter;
			};
			wsView._drawHeaderFooter = function (drawingCtx, printPagesData, curHeader, indexPrintPage, countPrintPages, bFooter) {
				let text = "";
				curHeader.parser.tokens.forEach(function (portion) {
					portion && portion.forEach(function (token) {
						text += [].concat(token._calculatedText).join("");
					});
				});
				drawn.push(text);
			};
			wb.getWorksheet = sheetsWithNumbers([null, 5]);
			try {
				let printPages = printJob([0, 0, 1, 1, 1]);
				let pageTexts = function () {
					let res = [];
					for (let i = 0; i < printPages.length; i++) {
						let odf = wsView._getOdfPrintPageNumbering(printPages, i, [0, 1]);
						drawn = [];
						wsView.drawHeaderFooter({}, {}, odf.index, printPages.length, odf.firstPageNumber);
						res.push(drawn.join(" "));
					}
					return res;
				};
				assert.deepEqual(pageTexts(), ["F1 f1", "E2 e2", "F5 f5", "E6 e6", "O7 o7"],
					"ODS, drawHeaderFooter: page numbers of the run, first page header at the start of each run, then even/odd");
				drawn = [];
				wsView.drawHeaderFooter({}, {}, 2, 5, null);
				assert.deepEqual(drawn, ["O3", "o3"], "drawHeaderFooter without a run number: the sheet's own numbering");
			} finally {
				wb.getWorksheet = getWorksheet;
				api.wb.getPrintHeaderFooterFromJson = getPrintHeaderFooterFromJson;
				delete wsView._drawHeaderFooter;
			}

			// print preview: the page settings of the preview win over the sheet's own
			let printPreviewState = api.wb.printPreviewState;
			let preview = function (started, pageOptionsMap) {
				return {
					isStart: function () {
						return started;
					},
					getActivePageSetup: function () {
						return null;
					},
					advancedOptions: {pageOptionsMap: pageOptionsMap}
				};
			};
			let previewSetup = function (number) {
				return {pageSetup: {useFirstPageNumber: number !== null, firstPageNumber: number}};
			};
			try {
				api.wb.printPreviewState = preview(true, {1: previewSetup(null)});
				assert.deepEqual(numbering([null, 5, null], [0, 1, 2], [0, 1, 2]), [[0, 1], [1, 1], [2, 1]],
					"ODS, print preview: a first page number cleared in the preview continues");
				api.wb.printPreviewState = preview(true, {2: previewSetup(9)});
				assert.deepEqual(numbering([null, 5, null], [0, 1, 2], [0, 1, 2]), [[0, 1], [0, 5], [0, 9]],
					"ODS, print preview: a first page number set in the preview starts again");
				api.wb.printPreviewState = preview(false, {1: previewSetup(null), 2: previewSetup(9)});
				assert.deepEqual(numbering([null, 5, null], [0, 1, 2], [0, 1, 2]), [[0, 1], [0, 5], [1, 5]],
					"ODS, print preview not started: the sheets' own settings");
			} finally {
				api.wb.printPreviewState = printPreviewState;
			}
		} finally {
			wb.getWorksheet = getWorksheet;
			wb.sourceFileType = oldSourceFileType;
		}
	});

	QUnit.test("calcPagesPrint records the sheets of the print job", function (assert) {
		let wbView = api.wb;
		// no layout: sheet 1 has no pages, the others one each; a sheet in hidden is not printed
		let hidden = [];
		wbView.model.getWorksheet = function (index) {
			return {
				PagePrintOptions: null, getHidden: function () {
					return hidden.indexOf(index) !== -1;
				}
			};
		};
		wbView.getWorksheet = function (index) {
			return {
				calcPagesPrint: function (pageOptions, onlySelection, indexWorksheet, arrPages) {
					if (indexWorksheet !== 1) {
						arrPages.push({indexWorksheet: indexWorksheet});
					}
				}
			};
		};
		// no drawing context here: the zoom is read and set around the page layout
		wbView.getZoom = function () {
			return 1;
		};
		wbView.changeZoom = function () {
		};
		wbView.model.getWorksheetCount = function () {
			return 3;
		};
		wbView.model.getActive = function () {
			return 2;
		};
		let jobSheets = function (printType, opt_activeSheets) {
			let adjustPrint = new Asc.asc_CAdjustPrint();
			adjustPrint.asc_setPrintType(printType);
			if (opt_activeSheets) {
				adjustPrint.asc_setActiveSheetsArray(opt_activeSheets);
			}
			let printPagesData = wbView.calcPagesPrint(adjustPrint);
			assert.notOk(printPagesData.arrPages.hasOwnProperty("jobSheets"), "the sheets of the job are not kept on the page array");
			return [printPagesData.jobSheets, printPagesData.arrPages.map(function (page) {
				return page.indexWorksheet;
			})];
		};
		try {
			assert.deepEqual(jobSheets(Asc.c_oAscPrintType.EntireWorkbook), [[0, 1, 2], [0, 2]],
				"whole workbook: every sheet, also the one without pages");
			assert.deepEqual(jobSheets(Asc.c_oAscPrintType.ActiveSheets, [0, 1]), [[0, 1], [0]], "selected sheets");
			assert.deepEqual(jobSheets(Asc.c_oAscPrintType.ActiveSheets, [2, 0]), [[2, 0], [2, 0]],
				"selected sheets keep the order they were selected in");
			assert.deepEqual(jobSheets(Asc.c_oAscPrintType.ActiveSheets), [[2], [2]], "active sheet");
			assert.deepEqual(jobSheets(Asc.c_oAscPrintType.Selection), [[2], [2]], "selection of the active sheet");
			hidden = [1];
			assert.deepEqual(jobSheets(Asc.c_oAscPrintType.EntireWorkbook), [[0, 2], [0, 2]],
				"whole workbook: a hidden sheet is not printed and not one of the sheets of the job");
			hidden = [0];
			assert.deepEqual(jobSheets(Asc.c_oAscPrintType.EntireWorkbook), [[1, 2], [2]],
				"whole workbook: a hidden first sheet is left out, the sheet without pages stays");
		} finally {
			delete wbView.model.getWorksheet;
			delete wbView.getWorksheet;
			delete wbView.getZoom;
			delete wbView.changeZoom;
			delete wbView.model.getWorksheetCount;
			delete wbView.model.getActive;
		}
	});
});
