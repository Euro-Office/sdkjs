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
	AscCommonExcel.WorksheetView.prototype._onUpdateFormatTable = function () {
	};
	AscCommonExcel.WorksheetView.prototype.setSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
	};
	AscCommonExcel.WorksheetView.prototype._reinitializeScroll = function () {
	};
	AscCommonExcel.WorksheetView.prototype.getZoom = function () {
	};
	AscCommonExcel.WorksheetView.prototype._getPPIY = function () {
	};
	AscCommonExcel.WorksheetView.prototype._getPPIX = function () {
	};
	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
	};
	Asc.ReadDefTableStyles = function(){};


	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function() {}
	};
	window["Asc"]["editor"] = api;
	AscCommon.g_oTableId.init();
	api._onEndLoadSdk();
	api.isOpenOOXInBrowser = false;
	api.OpenDocumentFromBin(null, AscCommon.getEmpty());
	api.initCollaborativeEditing({});
	api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
		api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
	var wb = api.wbModel;
	wb.handlers.add("getSelectionState", function () {
		return null;
	});
	wb.handlers.add("getLockDefNameManagerStatus", function () {
		return true;
	});
	api.wb.cellCommentator = new AscCommonExcel.CCellCommentator({
		model: api.wbModel.aWorksheets[0],
		collaborativeEditing: null,
		draw: function() {
		},
		handlers: {
			trigger: function() {
				return false;
			}
		}
	});

	AscCommonExcel.CCellCommentator.prototype.isLockedComment = function (oComment, callbackFunc) {
		callbackFunc(true);
	};
	AscCommonExcel.CCellCommentator.prototype.drawCommentCells = function () {
	};
	AscCommonExcel.CCellCommentator.prototype.ascCvtRatio = function () {
	};

	var wsView = api.wb.getWorksheet(0);
	wsView.handlers = api.handlers;
	wsView.objectRender = new AscFormat.DrawingObjects();
	wsView.objectRender.controller = new AscFormat.DrawingObjectsController(wsView.objectRender);
	var ws = api.wbModel.aWorksheets[0];

	var getRange = function (c1, r1, c2, r2) {
		return new window["Asc"].Range(c1, r1, c2, r2);
	};

	QUnit.test("Test: \"simple tests\"", function (assert) {

		ws.getRange2("A1").setValue("-4");

		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		var base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(0, 1, 0, 1)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("A2").getValue(), ws.getRange2("A1").getValue());

		ws.selectionRange.ranges = [getRange(0, 5, 0, 5), getRange(1, 5, 1, 8)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("A6").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B6").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B7").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B8").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B9").getValue(), "-4");
	});

	QUnit.test("Test: \"formula tests\"", function (assert) {
		var val = "=SIN(1)";

		ws.getRange2("A1").setValue(val);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		var base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(0, 1, 0, 1)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("A2").getValueForEdit(), ws.getRange2("A1").getValueForEdit());

		ws.selectionRange.ranges = [getRange(0, 5, 0, 5), getRange(1, 5, 1, 8)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("A6").getValueForEdit(), val);
		assert.strictEqual(ws.getRange2("B6").getValueForEdit(), val);
		assert.strictEqual(ws.getRange2("B7").getValueForEdit(), val);
		assert.strictEqual(ws.getRange2("B8").getValueForEdit(), val);
		assert.strictEqual(ws.getRange2("B9").getValueForEdit(), val);


		var val1 = "=SIN(A2)";
		var val2 = "=SIN(A3)";

		ws.getRange2("A1").setValue(val1);
		ws.getRange2("A2").setValue(val2);

		ws.selectionRange.ranges = [getRange(0, 0, 0, 1)];
		base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(2, 1, 2, 6), getRange(3, 5, 3, 8), getRange(4, 5, 4, 6)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("C2").getValueForEdit(), "=SIN(C3)");
		assert.strictEqual(ws.getRange2("C3").getValueForEdit(), "=SIN(C4)");
		assert.strictEqual(ws.getRange2("C4").getValueForEdit(), "=SIN(C5)");
		assert.strictEqual(ws.getRange2("C5").getValueForEdit(), "=SIN(C6)");
		assert.strictEqual(ws.getRange2("C6").getValueForEdit(), "=SIN(C7)");
		assert.strictEqual(ws.getRange2("C7").getValueForEdit(), "=SIN(C8)");

		assert.strictEqual(ws.getRange2("D6").getValueForEdit(), "=SIN(D7)");
		assert.strictEqual(ws.getRange2("D7").getValueForEdit(), "=SIN(D8)");
		assert.strictEqual(ws.getRange2("D8").getValueForEdit(), "=SIN(D9)");
		assert.strictEqual(ws.getRange2("D9").getValueForEdit(), "=SIN(D10)");

		assert.strictEqual(ws.getRange2("E6").getValueForEdit(), "=SIN(E7)");
		assert.strictEqual(ws.getRange2("E7").getValueForEdit(), "=SIN(E8)");
	});

	QUnit.test("Test: \"comment tests\"", function (assert) {

		ws.getRange2("E10").setValue("-4");
		var comment = new  window["Asc"].asc_CCommentData(null);
		comment.asc_putText("test");
		comment.bDocument = false;
		/*comment.asc_putTime(this.utcDateToString(new Date()));
		comment.asc_putOnlyOfficeTime(this.ooDateToString(new Date()));
		comment.asc_putUserId(this.currentUserId);
		comment.asc_putUserName(this.currentUserName);
		comment.asc_putSolved(false);*/
		api.asc_addComment(comment);
		comment.nCol = 4;
		comment.nRow = 9;
		comment.coords.nCol = 4;
		comment.coords.nRow = 9;

		ws.selectionRange.ranges = [getRange(4, 9, 4, 9)];

		var base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(0, 1, 0, 1)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("A2").getValue(), ws.getRange2("E10").getValue());
		assert.strictEqual(wsView.cellCommentator.getComment(4,9).nCol, 4);

		ws.selectionRange.ranges = [getRange(0, 5, 0, 5), getRange(1, 5, 1, 8)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("A6").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B6").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B7").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B8").getValue(), "-4");
		assert.strictEqual(ws.getRange2("B9").getValue(), "-4");

		assert.strictEqual(wsView.cellCommentator.getComment(0,5).nRow, 5);
		assert.strictEqual(wsView.cellCommentator.getComment(1,5).nRow, 5);
		assert.strictEqual(wsView.cellCommentator.getComment(1,6).nRow, 6);
		assert.strictEqual(wsView.cellCommentator.getComment(1,7).nRow, 7);
		assert.strictEqual(wsView.cellCommentator.getComment(1,8).nRow, 8);

		assert.strictEqual(wsView.cellCommentator.getComment(1,9), null);
	});

	QUnit.test("Test: \"tables\"", function (assert) {
		ws.autoFilters.addAutoFilter("TableStyleMedium2", getRange(3, 5, 3, 8));

		ws.selectionRange.ranges = [getRange(3, 5, 3, 9)];
		var base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(4, 10, 4, 10)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.TableParts[ws.TableParts.length - 1].Ref.r1, 10);
		assert.strictEqual(ws.TableParts[ws.TableParts.length - 1].Ref.c1, 4);

		ws.selectionRange.ranges = [getRange(5, 10, 5, 10), getRange(6, 10, 6, 10)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.TableParts[ws.TableParts.length - 2].Ref.r1, 10);
		assert.strictEqual(ws.TableParts[ws.TableParts.length - 2].Ref.c1, 5);

		assert.strictEqual(ws.TableParts[ws.TableParts.length - 1].Ref.r1, 10);
		assert.strictEqual(ws.TableParts[ws.TableParts.length - 1].Ref.c1, 6)
	});

	QUnit.test("Test: \"formulas with unar operators\"", function (assert) {
		let originalFormula = "+++1", receivedFormula;

		ws.getRange2("A1").setValue(originalFormula);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		let base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(0, 1, 0, 1)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);
		receivedFormula = ws.getRange2("A2").getValueForEdit();

		assert.strictEqual(originalFormula, receivedFormula, "Copy without formula changing");


		originalFormula = '++++"STR"';
		ws.getRange2("A1").setValue(originalFormula);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(0, 1, 0, 1)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);
		receivedFormula = ws.getRange2("A2").getValueForEdit();

		assert.strictEqual(originalFormula, receivedFormula, "Copy without formula changing");


		originalFormula = '++++FALSE';
		ws.getRange2("A1").setValue(originalFormula);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(0, 1, 0, 1)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);
		receivedFormula = ws.getRange2("A2").getValueForEdit();
		
		assert.strictEqual(originalFormula, receivedFormula, "Copy without formula changing");


		originalFormula = "+SUM(+++1)+++1";
		ws.getRange2("A1").setValue(originalFormula);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(0, 1, 0, 1)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);
		receivedFormula = ws.getRange2("A2").getValueForEdit();

		assert.strictEqual(originalFormula, receivedFormula, "Copy without formula changing");


		originalFormula = "+++-SIN(+-+1-+-+1)+-+1+-+1";
		ws.getRange2("A1").setValue(originalFormula);
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(0, 1, 0, 1)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);
		receivedFormula = ws.getRange2("A2").getValueForEdit();

		assert.strictEqual(originalFormula, receivedFormula, "Copy without formula changing");

		// wsView.cellPasteHelper.loadDataBeforePaste(false, val, true, 0, ws.selectionRange.ranges);

	});

	QUnit.test("Test: \"callback tests paste text\"", function (assert) {
		let done = assert.async();
		api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, "test", undefined, undefined, undefined, function (success) {
			assert.ok(success);
			done();
		});
	});

	QUnit.test("Test: \"callback tests paste HTML\"", function (assert) {
		let done = assert.async();
		let htmlElement = document.createElement("div");
		htmlElement.innerHTML = "test HTML content";
		api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, htmlElement, undefined, undefined, undefined, function (success) {
			assert.ok(success);
			done();
		});
	});

	QUnit.test("Test: \"callback tests paste Binary\"", function (assert) {
		let done = assert.async();
		let binaryData = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);
		api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Internal, binaryData, undefined, undefined, undefined, function (success) {
			assert.ok(success);
			done();
		});
	});

	QUnit.test("Test: paste size is counted from the clipboard data", function (assert) {
		let html = document.createElement("div");
		html.innerHTML = "<table><tr><th>a</th><th>b</th></tr><tr><td>1</td><td>2</td></tr><tr><td>3</td><td>4</td></tr></table>";
		assert.strictEqual(AscCommonExcel.getPasteCellCount(AscCommon.c_oAscClipboardDataFormat.HtmlElement, html), 6, "HTML table cells");

		let paragraph = document.createElement("div");
		paragraph.innerHTML = "<p>a</p><p>b</p><p>c</p>";
		assert.strictEqual(AscCommonExcel.getPasteCellCount(AscCommon.c_oAscClipboardDataFormat.HtmlElement, paragraph), 0, "HTML without a table or text copy");
		assert.strictEqual(AscCommonExcel.getPasteCellCount(AscCommon.c_oAscClipboardDataFormat.HtmlElement, paragraph, "a\nb\nc"), 3, "HTML without a table counted from its text copy");

		assert.strictEqual(AscCommonExcel.getPasteCellCount(AscCommon.c_oAscClipboardDataFormat.Text, "a\tb\tc\n1\t2\t3\n"), 6, "Tab-separated text with a trailing newline");
		assert.strictEqual(AscCommonExcel.getPasteCellCount(AscCommon.c_oAscClipboardDataFormat.Text, "single"), 1, "One cell of text");
		assert.strictEqual(AscCommonExcel.getPasteCellCount(AscCommon.c_oAscClipboardDataFormat.Text, "a\tb\r\n1\t2\r\n"), 4, "Lines ending in CRLF");
		assert.strictEqual(AscCommonExcel.getPasteCellCount(AscCommon.c_oAscClipboardDataFormat.Text, "a\tb\r1\t2\r3"), 5, "Lines ending in a lone CR");
		assert.strictEqual(AscCommonExcel.getPasteCellCount(AscCommon.c_oAscClipboardDataFormat.Text, "a\n\t\t\t\n\t\n"), 7, "Ragged rows counted line by line");
		assert.strictEqual(AscCommonExcel.getPasteCellCount(AscCommon.c_oAscClipboardDataFormat.Text, "\t".repeat(1000000)), 1000001, "One line of many cells");
		assert.strictEqual(AscCommonExcel.getPasteCellCount(AscCommon.c_oAscClipboardDataFormat.Internal, "binary", "a\tb\n1\t2\n"), 4, "Internal paste counted from its text copy");
		assert.strictEqual(AscCommonExcel.getPasteCellCount(AscCommon.c_oAscClipboardDataFormat.Internal, "x".repeat(450), ""), 10, "Internal paste without a text copy estimated from its size");
	});

	QUnit.test("Test: paste size decides between pasting, confirming and refusing", function (assert) {
		let action = AscCommonExcel.c_oAscPasteSizeAction;
		let format = AscCommon.c_oAscClipboardDataFormat;
		let maxChangesSize = 150 * 1024 * 1024;
		let size = function (fmt, cells, max, changes) {
			return AscCommonExcel.getPasteSizeAction(fmt, cells, max, function () {
				return changes;
			});
		};

		assert.strictEqual(size(format.HtmlElement, 50000, maxChangesSize, 0), action.Paste, "HTML up to the confirm threshold");
		assert.strictEqual(size(format.HtmlElement, 50001, maxChangesSize, 0), action.Confirm, "HTML over the confirm threshold");
		assert.strictEqual(size(format.HtmlElement, 150000, maxChangesSize, 0), action.Confirm, "HTML up to the cell limit");
		assert.strictEqual(size(format.HtmlElement, 150001, maxChangesSize, 0), action.Refuse, "HTML over the cell limit");
		assert.strictEqual(size(format.HtmlElement, 150001, 0, 0), action.Refuse, "HTML over the cell limit without maxChangesSize");
		assert.strictEqual(size(format.HtmlElement, 100000, maxChangesSize, maxChangesSize - 1024), action.Refuse, "HTML over maxChangesSize with the session's changes");
		assert.strictEqual(size(format.HtmlElement, 100000, 1024 * 1024, 0), action.Refuse, "HTML over a lower maxChangesSize");

		assert.strictEqual(size(format.Internal, 380000, maxChangesSize, 0), action.Paste, "Internal under the default maxChangesSize, not confirmed");
		assert.strictEqual(size(format.Internal, 390000, maxChangesSize, 0), action.Refuse, "Internal over the default maxChangesSize");
		assert.strictEqual(size(format.Internal, 500000, 0, 0), action.Paste, "Internal up to its cell limit without maxChangesSize, not confirmed");
		assert.strictEqual(size(format.Internal, 500001, 0, 0), action.Refuse, "Internal over its cell limit without maxChangesSize");
		assert.strictEqual(size(format.Internal, 300000, maxChangesSize, 50 * 1024 * 1024), action.Refuse, "Internal over maxChangesSize with the session's changes");

		assert.strictEqual(size(format.Text, 700000, maxChangesSize, 0), action.Paste, "Text under maxChangesSize, not confirmed");
		assert.strictEqual(size(format.Text, 760000, maxChangesSize, 0), action.Refuse, "Text over the default maxChangesSize");
		assert.strictEqual(size(format.Text, 780000, 0, 0), action.Paste, "Text up to the confirm threshold without maxChangesSize");
		assert.strictEqual(size(format.Text, 780001, 0, 0), action.Confirm, "Text over the confirm threshold without maxChangesSize");
		assert.strictEqual(size(format.Text, 2340001, 0, 0), action.Refuse, "Text over the cell limit");

		assert.strictEqual(size(format.Text, 50000, maxChangesSize, maxChangesSize), action.Paste, "Small paste into a full session is left to checkChangesSize");
		assert.strictEqual(size(format.Rtf, 10000000, maxChangesSize, 0), action.Paste, "Formats without limits aren't checked");

		let changesRead = false;
		AscCommonExcel.getPasteSizeAction(format.Text, 50000, maxChangesSize, function () {
			changesRead = true;
			return 0;
		});
		assert.notOk(changesRead, "The session's changes aren't measured for a small paste");
	});

	QUnit.test("Test: paste time is a lower bound in whole seconds", function (assert) {
		let format = AscCommon.c_oAscClipboardDataFormat;
		assert.strictEqual(AscCommonExcel.getPasteMinSeconds(format.HtmlElement, 50001), 2, "HTML at the confirm threshold");
		assert.strictEqual(AscCommonExcel.getPasteMinSeconds(format.HtmlElement, 100000), 5, "HTML, 100,000 cells");
		assert.strictEqual(AscCommonExcel.getPasteMinSeconds(format.HtmlElement, 150000), 8, "HTML at the cell limit");
		assert.strictEqual(AscCommonExcel.getPasteMinSeconds(format.Text, 780001), 3, "Text at the confirm threshold");
		assert.strictEqual(AscCommonExcel.getPasteMinSeconds(format.Text, 1), 1, "Never less than a second");
	});

	QUnit.test("Test: large paste over maxChangesSize is refused", function (assert) {
		let done = assert.async();
		let largeText = "x\n".repeat(100000) + "x";
		let errors = [];
		let onError = function (id) {
			errors.push(id);
		};
		let maxChangesSize = api.maxChangesSize;
		api.maxChangesSize = 1024;
		api.asc_registerCallback("asc_onError", onError);
		api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, largeText, undefined, undefined, undefined, function (success) {
			api.asc_unregisterCallback("asc_onError", onError);
			api.maxChangesSize = maxChangesSize;
			assert.notOk(success, "Paste not done");
			assert.deepEqual(errors, [Asc.c_oAscError.ID.PasteSizeLimitError], "Size limit error sent");
			done();
		});
	});

	QUnit.test("Test: large paste is not done when the confirm is declined", function (assert) {
		let done = assert.async();
		// Text is only confirmed without maxChangesSize
		let largeText = "x\n".repeat(800000) + "x";
		let maxChangesSize = api.maxChangesSize;
		api.maxChangesSize = 0;
		let confirmedCells = null;
		let confirmedSeconds = null;
		let onConfirm = function (cellCount, minSeconds, answer) {
			confirmedCells = cellCount;
			confirmedSeconds = minSeconds;
			answer(false);
		};
		api.asc_registerCallback("asc_onConfirmLargePaste", onConfirm);
		api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, largeText, undefined, undefined, undefined, function (success) {
			api.asc_unregisterCallback("asc_onConfirmLargePaste", onConfirm);
			api.maxChangesSize = maxChangesSize;
			assert.notOk(success, "Paste not done");
			assert.strictEqual(confirmedCells, 800001, "Confirm asked with the cell count");
			assert.strictEqual(confirmedSeconds, 3, "Confirm asked with the minimum time");
			done();
		});
	});

	// Forces the size decision for a small paste, so the confirm paths run without pasting a large range
	function withPasteSizeAction(action, run) {
		let getPasteSizeAction = AscCommonExcel.getPasteSizeAction;
		AscCommonExcel.getPasteSizeAction = function () {
			return action;
		};
		try {
			run();
		} finally {
			AscCommonExcel.getPasteSizeAction = getPasteSizeAction;
		}
	}

	QUnit.test("Test: large paste is done when the confirm is accepted, once", function (assert) {
		let done = assert.async();
		let pastes = 0;
		let onConfirm = function (cellCount, minSeconds, answer) {
			answer(true);
			answer(true);
		};
		api.asc_registerCallback("asc_onConfirmLargePaste", onConfirm);
		withPasteSizeAction(AscCommonExcel.c_oAscPasteSizeAction.Confirm, function () {
			api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, "accepted", undefined, undefined, undefined, function (success) {
				pastes++;
				assert.ok(success, "Paste done");
			});
		});
		api.asc_unregisterCallback("asc_onConfirmLargePaste", onConfirm);
		setTimeout(function () {
			assert.strictEqual(pastes, 1, "A second answer doesn't paste again");
			done();
		}, 0);
	});

	QUnit.test("Test: large paste goes ahead when no UI listens for the confirm", function (assert) {
		// Without the fallback nothing answers, so fail rather than wait
		assert.timeout(5000);
		let done = assert.async();
		withPasteSizeAction(AscCommonExcel.c_oAscPasteSizeAction.Confirm, function () {
			api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, "no listener", undefined, undefined, undefined, function (success) {
				assert.ok(success, "Pasted without asking");
				done();
			});
		});
	});

	QUnit.test("Test: confirmed HTML paste shows its parsing iframe while it runs", function (assert) {
		let done = assert.async();
		let clipboard = AscCommon.g_clipboardBase;
		let commonIframe = clipboard.CommonIframe;
		let frame = {style: {display: "none", width: "100px"}, blur: function () {}};
		clipboard.CommonIframe = frame;
		let html = document.createElement("div");
		html.innerHTML = "<table><tr><td>1</td></tr></table>";
		let onConfirm = function (cellCount, minSeconds, answer) {
			answer(true);
		};
		api.asc_registerCallback("asc_onConfirmLargePaste", onConfirm);
		withPasteSizeAction(AscCommonExcel.c_oAscPasteSizeAction.Confirm, function () {
			api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, html, frame, undefined, undefined, function () {
				assert.strictEqual(frame.style.display, "block", "Shown during the paste");
			});
		});
		api.asc_unregisterCallback("asc_onConfirmLargePaste", onConfirm);
		clipboard.CommonIframe = commonIframe;
		assert.strictEqual(frame.style.display, "none", "Hidden again afterwards");
		assert.strictEqual(frame.style.width, "100px", "Narrowed again afterwards");
		done();
	});

	QUnit.test("Test: internal paste with a cut pending isn't refused", function (assert) {
		let done = assert.async();
		let binaryData = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);
		let cutIdSheet = api.wb.cutIdSheet;
		let errors = [];
		let onError = function (id) {
			errors.push(id);
		};
		api.asc_registerCallback("asc_onError", onError);
		api.wb.cutIdSheet = ws.getId();
		withPasteSizeAction(AscCommonExcel.c_oAscPasteSizeAction.Refuse, function () {
			api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Internal, binaryData, undefined, undefined, undefined, function () {
			});
		});
		api.wb.cutIdSheet = cutIdSheet;
		api.asc_unregisterCallback("asc_onError", onError);
		assert.strictEqual(errors.indexOf(Asc.c_oAscError.ID.PasteSizeLimitError), -1, "Not refused for size");
		done();
	});

	/**
	 * Test suite for parseText CSV parsing functionality
	 */
	QUnit.test("Test: parseText CSV parsing", function (assert) {

		/**
		 * Helper function to create text options for CSV parsing
		 */
		function createTextOptions(delimiter, delimiterChar, textQualifier) {
			var options = new Asc.asc_CTextOptions(AscCommon.c_oAscCodePageUtf8, delimiter, delimiterChar);
			if (textQualifier !== undefined) {
				options.asc_setTextQualifier(textQualifier);
			}
			return options;
		}

		// Test 1: Basic comma-separated values
		var text1 = "a,b,c\nd,e,f";
		var options1 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma);
		var result1 = AscCommon.parseText(text1, options1, true);
		assert.deepEqual(result1, [["a", "b", "c"], ["d", "e", "f"]], "Basic comma separation should work");

		// Test 2: Semicolon delimiter
		var text2 = "x;y;z\n1;2;3";
		var options2 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Semicolon);
		var result2 = AscCommon.parseText(text2, options2, true);
		assert.deepEqual(result2, [["x", "y", "z"], ["1", "2", "3"]], "Semicolon delimiter should work");

		// Test 3: Tab delimiter
		var text3 = "col1\tcol2\tcol3\nval1\tval2\tval3";
		var options3 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Tab);
		var result3 = AscCommon.parseText(text3, options3, true);
		assert.deepEqual(result3, [["col1", "col2", "col3"], ["val1", "val2", "val3"]], "Tab delimiter should work");

		// Test 4: Space delimiter
		var text4 = "one two three\nfour five six";
		var options4 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Space);
		var result4 = AscCommon.parseText(text4, options4, true);
		assert.deepEqual(result4, [["one", "two", "three"], ["four", "five", "six"]], "Space delimiter should work");

		// Test 5: Custom delimiter character
		var text5 = "a|b|c\nd|e|f";
		var options5 = createTextOptions(undefined, "|");
		var result5 = AscCommon.parseText(text5, options5, true);
		assert.deepEqual(result5, [["a", "b", "c"], ["d", "e", "f"]], "Custom delimiter character should work");

		// Test 6: Text qualifiers (quotes) - basic
		var text6 = '"hello","world","test"\n"foo","bar","baz"';
		var options6 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma, undefined, '"');
		var result6 = AscCommon.parseText(text6, options6, true);
		assert.deepEqual(result6, [["hello", "world", "test"], ["foo", "bar", "baz"]], "Basic text qualifiers should work");

		// Test 7: Text qualifiers with embedded delimiters
		var text7 = '"hello, world","test,data",normal\n"a,b,c","x,y,z",simple';
		var options7 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma, undefined, '"');
		var result7 = AscCommon.parseText(text7, options7, true);
		assert.deepEqual(result7, [["hello, world", "test,data", "normal"], ["a,b,c", "x,y,z", "simple"]], "Text qualifiers with embedded delimiters should work");

		// Test 8: Empty fields
		var text8 = "a,,c\n,b,\nd,,f";
		var options8 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma);
		var result8 = AscCommon.parseText(text8, options8, true);
		assert.deepEqual(result8, [["a", "", "c"], ["", "b", ""], ["d", "", "f"]], "Empty fields should be preserved");

		// Test 9: Empty rows
		var text9 = "a,b,c\n\nd,e,f\n\n";
		var options9 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma);
		var result9 = AscCommon.parseText(text9, options9, true);
		assert.deepEqual(result9, [["a", "b", "c"], [""], ["d", "e", "f"], [""]], "Empty rows should be handled correctly");

		// Test 10: Different line endings - Windows (\r\n)
		var text10 = "a,b,c\r\nd,e,f";
		var options10 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma);
		var result10 = AscCommon.parseText(text10, options10, true);
		assert.deepEqual(result10, [["a", "b", "c"], ["d", "e", "f"]], "Windows line endings (\\r\\n) should work");

		// Test 11: Different line endings - Mac Classic (\r)
		var text11 = "a,b,c\rd,e,f";
		var options11 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma);
		var result11 = AscCommon.parseText(text11, options11, true);
		assert.deepEqual(result11, [["a", "b", "c"], ["d", "e", "f"]], "Mac Classic line endings (\\r) should work");

		// Test 12: Mixed content without text qualifiers
		var text12 = "Name,Age,City\nJohn,25,New York\nJane,30,London";
		var options12 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma);
		var result12 = AscCommon.parseText(text12, options12, true);
		assert.deepEqual(result12, [["Name", "Age", "City"], ["John", "25", "New York"], ["Jane", "30", "London"]], "Mixed content without qualifiers should work");

		// Test 13: Single column data
		var text13 = "item1\nitem2\nitem3";
		var options13 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma);
		var result13 = AscCommon.parseText(text13, options13, true);
		assert.deepEqual(result13, [["item1"], ["item2"], ["item3"]], "Single column data should work");

		// Test 14: Single row data
		var text14 = "col1,col2,col3,col4";
		var options14 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma);
		var result14 = AscCommon.parseText(text14, options14, true);
		assert.deepEqual(result14, [["col1", "col2", "col3", "col4"]], "Single row data should work");

		// Test 15: Trailing delimiter
		var text15 = "a,b,c,\nd,e,f,";
		var options15 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma);
		var result15 = AscCommon.parseText(text15, options15, true);
		assert.deepEqual(result15, [["a", "b", "c", ""], ["d", "e", "f", ""]], "Trailing delimiters should create empty fields");

		// Test 16: Leading delimiter
		var text16 = ",a,b,c\n,d,e,f";
		var options16 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma);
		var result16 = AscCommon.parseText(text16, options16, true);
		assert.deepEqual(result16, [["", "a", "b", "c"], ["", "d", "e", "f"]], "Leading delimiters should create empty fields");

		// Test 17: No delimiter (single field per row)
		var text17 = "hello\nworld\ntest";
		var options17 = createTextOptions(AscCommon.c_oAscCsvDelimiter.None);
		var result17 = AscCommon.parseText(text17, options17, true);
		assert.deepEqual(result17, [["hello"], ["world"], ["test"]], "No delimiter should treat each row as single field");

		// Test 18: Quoted empty fields
		var text18 = '"","data",""\n"empty","","filled"';
		var options18 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma, undefined, '"');
		var result18 = AscCommon.parseText(text18, options18, true);
		assert.deepEqual(result18, [["", "data", ""], ["empty", "", "filled"]], "Quoted empty fields should be preserved");

		// Test 19: Mixed quoted and unquoted fields
		var text19 = 'unquoted,"quoted",normal,"with spaces"\nplain,"text",simple,"more text"';
		var options19 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma, undefined, '"');
		var result19 = AscCommon.parseText(text19, options19, true);
		assert.deepEqual(result19, [["unquoted", "quoted", "normal", "with spaces"], ["plain", "text", "simple", "more text"]], "Mixed quoted and unquoted fields should work");

		// Test 20: Text qualifier at end of line
		var text20 = 'a,b,"quoted"\nc,d,"final"';
		var options20 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma, undefined, '"');
		var result20 = AscCommon.parseText(text20, options20, true);
		assert.deepEqual(result20, [["a", "b", "quoted"], ["c", "d", "final"]], "Text qualifiers at end of line should work");

		// Test 21: Whitespace handling with space delimiter
		var text21 = " leading trailing \n  double  spaces  ";
		var options21 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Space);
		var result21 = AscCommon.parseText(text21, options21, true);
		// With bTrimSpaces=true, leading/trailing spaces should be handled specially
		assert.ok(Array.isArray(result21), "Space delimiter with whitespace should return array");
		assert.ok(result21.length > 0, "Should parse at least one row");

		// Test 22: Complex real-world CSV example
		var text22 = 'Name,"Address",Phone,"Email"\n"John Doe","123 Main St, Apt 4",555-1234,"john@example.com"\n"Jane Smith","456 Oak Ave",555-5678,"jane@test.org"';
		var options22 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma, undefined, '"');
		var result22 = AscCommon.parseText(text22, options22, true);
		assert.deepEqual(result22, [
			["Name", "Address", "Phone", "Email"],
			["John Doe", "123 Main St, Apt 4", "555-1234", "john@example.com"],
			["Jane Smith", "456 Oak Ave", "555-5678", "jane@test.org"]
		], "Complex real-world CSV should parse correctly");

		// Test 23: Empty text input
		var text23 = "";
		var options23 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma);
		var result23 = AscCommon.parseText(text23, options23, true);
		assert.deepEqual(result23, [[]], "Empty text should return array with one empty row");

		// Test 24: Only delimiters
		var text24 = ",,,\n,,,";
		var options24 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma);
		var result24 = AscCommon.parseText(text24, options24, true);
		assert.deepEqual(result24, [["", "", "", ""], ["", "", "", ""]], "Only delimiters should create empty fields");

		// Test 25: Escaped quotes inside quoted fields
		var text25 = '"He said ""Hello""","World","Quote: ""test"""\n"A ""quoted"" word","Normal","End ""here"""';
		var options25 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma, undefined, '"');
		var result25 = AscCommon.parseText(text25, options25, true);
		assert.deepEqual(result25, [
			["He said \"Hello\"", "World", "Quote: \"test\""],
			["A \"quoted\" word", "Normal", "End \"here\""]
		], "Escaped quotes inside quoted fields should be handled correctly");

		// Test 26: Quoted field spanning lines (CRLF normalized to LF)
		let text26 = '"A","line1\r\nline2","B"\n"C","D","E"';
		let options26 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma, undefined, '"');
		let result26 = AscCommon.parseText(text26, options26, true);
		assert.deepEqual(result26, [["A", "line1\nline2", "B"], ["C", "D", "E"]], "Quoted newline should not split record, normalize to LF");

		// Test 27: Quoted field spanning lines (CR normalized to LF)
		let text27 = '"A","l1\rl2","B"\n"X","Y","Z"';
		let options27 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma, undefined, '"');
		let result27 = AscCommon.parseText(text27, options27, true);
		assert.deepEqual(result27, [["A", "l1\nl2", "B"], ["X", "Y", "Z"]], "Quoted newline should not split record, normalize to LF");

		// Test 28: Escaped quotes around a newline inside quoted field
		let text28 = '"multi ""line""\nvalue",x,y';
		let options28 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma, undefined, '"');
		let result28 = AscCommon.parseText(text28, options28, true);
		assert.deepEqual(result28, [["multi \"line\"\nvalue", "x", "y"]], "Escaped quotes with embedded newline should parse correctly");

		// Test 29: Custom delimiter with quoted newline
		let text29 = '"x\ny"|z';
		let options29 = createTextOptions(undefined, '|', '"');
		let result29 = AscCommon.parseText(text29, options29, true);
		assert.deepEqual(result29, [["x\ny", "z"]], "Custom delimiter with quoted newline should parse correctly");

		// Test 30: Quoted field across multiple lines (normalized to LF)
		let text30 = '"start\nmiddle\r\nend",foo';
		let options30 = createTextOptions(AscCommon.c_oAscCsvDelimiter.Comma, undefined, '"');
		let result30 = AscCommon.parseText(text30, options30, true);
		assert.deepEqual(result30, [["start\nmiddle\nend", "foo"]], "Quoted field spanning lines normalizes to LF");

		console.log("All parseText CSV tests completed successfully!");
	});

	QUnit.module("CopyPaste");
});
