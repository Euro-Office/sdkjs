/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	AscCommon.CViewer.prototype.createComponents = function(){};
	AscCommon.CViewer.prototype.resize = function(){};
	// no fonts are needed for pictures (the test page has no font files)
	AscPDF.CPDFDoc.prototype.checkDefaultFonts = function(){};
	AscTest.Editor.sync_SearchEndCallback = function(){};

	const PT_TO_MM = 25.4 / 72;

	// one document with a portrait page (612 x 792 pt) and a landscape page (842 x 595 pt)
	let viewer = new AscCommon.CViewer(null, AscTest.Editor);
	viewer.file = window["AscViewer"].createEmptyFile();
	viewer.file.pages = [
		{W : 612, H : 792, Dpi : 72, Rotate : 0, fonts : []},
		{W : 842, H : 595, Dpi : 72, Rotate : 0, fonts : []}
	];
	viewer.pagesInfo.setCount(2);
	let pdfDoc = viewer.doc;
	AscTest.Editor.DocumentRenderer = viewer;
	AscTest.Editor.getDocumentRenderer = function () { return viewer; };
	AscTest.Editor.getPDFDoc = function () { return pdfDoc; };

	// every test starts with empty pages and an empty history
	function resetDocument()
	{
		for (let page = 0; page < pdfDoc.GetPagesCount(); page++)
			pdfDoc.GetPageInfo(page).drawings.length = 0;
		AscCommon.History.Clear();
		return pdfDoc;
	}

	// what asc_AddWatermarkImage does once the picture is uploaded, with a picture loader that needs no network
	function addWatermark(doc, width, height)
	{
		let api = {
			getPDFDoc : function () { return doc; },
			ImageLoader : {
				LoadImagesWithCallback : function (urls, callback) { callback(); },
				LoadImage : function (url) { return {src : url, Image : {width : width, height : height}}; }
			}
		};
		AscCommon.PDFEditorApi.prototype._addImageUrl.call(api, ["data:image/png;base64,AAAA"], {isWatermark : true});
	}

	function getDrawings(doc, page)
	{
		return doc.GetPageInfo(page).drawings;
	}

	QUnit.module("PDF watermark");

	QUnit.test("asc_AddWatermarkImage is exported", function (assert)
	{
		assert.strictEqual(typeof AscCommon.PDFEditorApi.prototype["asc_AddWatermarkImage"], "function", "the interface can call it");
	});

	QUnit.test("One picture per page, centred, at most 80 % of the page", function (assert)
	{
		let doc = resetDocument();
		addWatermark(doc, 400, 200);

		for (let page = 0; page < 2; page++)
		{
			let drawings = getDrawings(doc, page);
			assert.strictEqual(drawings.length, 1, "page " + (page + 1) + ": one picture");
			assert.ok(drawings[0] instanceof AscPDF.CPdfImage, "page " + (page + 1) + ": it is a picture");

			let pageW = doc.GetPageWidthMM(page);
			let pageH = doc.GetPageHeightMM(page);
			let xfrm = drawings[0].getXfrm();
			assert.close(xfrm.offX + xfrm.extX / 2, pageW / 2, 0.001, "page " + (page + 1) + ": centred horizontally");
			assert.close(xfrm.offY + xfrm.extY / 2, pageH / 2, 0.001, "page " + (page + 1) + ": centred vertically");
			assert.close(xfrm.extX / xfrm.extY, 2, 0.001, "page " + (page + 1) + ": the proportions are kept");
			assert.ok(xfrm.extX <= 0.8 * pageW + 0.001 && xfrm.extY <= 0.8 * pageH + 0.001, "page " + (page + 1) + ": not larger than 80 % of the page");
		}

		// portrait page: the width limits the picture (80 % of 612 pt)
		assert.close(getDrawings(doc, 0)[0].getXfrm().extX, 0.8 * 612 * PT_TO_MM, 0.001, "portrait page: 80 % of the page width");
		// landscape page: the width limits it too (0.8 * 842 / 400 < 0.8 * 595 / 200)
		assert.close(getDrawings(doc, 1)[0].getXfrm().extX, 0.8 * 842 * PT_TO_MM, 0.001, "landscape page: 80 % of the page width");
	});

	QUnit.test("A tall picture is limited by the page height", function (assert)
	{
		let doc = resetDocument();
		addWatermark(doc, 100, 1000);
		let xfrm = getDrawings(doc, 1)[0].getXfrm();
		assert.close(xfrm.extY, 0.8 * 595 * PT_TO_MM, 0.001, "80 % of the page height");
		assert.close(xfrm.offY + xfrm.extY / 2, doc.GetPageHeightMM(1) / 2, 0.001, "centred vertically");
	});

	QUnit.test("All pages are one undo step", function (assert)
	{
		let doc = resetDocument();
		let history = AscCommon.History;
		addWatermark(doc, 400, 200);
		assert.strictEqual(history.Index, 0, "one history point for all pages");
		assert.strictEqual(history.Points[0].Description, AscDFH.historydescription_Presentation_AddFlowImage, "it is an 'add picture' step");

		history.Undo();
		assert.strictEqual(getDrawings(doc, 0).length, 0, "one undo removes the picture on page 1");
		assert.strictEqual(getDrawings(doc, 1).length, 0, "and on page 2");
	});
});
