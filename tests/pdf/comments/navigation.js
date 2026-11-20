/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	AscCommon.CViewer.prototype.createComponents = function(){};
	AscCommon.CViewer.prototype.resize = function(){};
	// no fonts are needed (the test page has no font files)
	AscPDF.CPDFDoc.prototype.checkDefaultFonts = function(){};
	AscTest.Editor.sync_SearchEndCallback = function(){};
	// adding an annotation with a comment tells the interface about it, the test editor has no interface
	AscTest.Editor.sync_RemoveComment = function(){};

	// a document with three pages, the viewer shows the first one
	let viewer = new AscCommon.CViewer(null, AscTest.Editor);
	viewer.file = window["AscViewer"].createEmptyFile();
	viewer.file.pages = [
		{W : 612, H : 792, Dpi : 72, Rotate : 0, fonts : []},
		{W : 612, H : 792, Dpi : 72, Rotate : 0, fonts : []},
		{W : 612, H : 792, Dpi : 72, Rotate : 0, fonts : []}
	];
	viewer.pagesInfo.setCount(3);
	let pdfDoc = viewer.doc;
	AscTest.Editor.DocumentRenderer = viewer;
	AscTest.Editor.getDocumentRenderer = function () { return viewer; };
	AscTest.Editor.getPDFDoc = function () { return pdfDoc; };

	// the test page can't scroll: going to an annotation only selects it, as GoToAnnot does
	pdfDoc.GoToAnnot = function (id)
	{
		this.mouseDownAnnot = this.GetAnnotById(id);
	};

	const HIDDEN = window["AscPDF"].Api.Types.display["hidden"];
	const SEPARATOR = AscCommon.UserInfoParser.getSeparator();

	function addAnnot(page, x, y, contents, author, type)
	{
		let annot = AscPDF.CreateAnnotByProps({
			rect     : [x, y, x + 20, y + 20],
			page     : page,
			name     : AscCommon.CreateGUID(),
			type     : undefined === type ? AscPDF.ANNOTATIONS_TYPES.Text : type,
			author   : author || "Reviewer",
			contents : contents,
			hidden   : false
		}, pdfDoc);
		pdfDoc.AddAnnot(annot, page);
		return annot;
	}

	function clearDocument()
	{
		pdfDoc.annots.slice().forEach(function (annot)
		{
			let pageAnnots = pdfDoc.GetPageInfo(annot.GetPage()).annots;
			pageAnnots.splice(pageAnnots.indexOf(annot), 1);
		});
		pdfDoc.mouseDownAnnot = null;
		pdfDoc.annotsHidden = false;
		viewer.currentPage = 0;
	}

	// added out of order on purpose: the order must follow the pages, not the creation time
	function setupDocument()
	{
		clearDocument();
		addAnnot(2, 100, 50, "page 3");
		addAnnot(0, 300, 400, "page 1 bottom");
		addAnnot(1, 100, 100, "page 2 top");
		addAnnot(0, 100, 100, "page 1 top left");
		addAnnot(0, 300, 100, "page 1 top right");
	}

	function goTo(isNext)
	{
		let id = AscCommon.PDFEditorApi.prototype.asc_GotoComment.call(AscTest.Editor, isNext);
		return id ? pdfDoc.GetAnnotById(id).GetContents() : null;
	}

	QUnit.module("PDF comment navigation");

	QUnit.test("asc_GotoComment is exported", function (assert)
	{
		assert.strictEqual(typeof AscCommon.PDFEditorApi.prototype["asc_GotoComment"], "function", "the interface can call it");
	});

	QUnit.test("Page by page, top to bottom, and round in both directions", function (assert)
	{
		setupDocument();

		assert.strictEqual(goTo(true), "page 1 top left", "From the first page, next is the top comment, the left one first");
		assert.strictEqual(goTo(true), "page 1 top right", "Then the one beside it");
		assert.strictEqual(goTo(true), "page 1 bottom", "Then the one further down");
		assert.strictEqual(goTo(true), "page 2 top", "Then the next page");
		assert.strictEqual(goTo(true), "page 3", "Then the last page");
		assert.strictEqual(goTo(true), "page 1 top left", "After the last comment it wraps to the first");
		assert.strictEqual(goTo(false), "page 3", "Previous from the first wraps to the last");
		assert.strictEqual(goTo(false), "page 2 top", "Previous goes back one");
		assert.strictEqual(goTo(false), "page 1 bottom", "And back to the previous page");
	});

	QUnit.test("Without a selection, navigation starts from the current page", function (assert)
	{
		setupDocument();

		viewer.currentPage = 1;
		assert.strictEqual(goTo(true), "page 2 top", "Next is the first comment of the current page");

		pdfDoc.mouseDownAnnot = null;
		assert.strictEqual(goTo(false), "page 1 bottom", "Previous is the last comment before the current page");

		pdfDoc.mouseDownAnnot = null;
		viewer.currentPage = 2;
		pdfDoc.annots.forEach(function (annot)
		{
			if (2 === annot.GetPage())
				annot.SetDisplay(HIDDEN);
		});
		assert.strictEqual(goTo(true), "page 1 top left", "Nothing from the current page on: next wraps to the first comment");
	});

	QUnit.test("A selected annotation without a comment is the starting point", function (assert)
	{
		setupDocument();
		let square = addAnnot(0, 200, 200, null, "Reviewer", AscPDF.ANNOTATIONS_TYPES.Square);
		viewer.currentPage = 2;

		pdfDoc.mouseDownAnnot = square;
		assert.strictEqual(goTo(true), "page 1 bottom", "Next is the comment below the selected shape");

		pdfDoc.mouseDownAnnot = square;
		assert.strictEqual(goTo(false), "page 1 top right", "Previous is the comment above it");
	});

	QUnit.test("Comments the user can't see are skipped", function (assert)
	{
		setupDocument();
		addAnnot(0, 200, 200, "hidden annotation").SetDisplay(HIDDEN);
		addAnnot(0, 200, 250, null);

		assert.strictEqual(goTo(true), "page 1 top left", "First comment");
		assert.strictEqual(goTo(true), "page 1 top right", "Second comment");
		assert.strictEqual(goTo(true), "page 1 bottom", "The hidden annotation and the annotation without a comment are skipped");

		AscCommon.UserInfoParser.setParser(true);
		AscCommon.UserInfoParser.setCommentPermissions({"view" : ["Group A"]});
		clearDocument();
		addAnnot(0, 100, 100, "allowed", "Group A" + SEPARATOR + "Anna");
		addAnnot(0, 100, 200, "not allowed", "Group B" + SEPARATOR + "Ben");
		addAnnot(0, 100, 300, "allowed too", "Group A" + SEPARATOR + "Alex");
		assert.strictEqual(goTo(true), "allowed", "A comment the user may see");
		assert.strictEqual(goTo(true), "allowed too", "The comment of a group the user may not see is skipped");
		AscCommon.UserInfoParser.setCommentPermissions({});
		AscCommon.UserInfoParser.setParser(false);
	});

	QUnit.test("Hidden comments are not visited", function (assert)
	{
		setupDocument();

		pdfDoc.annotsHidden = true;
		assert.strictEqual(goTo(true), null, "Next finds nothing while comments are hidden");
		assert.strictEqual(goTo(false), null, "Previous finds nothing while comments are hidden");

		pdfDoc.annotsHidden = false;
		assert.strictEqual(goTo(true), "page 1 top left", "Shown again, next finds the first comment");
	});

	QUnit.test("A document without comments returns nothing", function (assert)
	{
		clearDocument();
		assert.strictEqual(goTo(true), null, "Next finds nothing");
		assert.strictEqual(goTo(false), null, "Previous finds nothing");
	});
});
