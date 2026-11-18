/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	AscCommon.CViewer.prototype.createComponents = function(){};
	AscCommon.CViewer.prototype.resize = function(){};
	AscPDF.CPDFDoc.prototype.checkDefaultFonts = function(){};
	AscTest.Editor.sync_SearchEndCallback = function(){};

	// the test editor has no interface, the comment panel is recorded here
	["sync_AddComment", "sync_RemoveComment", "sync_HideComment", "sync_UpdateCommentPosition"].forEach(function(name)
	{
		AscTest.Editor[name] = function(){};
	});
	let changedComments = [];
	AscTest.Editor.sync_ChangeCommentData = function(id, data)
	{
		changedComments.push(id + "|" + data.m_sUserName);
	};

	// one document with one page
	let viewer = new AscCommon.CViewer(null, AscTest.Editor);
	viewer.file = window["AscViewer"].createEmptyFile();
	viewer.file.pages = [{W : 612, H : 792, Dpi : 72, Rotate : 0, fonts : []}];
	viewer.pagesInfo.setCount(1);
	viewer.currentPage = 0;
	let pdfDoc = viewer.doc;
	AscTest.Editor.DocumentRenderer = viewer;
	AscTest.Editor.getDocumentRenderer = function () { return viewer; };
	AscTest.Editor.getPDFDoc = function () { return pdfDoc; };

	function createAnnot(type, author, uid, contents, rect)
	{
		return AscPDF.CreateAnnotByProps({
			rect     : rect || [10, 10, 30, 30],
			page     : 0,
			type     : type,
			author   : author,
			contents : contents,
			modDate  : "1700000000000",
			creationDate : "1700000000000"
		}, pdfDoc);
	}

	// a sticky note with a reply, and a highlight with a comment, by different people
	function createDocument()
	{
		AscCommon.History.Clear();
		pdfDoc.annots.slice().forEach(function(annot) { pdfDoc.RemoveAnnot(annot.GetId()); });
		AscCommon.History.Clear();

		let note = pdfDoc.AddAnnotByProps({
			rect : [10, 10, 30, 30], page : 0, type : AscPDF.ANNOTATIONS_TYPES.Text,
			author : "Alice", contents : "Please check", modDate : "1700000000000", creationDate : "1700000000000"
		});
		note.SetUserId("alice-id");
		let reply = createAnnot(AscPDF.ANNOTATIONS_TYPES.Text, "Bob", "bob-id", "Done");
		reply.SetUserId("bob-id");
		note.AddReply(reply);

		let highlight = pdfDoc.AddAnnotByProps({
			rect : [50, 50, 200, 70], page : 0, type : AscPDF.ANNOTATIONS_TYPES.Highlight,
			author : "Carol", contents : "Wrong term", modDate : "1700000000000", creationDate : "1700000000000"
		});
		highlight.SetUserId("carol-id");

		AscCommon.History.Clear();
		changedComments.length = 0;
		return {note : note, reply : reply, highlight : highlight};
	}

	function removePersonalInformation(name)
	{
		AscCommon.PDFEditorApi.prototype.asc_RemovePersonalInformation.call(AscTest.Editor, name);
	}

	function getAuthors(annots)
	{
		return [annots.note, annots.reply, annots.highlight].map(function(annot)
		{
			return annot.GetAuthor() + "|" + (annot.GetUserId() || "");
		});
	}

	QUnit.module("PDF remove personal information");

	QUnit.test("asc_RemovePersonalInformation is exported", function (assert)
	{
		assert.strictEqual(typeof AscCommon.PDFEditorApi.prototype["asc_RemovePersonalInformation"], "function", "the interface can call it");
	});

	QUnit.test("Annotations and replies get the new author, the user ids are removed", function (assert)
	{
		let annots = createDocument();
		assert.deepEqual(getAuthors(annots), ["Alice|alice-id", "Bob|bob-id", "Carol|carol-id"], "the annotations have their authors");

		removePersonalInformation("Author");

		assert.deepEqual(getAuthors(annots), ["Author|", "Author|", "Author|"], "every annotation and reply has the new author and no user id");
		assert.strictEqual(annots.note.GetContents(), "Please check", "the comment text is kept");
		assert.strictEqual(annots.reply.GetContents(), "Done", "the reply text is kept");
		assert.strictEqual(String(annots.note.GetModDate()), "1700000000000", "the date is kept");
		assert.notStrictEqual(changedComments.indexOf(annots.note.GetId() + "|Author"), -1, "the comment panel gets the changed comment");
		assert.notStrictEqual(changedComments.indexOf(annots.highlight.GetId() + "|Author"), -1, "and the changed highlight comment");
	});

	QUnit.test("Without a name the authors become \"Author\"", function (assert)
	{
		let annots = createDocument();
		removePersonalInformation();
		assert.deepEqual(getAuthors(annots), ["Author|", "Author|", "Author|"], "the default name is used");
	});

	QUnit.test("Removing personal information is one undo step", function (assert)
	{
		let annots = createDocument();
		let history = AscCommon.History;

		removePersonalInformation("Author");
		assert.strictEqual(history.Index, 0, "one history point for all annotations");
		assert.strictEqual(history.Points[0].Description, AscDFH.historydescription_Document_RemovePersonalInformation, "it is a 'remove personal information' step");

		history.Undo();
		assert.deepEqual(getAuthors(annots), ["Alice|alice-id", "Bob|bob-id", "Carol|carol-id"], "one undo brings all authors and user ids back");
	});
});
