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
	let commentPanel = {};
	AscTest.Editor.sync_ChangeCommentData = function(id, data)
	{
		changedComments.push(id + "|" + data.m_sUserName);
		commentPanel[id] = [data].concat(data.m_aReplies).map(function(item) { return item.m_sUserName + "|" + (item.m_sUserId || ""); });
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

		// a reply written in the comment panel: it is created from the comment data on every client
		let data = new AscCommon.CCommentData();
		data.Read_FromAscCommentData(highlight.GetAscCommentData());
		let replyData = new AscCommon.CCommentData();
		replyData.m_sText = "Agreed";
		replyData.m_sOOTime = "1700000000000";
		replyData.m_sUserName = "Dave";
		replyData.m_sUserId = "dave-id";
		data.m_aReplies.push(replyData);
		highlight.EditCommentData(data);

		AscCommon.History.Clear();
		changedComments.length = 0;
		commentPanel = {};
		return {note : note, reply : reply, highlight : highlight, panelReply : highlight.GetReply(0)};
	}

	function removePersonalInformation(name)
	{
		AscCommon.PDFEditorApi.prototype.asc_RemovePersonalInformation.call(AscTest.Editor, name);
	}

	function getAuthors(annots)
	{
		return [annots.note, annots.reply, annots.highlight, annots.highlight.GetReply(0)].map(function(annot)
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
		assert.deepEqual(getAuthors(annots), ["Alice|alice-id", "Bob|bob-id", "Carol|carol-id", "Dave|dave-id"], "the annotations have their authors");

		removePersonalInformation("Author");

		assert.deepEqual(getAuthors(annots), ["Author|", "Author|", "Author|", "Author|"], "every annotation and reply has the new author and no user id");
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
		assert.deepEqual(getAuthors(annots), ["Author|", "Author|", "Author|", "Author|"], "the default name is used");
	});

	QUnit.test("Removing personal information is one undo step", function (assert)
	{
		let annots = createDocument();
		let history = AscCommon.History;

		removePersonalInformation("Author");
		assert.strictEqual(history.Index, 0, "one history point for all annotations");
		assert.strictEqual(history.Points[0].Description, AscDFH.historydescription_Document_RemovePersonalInformation, "it is a 'remove personal information' step");

		history.Undo();
		assert.deepEqual(getAuthors(annots), ["Alice|alice-id", "Bob|bob-id", "Carol|carol-id", "Dave|dave-id"], "one undo brings all authors and user ids back");
	});

	QUnit.test("Undo and redo show the replies with the right author in the comment panel", function (assert)
	{
		let annots = createDocument();
		let history = AscCommon.History;
		removePersonalInformation("Author");

		history.Undo();
		assert.deepEqual(commentPanel[annots.note.GetId()], ["Alice|alice-id", "Bob|bob-id"], "after undo the panel shows the old names of the comment and its reply");
		assert.deepEqual(commentPanel[annots.highlight.GetId()], ["Carol|carol-id", "Dave|dave-id"], "and of the highlight comment and its reply");

		history.Redo();
		assert.deepEqual(getAuthors(annots), ["Author|", "Author|", "Author|", "Author|"], "after redo every annotation and reply has the new author again");
		assert.deepEqual(commentPanel[annots.note.GetId()], ["Author|", "Author|"], "the panel shows the new name for the comment and its reply");
		assert.deepEqual(commentPanel[annots.highlight.GetId()], ["Author|", "Author|"], "and for the highlight comment and its reply");
	});

	// The changes of the history point as a co-author (or the server that saves the file) receives them
	function writeChanges(point)
	{
		return point.Items.map(function(item)
		{
			let writer = AscTest.GetBinaryWriter();
			writer.WriteString2(item.Class.Get_Id());
			writer.WriteLong(item.Data.Type);
			item.Data.WriteToBinary(writer);
			return writer;
		});
	}
	function applyChanges(changes)
	{
		AscCommon.History.StartNoHistoryMode();
		changes.forEach(function(writer)
		{
			let reader = AscTest.GetBinaryReader(writer);
			let annot = AscCommon.g_oTableId.Get_ById(reader.GetString2());
			let type = reader.GetLong();
			if (!annot)
				return; // as in CCollaborativeChanges.Apply_Data: an object this copy doesn't have is skipped
			let change = new AscDFH.changesFactory[type](annot);
			change.ReadFromBinary(reader);
			change.Load();
		});
		AscCommon.History.EndNoHistoryMode();
	}

	QUnit.test("Co-authors and the saved file get the new author of replies", function (assert)
	{
		let annots = createDocument();
		let history = AscCommon.History;
		// annotations and replies read from the file are in the id table, so every copy has them with the same id
		[annots.note, annots.reply, annots.highlight].forEach(function(annot)
		{
			if (!AscCommon.g_oTableId.Get_ById(annot.GetId()))
				AscCommon.g_oTableId.Add(annot, annot.GetId());
		});

		removePersonalInformation("Author");
		let changes = writeChanges(history.Points[history.Index]);

		// the other copy: old names, and its own object for the reply written in the comment panel
		history.Undo();
		history.Clear();
		history.StartNoHistoryMode();
		let data = new AscCommon.CCommentData();
		data.Read_FromAscCommentData(annots.highlight.GetAscCommentData());
		data.m_aReplies[0].m_sUserData = "";
		annots.highlight.EditCommentData(data);
		history.EndNoHistoryMode();
		let otherReply = annots.highlight.GetReply(0);
		assert.notStrictEqual(otherReply.GetId(), annots.panelReply.GetId(), "the other copy has its own reply object");
		assert.deepEqual(getAuthors(annots), ["Alice|alice-id", "Bob|bob-id", "Carol|carol-id", "Dave|dave-id"], "the other copy starts with the old names");

		applyChanges(changes);

		assert.deepEqual(getAuthors(annots), ["Author|", "Author|", "Author|", "Author|"], "the other copy gets the new author and no user id for every annotation and reply");
		assert.strictEqual(annots.highlight.GetReply(0).GetContents(), "Agreed", "the reply text is kept");
		assert.strictEqual(annots.highlight.GetReply(0).IsChanged(), true, "the reply is written when the file is saved");
		assert.deepEqual(commentPanel[annots.highlight.GetId()], ["Author|", "Author|"], "the comment panel of the other copy shows the new name for the reply");
	});
});
