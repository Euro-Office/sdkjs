/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	// undo of a comment change tells the interface about it, the test editor has no interface
	if (!editor.sync_ChangeCommentData)
		editor.sync_ChangeCommentData = function() {};

	function createInfo(userName, dateTime)
	{
		let info = new AscWord.ReviewInfo();
		info.UserName = userName;
		info.UserId   = userName.toLowerCase();
		info.DateTime = dateTime;
		return info;
	}

	function createRun(text, reviewType, userName)
	{
		let run = AscTest.CreateRun();
		run.AddText(text);
		if (undefined !== reviewType)
			run.SetReviewTypeWithInfo(reviewType, createInfo(userName, 1000));
		return run;
	}

	function createCommentData(userName, text)
	{
		let data = new AscCommon.CCommentData();
		data.m_sText     = text;
		data.m_sUserName = userName;
		data.m_sUserId   = userName.toLowerCase();
		data.m_sInitials = userName.charAt(0);
		data.m_sTime     = "1700000000000";
		return data;
	}

	function getReviewNames(paragraph)
	{
		let names = [];
		paragraph.CheckRunContent(function(run)
		{
			if (reviewtype_Common !== run.GetReviewType())
				names.push(run.GetReviewInfo().GetUserName() + "|" + run.GetReviewInfo().GetUserId());
			if (run.HavePrChange())
				names.push("pr:" + run.Pr.ReviewInfo.GetUserName());
			return false;
		});
		return names;
	}

	QUnit.module("Remove personal information");

	QUnit.test("Authors of tracked changes, comments and document properties are replaced", function (assert)
	{
		AscTest.ClearDocument();

		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		paragraph.AddToContentToEnd(createRun("added ", reviewtype_Add, "Alice"));
		paragraph.AddToContentToEnd(createRun("removed ", reviewtype_Remove, "Bob"));
		let formatted = createRun("formatted");
		paragraph.AddToContentToEnd(formatted);
		formatted.AddPrChange();
		formatted.Pr.ReviewInfo = createInfo("Carol", 2000);

		let table = AscTest.CreateTable(2, 2);
		logicDocument.PushToContent(table);
		table.GetRow(1).SetReviewTypeWithInfo(reviewtype_Add, createInfo("Dave", 3000));

		let commentData = createCommentData("Erin", "Please check");
		commentData.Add_Reply(createCommentData("Frank", "Done"));
		let comment = new AscCommon.CComment(logicDocument.Comments, commentData);
		logicDocument.Comments.Add(comment);

		if (!logicDocument.Core)
			logicDocument.Core = new AscCommon.CCore();
		logicDocument.Core.creator        = "Grace";
		logicDocument.Core.lastModifiedBy = "Heidi";

		assert.deepEqual(getReviewNames(paragraph), ["Alice|alice", "Bob|bob", "pr:Carol"], "The changes have their authors");

		let changedComments = logicDocument.RemovePersonalInformation("Author");

		assert.deepEqual(getReviewNames(paragraph), ["Author|", "Author|", "pr:Author"], "Inserted, deleted and formatted text has the new author and no id");
		assert.strictEqual(table.GetRow(1).GetReviewInfo().GetUserName(), "Author", "The inserted table row has the new author");
		assert.strictEqual(table.GetRow(1).GetReviewInfo().GetDateTime(), 3000, "The date of the change is kept");

		let data = logicDocument.Comments.Get_ById(comment.GetId()).GetData();
		assert.strictEqual(data.m_sUserName, "Author", "The comment has the new author");
		assert.strictEqual(data.m_sUserId + data.m_sInitials, "", "The comment has no user id and no initials");
		assert.strictEqual(data.m_sText, "Please check", "The comment text is kept");
		assert.strictEqual(data.m_aReplies[0].m_sUserName, "Author", "The reply has the new author");
		assert.strictEqual(data.m_aReplies[0].m_sText, "Done", "The reply text is kept");
		assert.deepEqual(changedComments, [comment.GetId()], "The changed comment is reported");

		assert.strictEqual(logicDocument.Core.creator, null, "The document author is removed");
		assert.strictEqual(logicDocument.Core.lastModifiedBy, null, "Last modified by is removed");
	});

	QUnit.test("Removing personal information can be undone", function (assert)
	{
		AscTest.ClearDocument();

		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		paragraph.AddToContentToEnd(createRun("added", reviewtype_Add, "Alice"));

		AscCommon.History.Clear();
		logicDocument.StartAction(AscDFH.historydescription_Document_RemovePersonalInformation);
		logicDocument.RemovePersonalInformation("Author");
		logicDocument.FinalizeAction();
		assert.deepEqual(getReviewNames(paragraph), ["Author|"], "The author is replaced");

		logicDocument.Document_Undo();
		assert.deepEqual(getReviewNames(paragraph), ["Alice|alice"], "Undo brings the author back");
	});
	QUnit.test("The editor command anonymises a document with comments and tells the interface", function (assert)
	{
		AscTest.ClearDocument();

		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		paragraph.AddToContentToEnd(createRun("added", reviewtype_Add, "Alice"));

		let comment = new AscCommon.CComment(logicDocument.Comments, createCommentData("Erin", "Please check"));
		logicDocument.Comments.Add(comment);

		let changedComments = [];
		let api = {
			private_GetLogicDocument : function() { return logicDocument; },
			sync_ChangeCommentData   : function(id, data) { changedComments.push(id + "|" + data.m_sUserName); },
			UpdateInterfaceState     : function() {}
		};
		Asc.asc_docs_api.prototype.asc_RemovePersonalInformation.call(api);

		assert.deepEqual(getReviewNames(paragraph), ["Author|"], "The tracked change has the default name");
		assert.notStrictEqual(changedComments.indexOf(comment.GetId() + "|Author"), -1, "The interface gets the changed comment");
	});

	QUnit.test("Document properties are removed, dates and the template are kept", function (assert)
	{
		AscTest.ClearDocument();

		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		paragraph.AddToContentToEnd(createRun("added", reviewtype_Add, "Alice"));

		let commentData = createCommentData("Erin", "Please check");
		commentData.m_sOOTime = "1700000000000";
		let comment = new AscCommon.CComment(logicDocument.Comments, commentData);
		logicDocument.Comments.Add(comment);

		if (!logicDocument.Core)
			logicDocument.Core = new AscCommon.CCore();
		logicDocument.Core.title    = "Report for Acme";
		logicDocument.Core.keywords = "secret";
		logicDocument.Core.created  = new Date(1700000000000);

		logicDocument.CustomProperties.AddProperty("Client", AscCommon.c_oVariantTypes.vtLpwstr, "Acme");

		let app = logicDocument.App;
		logicDocument.App = new AscCommon.CApp();
		logicDocument.App.Company  = "Acme Ltd";
		logicDocument.App.Manager  = "Ivan";
		logicDocument.App.Template = "Normal.dotm";

		AscCommon.History.Clear();
		logicDocument.StartAction(AscDFH.historydescription_Document_RemovePersonalInformation);
		logicDocument.RemovePersonalInformation("Author");
		logicDocument.FinalizeAction();

		let reviewInfo = paragraph.GetElement(0).GetReviewInfo();
		assert.strictEqual(reviewInfo.GetUserName() + "|" + reviewInfo.GetDateTime(), "Author|1000", "The tracked change has the new author and keeps its date");

		let data = logicDocument.Comments.Get_ById(comment.GetId()).GetData();
		assert.strictEqual(data.m_sUserName + "|" + data.m_sTime + "|" + data.m_sOOTime, "Author|1700000000000|1700000000000", "The comment has the new author and keeps its date");

		let core = logicDocument.Core;
		assert.deepEqual([core.title, core.keywords, core.created], [null, null, null], "The document properties are empty");
		assert.strictEqual(logicDocument.CustomProperties.getAllProperties().length, 0, "The custom properties are removed");
		assert.deepEqual([logicDocument.App.Company, logicDocument.App.Manager], [null, null], "Company and manager are removed");
		assert.strictEqual(logicDocument.App.Template, "Normal.dotm", "The template is kept, as in Word");

		logicDocument.Document_Undo();
		assert.strictEqual(logicDocument.Core.title, "Report for Acme", "Undo brings the title back");
		assert.strictEqual(logicDocument.CustomProperties.getAllProperties().length, 1, "Undo brings the custom property back");

		logicDocument.App = app;
	});

	QUnit.test("The removed dates can be saved and replayed", function (assert)
	{
		// the desktop app and the server save by replaying the recorded changes, so they must read back the same
		AscTest.ClearDocument();
		logicDocument.PushToContent(AscTest.CreateParagraph());
		if (!logicDocument.Core)
			logicDocument.Core = new AscCommon.CCore();
		logicDocument.Core.title       = "Report";
		logicDocument.Core.created     = new Date(1700000000000);
		logicDocument.Core.lastPrinted = new Date(1700000000000);

		AscCommon.History.Clear();
		logicDocument.StartAction(AscDFH.historydescription_Document_RemovePersonalInformation);
		logicDocument.RemovePersonalInformation("Author");
		logicDocument.FinalizeAction();

		let change = null;
		AscCommon.History.Points[AscCommon.History.Index].Items.forEach(function(item)
		{
			if (AscDFH.historyitem_CoreProperties === item.Data.Type)
				change = item.Data;
		});
		assert.ok(change, "Removing the properties is recorded");

		let writer = AscTest.GetBinaryWriter();
		change.WriteToBinary(writer);
		let replayed = new AscDFH.changesFactory[AscDFH.historyitem_CoreProperties](logicDocument.Core);
		replayed.ReadFromBinary(AscTest.GetBinaryReader(writer));

		assert.strictEqual(replayed.NewCreated, null, "The created date is removed, not an invalid date");
		assert.strictEqual(replayed.NewLastPrinted, null, "The last printed date is removed, not an invalid date");
		assert.strictEqual(replayed.NewTitle, null, "The title is removed");
	});

	QUnit.test("Company and manager are removed by a recorded change that the save replays", function (assert)
	{
		// the server saves by opening the original file and replaying the changes, so a value set only in memory
		// would stay in docProps/app.xml
		AscTest.ClearDocument();
		logicDocument.PushToContent(AscTest.CreateParagraph());

		let app = logicDocument.App;
		logicDocument.App = new AscCommon.CApp();
		logicDocument.App.Company       = "Acme Ltd";
		logicDocument.App.Manager       = "Ivan";
		logicDocument.App.HyperlinkBase = "https://acme.example/";
		logicDocument.App.Template      = "Normal.dotm";

		AscCommon.History.Clear();
		logicDocument.StartAction(AscDFH.historydescription_Document_RemovePersonalInformation);
		logicDocument.RemovePersonalInformation("Author");
		logicDocument.FinalizeAction();

		let change = null;
		AscCommon.History.Points[AscCommon.History.Index].Items.forEach(function(item)
		{
			if (AscDFH.historyitem_Document_AppProps === item.Data.Type)
				change = item.Data;
		});
		assert.ok(change, "Removing company and manager is recorded as a change");
		if (!change)
		{
			logicDocument.App = app;
			return;
		}

		// replay onto the document as it was opened from the original file
		let writer = AscTest.GetBinaryWriter();
		writer.WriteString2(change.Class.Get_Id());
		writer.WriteLong(change.Type);
		change.WriteToBinary(writer);

		logicDocument.App.Company       = "Acme Ltd";
		logicDocument.App.Manager       = "Ivan";
		logicDocument.App.HyperlinkBase = "https://acme.example/";

		let reader   = AscTest.GetBinaryReader(writer);
		let classId  = reader.GetString2();
		let type     = reader.GetLong();
		let replayed = new AscDFH.changesFactory[type](AscCommon.g_oTableId.Get_ById(classId));
		replayed.ReadFromBinary(reader);
		replayed.Redo();

		let props = logicDocument.App;
		assert.deepEqual([props.Company, props.Manager, props.HyperlinkBase], [null, null, null], "The replayed change removes company, manager and hyperlink base");
		assert.strictEqual(props.Template, "Normal.dotm", "The template is kept");

		replayed.CreateReverseChange().Redo();
		assert.deepEqual([props.Company, props.Manager], ["Acme Ltd", "Ivan"], "The reverse change brings them back");

		logicDocument.Document_Undo();
		logicDocument.Document_Redo();
		assert.deepEqual([props.Company, props.Manager], [null, null], "Redo removes them again");
		logicDocument.Document_Undo();
		assert.deepEqual([props.Company, props.Manager, props.HyperlinkBase], ["Acme Ltd", "Ivan", "https://acme.example/"], "Undo brings them back");

		logicDocument.App = app;
	});

	QUnit.test("The mark of the last paragraph of the document and of a table cell gets the new author", function (assert)
	{
		// Word keeps an author on these marks too, e.g. when the paragraph was inserted with tracked changes
		AscTest.ClearDocument();

		let table = AscTest.CreateTable(1, 1);
		logicDocument.PushToContent(table);
		let cellParagraph = table.GetRow(0).GetCell(0).GetContent().GetElement(0);
		cellParagraph.GetParaEndRun().SetReviewTypeWithInfo(reviewtype_Add, createInfo("Alice", 1000), false);

		let lastParagraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(lastParagraph);
		lastParagraph.GetParaEndRun().SetReviewTypeWithInfo(reviewtype_Add, createInfo("Bob", 2000), false);

		assert.strictEqual(lastParagraph.GetParaEndRun().GetReviewInfo().GetUserName(), "Bob", "The last paragraph mark has its author");

		AscCommon.History.Clear();
		logicDocument.StartAction(AscDFH.historydescription_Document_RemovePersonalInformation);
		logicDocument.RemovePersonalInformation("Author");
		logicDocument.FinalizeAction();

		let lastInfo = lastParagraph.GetParaEndRun().GetReviewInfo();
		assert.strictEqual(lastInfo.GetUserName() + "|" + lastInfo.GetUserId() + "|" + lastInfo.GetDateTime(), "Author||2000", "The mark of the last paragraph of the document has the new author and keeps its date");
		let cellInfo = cellParagraph.GetParaEndRun().GetReviewInfo();
		assert.strictEqual(cellInfo.GetUserName() + "|" + cellInfo.GetDateTime(), "Author|1000", "The mark of the last paragraph of the table cell has the new author");

		logicDocument.Document_Undo();
		assert.strictEqual(lastParagraph.GetParaEndRun().GetReviewInfo().GetUserName(), "Bob", "Undo brings the author of the last paragraph mark back");
		assert.strictEqual(cellParagraph.GetParaEndRun().GetReviewInfo().GetUserName(), "Alice", "Undo brings the author of the cell paragraph mark back");
	});

	QUnit.test("The author of moved text is replaced by a recorded change", function (assert)
	{
		// the change is what the save, the co-authors and undo see
		AscTest.ClearDocument();

		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		let moveStart = new AscCommon.CParaRevisionMove(true, true, "move1", createInfo("Alice", 1000));
		paragraph.AddToContent(0, moveStart);
		paragraph.AddToContentToEnd(createRun("moved", reviewtype_Remove, "Alice"));
		let moveEnd = new AscCommon.CParaRevisionMove(false, true, "move1", createInfo("Alice", 1000));
		paragraph.AddToContentToEnd(moveEnd);

		AscCommon.History.Clear();
		logicDocument.StartAction(AscDFH.historydescription_Document_RemovePersonalInformation);
		logicDocument.RemovePersonalInformation("Author");
		logicDocument.FinalizeAction();

		let info = moveStart.GetReviewInfo();
		assert.strictEqual(info.GetUserName() + "|" + info.GetUserId() + "|" + info.GetDateTime(), "Author||1000", "The move mark has the new author and keeps its date");

		let change = null;
		AscCommon.History.Points[AscCommon.History.Index].Items.forEach(function(item)
		{
			if (AscDFH.historyitem_ParaRevisionMove_ReviewInfo === item.Data.Type && moveStart === item.Data.Class)
				change = item.Data;
		});
		assert.ok(change, "The new author of the move mark is recorded as a change");
		if (change)
		{
			// replay onto the mark as it was opened from the original file
			let writer = AscTest.GetBinaryWriter();
			writer.WriteString2(change.Class.Get_Id());
			writer.WriteLong(change.Type);
			change.WriteToBinary(writer);

			moveStart.ReviewInfo = createInfo("Alice", 1000);

			let reader   = AscTest.GetBinaryReader(writer);
			let classId  = reader.GetString2();
			let type     = reader.GetLong();
			let replayed = new AscDFH.changesFactory[type](AscCommon.g_oTableId.Get_ById(classId));
			replayed.ReadFromBinary(reader);
			replayed.Redo();
			assert.strictEqual(moveStart.GetReviewInfo().GetUserName() + "|" + moveStart.GetReviewInfo().GetDateTime(), "Author|1000", "The replayed change gives the move mark the new author");
		}

		logicDocument.Document_Undo();
		assert.strictEqual(moveStart.GetReviewInfo().GetUserName(), "Alice", "Undo brings the author of the move mark back");
		assert.strictEqual(moveEnd.GetReviewInfo().GetUserName(), "Alice", "Undo brings the author of the end mark back");
		logicDocument.Document_Redo();
		assert.strictEqual(moveStart.GetReviewInfo().GetUserName(), "Author", "Redo replaces it again");
	});

	QUnit.test("The saved file has no last modified by after removing personal information", function (assert)
	{
		// the save writes the name of the current user as "last modified by" (docProps/core.xml)
		AscTest.ClearDocument();
		logicDocument.PushToContent(AscTest.CreateParagraph());

		let core = logicDocument.Core;
		logicDocument.Core = new AscCommon.CCore();
		logicDocument.Core.creator        = "Grace";
		logicDocument.Core.lastModifiedBy = "Heidi";

		let api = {DocInfo : {get_UserName : function() { return "Mallory"; }}};
		function getSavedCore(coreToSave)
		{
			let writer = new AscCommon.CBinaryFileWriter();
			coreToSave.toStream(writer, api);
			let saved = new AscCommon.CCore();
			saved.fromStream(new AscCommon.FileStream(writer.data, writer.pos));
			return saved;
		}

		assert.strictEqual(getSavedCore(logicDocument.Core).lastModifiedBy, "Mallory", "Normally the save writes the current user");

		AscCommon.History.Clear();
		logicDocument.StartAction(AscDFH.historydescription_Document_RemovePersonalInformation);
		logicDocument.RemovePersonalInformation("Author");
		logicDocument.FinalizeAction();

		let saved = getSavedCore(logicDocument.Core);
		assert.strictEqual(saved.lastModifiedBy, null, "The saved file has no last modified by");
		assert.strictEqual(saved.creator, null, "The saved file has no author");

		let change = null;
		AscCommon.History.Points[AscCommon.History.Index].Items.forEach(function(item)
		{
			if (AscDFH.historyitem_Core_LastModifiedByRemoved === item.Data.Type)
				change = item.Data;
		});
		assert.ok(change, "Keeping last modified by out of the file is recorded as a change");
		if (change)
		{
			// the server saves by replaying the changes onto the core properties of the original file
			let writer = AscTest.GetBinaryWriter();
			writer.WriteString2(change.Class.Get_Id());
			writer.WriteLong(change.Type);
			change.WriteToBinary(writer);

			let reader   = AscTest.GetBinaryReader(writer);
			let classId  = reader.GetString2();
			let type     = reader.GetLong();
			let original = AscCommon.g_oTableId.Get_ById(classId);
			original.lastModifiedByRemoved = false;
			let replayed = new AscDFH.changesFactory[type](original);
			replayed.ReadFromBinary(reader);
			replayed.Redo();
			assert.strictEqual(getSavedCore(original).lastModifiedBy, null, "After the replayed change the save does not write last modified by");
		}

		logicDocument.Document_Undo();
		assert.strictEqual(getSavedCore(logicDocument.Core).lastModifiedBy, "Mallory", "After undo the save writes the current user again");

		logicDocument.Core = core;
	});
});
