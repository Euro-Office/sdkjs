/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	if (!editor.sync_ChangeCommentData)
		editor.sync_ChangeCommentData = function() {};

	function createInfo(userName, dateTime, moveType)
	{
		let info = new AscWord.ReviewInfo();
		info.UserName = userName;
		info.UserId   = userName.toLowerCase();
		info.DateTime = dateTime;
		if (undefined !== moveType)
			info.SetMove(moveType);
		return info;
	}

	function createRun(text, reviewType, userName, dateTime, moveType)
	{
		let run = AscTest.CreateRun();
		run.AddText(text);
		if (undefined !== reviewType)
			run.SetReviewTypeWithInfo(reviewType, createInfo(userName, dateTime, moveType));
		return run;
	}

	function createMoveMark(isStart, isFrom)
	{
		// CParaRevisionMove registers itself in the global g_oTableId, which the test environment replaces in
		// AscCommon only
		let tableId = window.g_oTableId;
		window.g_oTableId = AscCommon.g_oTableId;
		try
		{
			return new AscCommon.CParaRevisionMove(isStart, isFrom, "move7", createInfo("Anna", 1000));
		}
		finally
		{
			window.g_oTableId = tableId;
		}
	}

	function addRuns(paragraph, runs)
	{
		runs.forEach(function(run) { paragraph.AddToContentToEnd(run); });
		return paragraph;
	}

	function setup()
	{
		AscTest.ClearDocument();
		let p1 = AscTest.CreateParagraph();
		p1.AddToContentToEnd(createRun("Hallo "));
		p1.AddToContentToEnd(createRun("Anna1", reviewtype_Add, "Anna", 3000));
		p1.AddToContentToEnd(createRun(" Welt"));
		let p2 = AscTest.CreateParagraph();
		p2.AddToContentToEnd(createRun("Text "));
		p2.AddToContentToEnd(createRun("Bernd", reviewtype_Remove, "Bernd", 1000));
		p2.AddToContentToEnd(createRun(" mehr "));
		p2.AddToContentToEnd(createRun("Anna2", reviewtype_Add, "Anna", 2000));
		logicDocument.PushToContent(p1);
		logicDocument.PushToContent(p2);
		AscTest.Recalculate();
		return [p1, p2];
	}

	// Main text with a table, a footnote and a header:
	//   Body1 (Anna, added) | table cell: Cell (Bernd, added) | Gone (Anna, deleted), with a footnote: Note (Anna, added)
	//   header: Head (Bernd, added)
	function setupParts()
	{
		// the footnote and the header find the document only in the document editor
		let isDocumentEditor = editor.isDocumentEditor;
		editor.isDocumentEditor = true;

		AscTest.ClearDocument();
		let p1 = addRuns(AscTest.CreateParagraph(), [createRun("Start "), createRun("Body1", reviewtype_Add, "Anna", 1000)]);
		let table = AscTest.CreateTable(1, 1);
		let cellParagraph = addRuns(table.GetRow(0).GetCell(0).GetContent().GetElement(0), [createRun("Cell", reviewtype_Add, "Bernd", 1000)]);
		let p2 = addRuns(AscTest.CreateParagraph(), [createRun("Body2 "), createRun("Gone", reviewtype_Remove, "Anna", 1000)]);
		logicDocument.PushToContent(p1);
		logicDocument.PushToContent(table);
		logicDocument.PushToContent(p2);

		logicDocument.SetDocPosType(docpostype_Content);
		AscTest.MoveCursorToParagraph(p2, false);
		let footnote = logicDocument.AddFootnote();
		let noteParagraph = addRuns(footnote.GetElement(0), [createRun("Note", reviewtype_Add, "Anna", 1000)]);

		let header = AscTest.CreateDefaultHeader(AscTest.GetFinalSection());
		let headerParagraph = addRuns(header.GetElement(0), [createRun("Head", reviewtype_Add, "Bernd", 1000)]);

		logicDocument.SetDocPosType(docpostype_Content);
		AscTest.MoveCursorToParagraph(p1, true);
		AscTest.Recalculate();
		editor.isDocumentEditor = isDocumentEditor;
		return {p1 : p1, cell : cellParagraph, p2 : p2, note : noteParagraph, header : headerParagraph};
	}

	// "moved" is moved by Anna from the first paragraph to the second one; Bernd added one word
	function setupMove()
	{
		AscTest.ClearDocument();
		let from = Asc.c_oAscRevisionsMove.MoveFrom, to = Asc.c_oAscRevisionsMove.MoveTo;
		let p1 = addRuns(AscTest.CreateParagraph(), [
			createRun("Start "),
			createMoveMark(true, true),
			createRun("moved", reviewtype_Remove, "Anna", 1000, from),
			createMoveMark(false, true),
			createRun(" end")
		]);
		let p2 = addRuns(AscTest.CreateParagraph(), [
			createRun("Target "),
			createMoveMark(true, false),
			createRun("moved", reviewtype_Add, "Anna", 1000, to),
			createMoveMark(false, false),
			createRun(" "),
			createRun("Bernd", reviewtype_Add, "Bernd", 1000)
		]);
		logicDocument.PushToContent(p1);
		logicDocument.PushToContent(p2);
		AscTest.Recalculate();
		return [p1, p2];
	}

	function describe(changes)
	{
		return changes.map(function (c) { return c.get_UserName() + ":" + c.get_Type() + ":" + c.get_Value(); });
	}

	let add = Asc.c_oAscRevisionsChangeType.TextAdd;
	let rem = Asc.c_oAscRevisionsChangeType.TextRem;

	QUnit.module("List of tracked changes");

	QUnit.test("All changes in the order of the document, with author and text", function (assert)
	{
		setup();
		let list = describe(logicDocument.GetRevisionChangesInOrder());
		assert.deepEqual(list, ["Anna:" + add + ":Anna1", "Bernd:" + rem + ":Bernd", "Anna:" + add + ":Anna2"], "document order, not date order");
	});

	QUnit.test("A change from the list is selected", function (assert)
	{
		setup();
		let change = logicDocument.GetRevisionChangesInOrder()[1];
		assert.true(logicDocument.SelectRevisionChange(change), "selected");
		assert.strictEqual(logicDocument.GetSelectedText(), "Bernd", "the changed text is selected");
	});

	QUnit.test("All changes of one author are accepted as one undo step", function (assert)
	{
		let paras = setup();
		let undoPoints = logicDocument.History.Index;
		assert.strictEqual(logicDocument.ProcessRevisionChangesOfUser("Anna", true), 2, "two changes of Anna");
		assert.strictEqual(logicDocument.History.Index, undoPoints + 1, "one undo step");
		assert.deepEqual(describe(logicDocument.GetRevisionChangesInOrder()), ["Bernd:" + rem + ":Bernd"], "only Bernd's change is left");
		assert.strictEqual(AscTest.GetParagraphText(paras[0]), "Hallo Anna1 Welt", "Anna's text stays");

		logicDocument.Document_Undo();
		assert.strictEqual(logicDocument.GetRevisionChangesInOrder().length, 3, "one undo brings all three changes back");
	});

	QUnit.test("All changes of one author are rejected", function (assert)
	{
		let paras = setup();
		assert.strictEqual(logicDocument.ProcessRevisionChangesOfUser("Bernd", false), 1, "one change of Bernd");
		assert.strictEqual(AscTest.GetParagraphText(paras[1]), "Text Bernd mehr Anna2", "the deleted text is back");
		assert.strictEqual(logicDocument.GetRevisionChangesInOrder().length, 2, "Anna's two changes are left");
		assert.strictEqual(logicDocument.ProcessRevisionChangesOfUser("Nobody", true), 0, "an author without changes changes nothing");
	});

	QUnit.test("Changes in a table, a footnote and a header come in the order of \"Next change\"", function (assert)
	{
		setupParts();
		let list = describe(logicDocument.GetRevisionChangesInOrder());
		assert.deepEqual(list, [
			"Anna:" + add + ":Body1",
			"Bernd:" + add + ":Cell",
			"Anna:" + rem + ":Gone",
			"Anna:" + add + ":Note",
			"Bernd:" + add + ":Head"
		], "main text (with the table in place), then footnotes, then headers");

		// "Next change" from the start of the document goes the same way
		let manager = logicDocument.TrackRevisionsManager;
		manager.ClearCurrentChange();
		let walk = [];
		for (let index = 0; index < list.length; ++index)
		{
			let change = manager.GetNextChange();
			if (!change)
				break;
			walk.push(change.get_UserName() + ":" + change.get_Type() + ":" + change.get_Value());
		}
		assert.deepEqual(walk, list, "same order as Next change");
	});

	QUnit.test("Accept and reject by author in a table, a footnote and a header, as one undo step", function (assert)
	{
		let parts = setupParts();
		let undoPoints = logicDocument.History.Index;
		assert.strictEqual(logicDocument.ProcessRevisionChangesOfUser("Anna", true), 3, "three changes of Anna");
		assert.strictEqual(logicDocument.History.Index, undoPoints + 1, "one undo step");
		assert.deepEqual(describe(logicDocument.GetRevisionChangesInOrder()), ["Bernd:" + add + ":Cell", "Bernd:" + add + ":Head"], "only Bernd's changes are left");
		assert.strictEqual(AscTest.GetParagraphText(parts.p2).indexOf("Gone"), -1, "the deleted text is gone");
		assert.true(AscTest.GetParagraphText(parts.note).indexOf("Note") >= 0, "the footnote text stays");

		logicDocument.Document_Undo();
		assert.strictEqual(logicDocument.GetRevisionChangesInOrder().length, 5, "one undo brings all five changes back");

		assert.strictEqual(logicDocument.ProcessRevisionChangesOfUser("Bernd", false), 2, "two changes of Bernd");
		assert.strictEqual(AscTest.GetParagraphText(parts.cell), "", "the added text in the table is gone");
		assert.strictEqual(AscTest.GetParagraphText(parts.header), "", "the added text in the header is gone");
		assert.strictEqual(logicDocument.GetRevisionChangesInOrder().length, 3, "Anna's three changes are left");
	});

	QUnit.test("A moved text is one entry, and accepting it accepts the whole move", function (assert)
	{
		let paras = setupMove();
		let list = logicDocument.GetRevisionChangesInOrder();
		assert.deepEqual(describe(list), ["Anna:" + rem + ":moved", "Bernd:" + add + ":Bernd"], "one entry for the move");
		assert.strictEqual(list[0].GetMoveType(), Asc.c_oAscRevisionsMove.MoveFrom, "the entry is the move");

		logicDocument.AcceptRevisionChange(list[0]);
		assert.strictEqual(AscTest.GetParagraphText(paras[0]), "Start  end", "the text is gone from the old place");
		assert.strictEqual(AscTest.GetParagraphText(paras[1]), "Target moved Bernd", "and stays at the new place");
		assert.deepEqual(describe(logicDocument.GetRevisionChangesInOrder()), ["Bernd:" + add + ":Bernd"], "no half of the move is left");
	});

	QUnit.test("A move is accepted or rejected by author as a whole", function (assert)
	{
		let paras = setupMove();
		assert.strictEqual(logicDocument.ProcessRevisionChangesOfUser("Anna", false), 1, "one move of Anna");
		assert.strictEqual(AscTest.GetParagraphText(paras[0]), "Start moved end", "rejected: the text is back at the old place");
		assert.strictEqual(AscTest.GetParagraphText(paras[1]), "Target  Bernd", "and gone from the new place");
		assert.deepEqual(describe(logicDocument.GetRevisionChangesInOrder()), ["Bernd:" + add + ":Bernd"], "only Bernd's change is left");

		paras = setupMove();
		assert.strictEqual(logicDocument.ProcessRevisionChangesOfUser("Anna", true), 1, "one move of Anna");
		assert.strictEqual(AscTest.GetParagraphText(paras[0]), "Start  end", "accepted: the text is gone from the old place");
		assert.strictEqual(AscTest.GetParagraphText(paras[1]), "Target moved Bernd", "and stays at the new place");
	});

	QUnit.test("Nothing is changed if a paragraph with a change of the author is locked", function (assert)
	{
		let paras = setup();
		paras[1].Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeOther, false);
		try
		{
			let undoPoints = logicDocument.History.Index;
			assert.strictEqual(logicDocument.ProcessRevisionChangesOfUser("Anna", true), -1, "locked");
			assert.strictEqual(logicDocument.GetRevisionChangesInOrder().length, 3, "no change is accepted, also not the one in the first paragraph");
			assert.strictEqual(logicDocument.History.Index, undoPoints, "no undo step");
		}
		finally
		{
			paras[1].Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeNone, false);
		}
	});

	QUnit.test("asc_onRevisionChangesUpdated is sent when the list of changes changes", function (assert)
	{
		let paras = setup();
		let count = 0;
		let onUpdate = function() { ++count; };
		editor.asc_registerCallback("asc_onRevisionChangesUpdated", onUpdate);
		try
		{
			logicDocument.ContinueTrackRevisions();
			count = 0;

			logicDocument.ContinueTrackRevisions();
			assert.strictEqual(count, 0, "nothing changed, no event");

			AscTest.MoveCursorToParagraph(paras[0], true);
			AscTest.EnterText("x");
			logicDocument.ContinueTrackRevisions();
			assert.strictEqual(count, 0, "typing in front of a change: the list is the same, no event");

			logicDocument.ProcessRevisionChangesOfUser("Anna", true);
			logicDocument.ContinueTrackRevisions();
			assert.strictEqual(count, 1, "changes accepted: one event");

			logicDocument.ContinueTrackRevisions();
			assert.strictEqual(count, 1, "and not again");
		}
		finally
		{
			editor.asc_unregisterCallback("asc_onRevisionChangesUpdated", onUpdate);
		}
	});
});
