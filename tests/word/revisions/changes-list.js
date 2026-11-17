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

	function createRun(text, reviewType, userName, dateTime)
	{
		let run = AscTest.CreateRun();
		run.AddText(text);
		if (undefined !== reviewType)
		{
			let info = new AscWord.ReviewInfo();
			info.UserName = userName;
			info.UserId   = userName.toLowerCase();
			info.DateTime = dateTime;
			run.SetReviewTypeWithInfo(reviewType, info);
		}
		return run;
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

	function describe(changes)
	{
		return changes.map(function (c) { return c.get_UserName() + ":" + c.get_Type() + ":" + c.get_Value(); });
	}

	QUnit.module("List of tracked changes");

	QUnit.test("All changes in the order of the document, with author and text", function (assert)
	{
		setup();
		let list = describe(logicDocument.GetRevisionChangesInOrder());
		let add = Asc.c_oAscRevisionsChangeType.TextAdd, rem = Asc.c_oAscRevisionsChangeType.TextRem;
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
		assert.strictEqual(logicDocument.ProcessRevisionChangesOfUser("Anna", true), 2, "two changes of Anna");
		assert.deepEqual(describe(logicDocument.GetRevisionChangesInOrder()), ["Bernd:" + Asc.c_oAscRevisionsChangeType.TextRem + ":Bernd"], "only Bernd's change is left");
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
});
