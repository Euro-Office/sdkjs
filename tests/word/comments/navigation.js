/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();
	logicDocument.Comments.Set_Use(true);

	// adding and selecting comments tells the interface about it, the test editor has no interface
	if (!editor.sync_AddComment)
		editor.sync_AddComment = function() {};
	if (!editor.sync_RemoveComment)
		editor.sync_RemoveComment = function() {};

	function addComment(paragraph, start, end, text)
	{
		AscTest.SelectParagraphRange(paragraph, start, end);
		let data = new AscCommon.CCommentData();
		data.SetText(text);
		return logicDocument.AddComment(data, false);
	}

	function getText(id)
	{
		return logicDocument.Comments.Get_ById(id).GetData().GetText();
	}

	function clearDocument()
	{
		AscTest.ClearDocument();
		let comments = logicDocument.Comments.GetAllComments();
		for (let id in comments)
			logicDocument.Comments.Remove_ById(id);
		logicDocument.Comments.Set_Current(null);
	}

	function setupDocument()
	{
		clearDocument();
		let p1 = AscTest.CreateParagraph();
		let p2 = AscTest.CreateParagraph();
		let p3 = AscTest.CreateParagraph();
		logicDocument.PushToContent(p1);
		logicDocument.PushToContent(p2);
		logicDocument.PushToContent(p3);
		AscTest.AddTextToParagraph(p1, "First paragraph");
		AscTest.AddTextToParagraph(p2, "Second paragraph");
		AscTest.AddTextToParagraph(p3, "Third paragraph");

		// added out of order on purpose: the order must follow the document, not the creation time
		addComment(p3, 0, 5, "third");
		addComment(p1, 0, 5, "first");
		addComment(p2, 0, 6, "second");
		logicDocument.RemoveSelection();
		logicDocument.Comments.Set_Current(null);
		return [p1, p2, p3];
	}

	function goTo(isNext)
	{
		let id = logicDocument.GoToNextComment(isNext);
		return id ? getText(id) : null;
	}

	QUnit.module("Comment navigation");

	QUnit.test("Next and previous follow the document order and wrap round", function (assert)
	{
		let paragraphs = setupDocument();
		AscTest.MoveCursorToParagraph(paragraphs[0], true);

		assert.strictEqual(goTo(true), "first", "From the start of the document, next is the first comment");
		assert.strictEqual(goTo(true), "second", "Then the second");
		assert.strictEqual(goTo(true), "third", "Then the third");
		assert.strictEqual(goTo(true), "first", "After the last comment it wraps to the first");
		assert.strictEqual(goTo(false), "third", "Previous from the first wraps to the last");
		assert.strictEqual(goTo(false), "second", "Previous goes back one");
	});

	QUnit.test("Navigation starts from where the cursor is", function (assert)
	{
		let paragraphs = setupDocument();

		AscTest.MoveCursorToParagraph(paragraphs[1], false);
		assert.strictEqual(goTo(true), "third", "From the end of the second paragraph, next is the third comment");

		AscTest.MoveCursorToParagraph(paragraphs[1], false);
		assert.strictEqual(goTo(false), "second", "From the end of the second paragraph, previous is the second comment");
	});

	QUnit.test("Resolved comments are skipped while they are hidden", function (assert)
	{
		let paragraphs = setupDocument();
		let comments = logicDocument.Comments.GetAllComments();
		for (let id in comments)
		{
			if ("second" === comments[id].GetData().GetText())
			{
				let data = comments[id].GetData().Copy();
				data.SetSolved(true);
				comments[id].SetData(data);
			}
		}

		let isUseSolved = logicDocument.Comments.IsUseSolved();
		logicDocument.Comments.SetUseSolved(false);
		AscTest.MoveCursorToParagraph(paragraphs[0], false);
		assert.strictEqual(goTo(true), "third", "The resolved second comment is skipped");

		logicDocument.Comments.SetUseSolved(true);
		AscTest.MoveCursorToParagraph(paragraphs[0], false);
		assert.strictEqual(goTo(true), "second", "It is visited when resolved comments are shown");
		logicDocument.Comments.SetUseSolved(isUseSolved);
	});

	QUnit.test("A document without comments returns nothing", function (assert)
	{
		clearDocument();
		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		AscTest.AddTextToParagraph(paragraph, "No comments");
		logicDocument.Comments.Set_Current(null);
		assert.strictEqual(logicDocument.GetNextCommentId(true), null, "Next finds nothing");
		assert.strictEqual(logicDocument.GetNextCommentId(false), null, "Previous finds nothing");
	});
});
