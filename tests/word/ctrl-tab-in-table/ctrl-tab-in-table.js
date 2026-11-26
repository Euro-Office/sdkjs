/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	window["NATIVE_EDITOR_ENJINE_SYNC_RECALC"] = true;
	let logicDocument = AscTest.CreateLogicDocument();

	function keyEvent(keyCode, ctrl, shift)
	{
		let e = new AscCommon.CKeyboardEvent();
		e.KeyCode  = keyCode;
		e.CtrlKey  = !!ctrl;
		e.ShiftKey = !!shift;
		e.AltKey   = false;
		return e;
	}

	function pressCtrlTab()
	{
		return logicDocument.OnKeyDown(keyEvent(9, true, false));
	}

	// 2 x 2 table; the first cell holds text, the second one is empty
	function createTable(text)
	{
		AscTest.ClearDocument();
		let table = AscTest.CreateTable(2, 2, [50, 50]);
		logicDocument.PushToContent(table);
		if (text)
			AscTest.AddTextToParagraph(table.GetRow(0).GetCell(0).GetContent().GetElement(0), text);
		AscTest.Recalculate();
		return table;
	}

	function cellParagraph(table, row, cell)
	{
		return table.GetRow(row).GetCell(cell).GetContent().GetElement(0);
	}

	// The tab characters in the paragraph and the text around them, e.g. "Hel<TAB>lo"
	function paragraphText(paragraph)
	{
		let text = "";
		paragraph.CheckRunContent(function(run)
		{
			for (let i = 0; i < run.GetElementsCount(); ++i)
			{
				let item = run.GetElement(i);
				if (item.IsTab())
					text += "<TAB>";
				else if (item.IsText())
					text += String.fromCodePoint(item.GetCodePoint());
			}
		});
		return text;
	}

	function putCursorInCell(table, row, cell, pos)
	{
		let paragraph = cellParagraph(table, row, cell);
		table.RemoveSelection();
		table.Document_SetThisElementCurrent();
		table.CurCell = table.GetRow(row).GetCell(cell);
		table.CurCell.GetContent().Set_CurrentElement(0, false);
		paragraph.MoveCursorToStartPos(false);
		for (let i = 0; i < pos; ++i)
			paragraph.MoveCursorRight(false, false);
		return paragraph;
	}

	QUnit.module("Ctrl+Tab in table cells");

	QUnit.test("Ctrl+Tab inside the text of a cell inserts a tab", function(assert)
	{
		let table     = createTable("Hello");
		let paragraph = putCursorInCell(table, 0, 0, 3);
		pressCtrlTab();
		assert.strictEqual(paragraphText(paragraph), "Hel<TAB>lo", "a tab character is inserted at the cursor");
		assert.strictEqual(table.CurCell, table.GetRow(0).GetCell(0), "the cursor stays in the same cell");
	});

	QUnit.test("Ctrl+Tab at the start of a cell inserts a tab (not an indent)", function(assert)
	{
		let table     = createTable("Hello");
		let paragraph = putCursorInCell(table, 0, 0, 0);
		let indBefore = paragraph.Get_CompiledPr2(false).ParaPr.Ind.Left;
		pressCtrlTab();
		assert.strictEqual(paragraphText(paragraph), "<TAB>Hello", "a tab character is inserted");
		assert.strictEqual(paragraph.Get_CompiledPr2(false).ParaPr.Ind.Left, indBefore, "the indent is unchanged");
	});

	QUnit.test("Ctrl+Tab in an empty cell inserts a tab", function(assert)
	{
		let table     = createTable("Hello");
		let paragraph = putCursorInCell(table, 0, 1, 0);
		pressCtrlTab();
		assert.strictEqual(paragraphText(paragraph), "<TAB>", "a tab character is inserted");
		assert.strictEqual(table.CurCell, table.GetRow(0).GetCell(1), "the cursor stays in the same cell");
	});

	QUnit.test("Ctrl+Tab in a numbered paragraph in a cell inserts a tab (not a list level)", function(assert)
	{
		let table     = createTable("Hello");
		let paragraph = putCursorInCell(table, 0, 0, 0);
		let num       = logicDocument.GetNumbering().CreateNum();
		num.CreateDefault(c_oAscMultiLevelNumbering.Numbered);
		paragraph.SetNumPr(num.GetId(), 0);
		AscTest.Recalculate();
		putCursorInCell(table, 0, 0, 0);
		pressCtrlTab();
		assert.strictEqual(paragraphText(paragraph), "<TAB>Hello", "a tab character is inserted");
		assert.strictEqual(paragraph.GetNumPr().Lvl, 0, "the list level is unchanged");
	});

	QUnit.test("Undo removes the tab; Tab alone still moves to the next cell", function(assert)
	{
		let table     = createTable("Hello");
		let paragraph = putCursorInCell(table, 0, 0, 2);
		pressCtrlTab();
		assert.strictEqual(paragraphText(paragraph), "He<TAB>llo", "inserted");
		logicDocument.Document_Undo();
		assert.strictEqual(paragraphText(paragraph), "Hello", "undo removes it");

		putCursorInCell(table, 0, 0, 2);
		logicDocument.OnKeyDown(keyEvent(9, false, false));
		assert.strictEqual(paragraphText(paragraph), "Hello", "Tab inserts nothing");
		assert.strictEqual(table.CurCell, table.GetRow(0).GetCell(1), "Tab moves to the next cell");
	});
});
