/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	function cellParagraph(table, row, cell)
	{
		return table.GetRow(row).GetCell(cell).GetContent().GetElement(0);
	}

	function setup()
	{
		AscTest.ClearDocument();
		let table = AscTest.CreateTable(3, 2);
		logicDocument.PushToContent(table);
		for (let r = 0; r < 3; ++r)
			AscTest.AddTextToParagraph(cellParagraph(table, r, 0), "R" + (r + 1));
		AscTest.Recalculate();
		return table;
	}

	function tables()
	{
		let result = [];
		for (let i = 0; i < logicDocument.GetElementsCount(); ++i)
		{
			if (logicDocument.GetElement(i).IsTable())
				result.push(logicDocument.GetElement(i));
		}
		return result;
	}

	QUnit.module("Split table");

	QUnit.test("Splits above the current row, with a paragraph in between", function (assert)
	{
		let table = setup();
		AscTest.MoveCursorToParagraph(cellParagraph(table, 1, 0));

		assert.true(logicDocument.CanSplitTable(), "a table in the main text can be split");
		assert.true(logicDocument.SplitTable(), "split done");

		let t = tables();
		assert.strictEqual(t.length, 2, "two tables");
		assert.strictEqual(t[0].GetRowsCount(), 1, "first table keeps row 1");
		assert.strictEqual(t[1].GetRowsCount(), 2, "second table has rows 2 and 3");
		assert.strictEqual(AscTest.GetParagraphText(cellParagraph(t[1], 0, 0)), "R2", "second table starts with row 2");
		let between = logicDocument.GetElement(t[0].GetIndex() + 1);
		assert.true(between.IsParagraph(), "a paragraph between the tables");
		assert.strictEqual(AscTest.GetParagraphText(between), "", "the paragraph is empty");
	});

	QUnit.test("Undo brings the table back", function (assert)
	{
		let table = setup();
		AscTest.MoveCursorToParagraph(cellParagraph(table, 2, 0));
		logicDocument.SplitTable();
		assert.strictEqual(tables().length, 2, "split");

		logicDocument.Document_Undo();
		let t = tables();
		assert.strictEqual(t.length, 1, "one table again");
		assert.strictEqual(t[0].GetRowsCount(), 3, "with all three rows");
	});

	QUnit.test("In the first row nothing is split, a paragraph is added above", function (assert)
	{
		let table = setup();
		let count = logicDocument.GetElementsCount();
		AscTest.MoveCursorToParagraph(cellParagraph(table, 0, 0));
		logicDocument.SplitTable();

		let t = tables();
		assert.strictEqual(t.length, 1, "still one table");
		assert.strictEqual(t[0].GetRowsCount(), 3, "with all rows");
		assert.strictEqual(logicDocument.GetElementsCount(), count + 1, "one paragraph more");
		assert.true(logicDocument.GetElement(t[0].GetIndex() - 1).IsParagraph(), "the paragraph is above the table");
	});

	QUnit.test("A selection of cells splits above its first row and deletes nothing", function (assert)
	{
		let table = setup();
		AscTest.SelectTableCells(table, 0, 1, 1, 2);
		logicDocument.SplitTable();

		let t = tables();
		assert.strictEqual(t.length, 2, "two tables");
		assert.strictEqual(t[0].GetRowsCount(), 1, "first table keeps row 1");
		assert.strictEqual(AscTest.GetParagraphText(cellParagraph(t[1], 0, 0)), "R2", "text of row 2 kept");
		assert.strictEqual(AscTest.GetParagraphText(cellParagraph(t[1], 1, 0)), "R3", "text of row 3 kept");
	});

	QUnit.test("A nested table is not split", function (assert)
	{
		let table = setup();
		let inner = AscTest.CreateTable(2, 2);
		table.GetRow(1).GetCell(1).GetContent().AddToContent(0, inner);
		AscTest.Recalculate();
		AscTest.MoveCursorToParagraph(cellParagraph(inner, 1, 0));

		assert.false(logicDocument.CanSplitTable(), "cannot split a nested table");
		assert.false(logicDocument.SplitTable(), "nothing done");
		assert.strictEqual(tables().length, 1, "still one table");
		assert.strictEqual(inner.GetRowsCount(), 2, "inner table unchanged");
	});
});
