/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	const SEPARATOR = {type : 3, separator : ";".charCodeAt(0), nested : false};

	function cellParagraph(table, row, cell)
	{
		return table.GetRow(row).GetCell(cell).GetContent().GetElement(0);
	}

	function setup(rows, cols)
	{
		AscTest.ClearDocument();
		let table = AscTest.CreateTable(rows, cols);
		logicDocument.PushToContent(table);
		for (let r = 0; r < rows; ++r)
		{
			for (let c = 0; c < cols; ++c)
				AscTest.AddTextToParagraph(cellParagraph(table, r, c), "R" + (r + 1) + "C" + (c + 1));
		}
		AscTest.Recalculate();
		return table;
	}

	function selectRows(table, startRow, endRow)
	{
		AscTest.SelectTableCells(table, 0, startRow, table.GetRow(endRow).GetCellsCount() - 1, endRow);
	}

	// The document as a short string: T[first cells of the rows] for a table, P(text) for a paragraph
	function describe()
	{
		let out = [];
		for (let i = 0; i < logicDocument.GetElementsCount(); ++i)
		{
			let element = logicDocument.GetElement(i);
			if (element.IsTable())
			{
				let rows = [];
				for (let r = 0; r < element.GetRowsCount(); ++r)
					rows.push(AscTest.GetParagraphText(cellParagraph(element, r, 0)));
				out.push("T[" + rows.join(",") + "]");
			}
			else
			{
				out.push("P(" + AscTest.GetParagraphText(element) + ")");
			}
		}
		return out.join(" ");
	}

	function convert(assert, table, startRow, endRow, props)
	{
		selectRows(table, startRow, endRow);
		try
		{
			logicDocument.ConvertTableToText(props || SEPARATOR);
			AscTest.Recalculate();
			return true;
		}
		catch (e)
		{
			assert.ok(false, "converting throws: " + e);
			return false;
		}
	}

	QUnit.module("Convert table to text");

	QUnit.test("Rows in the middle: the rest stays as one table before and one after", function (assert)
	{
		let table = setup(5, 3);
		if (convert(assert, table, 1, 2))
			assert.strictEqual(describe(), "T[R1C1] P(R2C1;R2C2;R2C3) P(R3C1;R3C2;R3C3) T[R4C1,R5C1] P()", "rows 2-3 converted");
	});

	QUnit.test("One row in the middle", function (assert)
	{
		let table = setup(5, 3);
		if (convert(assert, table, 2, 2))
			assert.strictEqual(describe(), "T[R1C1,R2C1] P(R3C1;R3C2;R3C3) T[R4C1,R5C1] P()", "row 3 converted");
	});

	QUnit.test("The first rows", function (assert)
	{
		let table = setup(5, 3);
		if (convert(assert, table, 0, 1))
			assert.strictEqual(describe(), "P(R1C1;R1C2;R1C3) P(R2C1;R2C2;R2C3) T[R3C1,R4C1,R5C1] P()", "rows 1-2 converted");
	});

	QUnit.test("The last rows", function (assert)
	{
		let table = setup(5, 3);
		if (convert(assert, table, 3, 4))
			assert.strictEqual(describe(), "T[R1C1,R2C1,R3C1] P(R4C1;R4C2;R4C3) P(R5C1;R5C2;R5C3) P()", "rows 4-5 converted");
	});

	QUnit.test("All rows still convert the whole table", function (assert)
	{
		let table = setup(3, 2);
		if (convert(assert, table, 0, 2))
			assert.strictEqual(describe(), "P(R1C1;R1C2) P(R2C1;R2C2) P(R3C1;R3C2) P()", "whole table converted");
	});

	QUnit.test("Tabs and paragraph marks as separators", function (assert)
	{
		let table = setup(4, 2);
		if (convert(assert, table, 1, 1, {type : 2, nested : false}))
			assert.strictEqual(describe(), "T[R1C1] P(R2C1 R2C2) T[R3C1,R4C1] P()", "tabs (a tab is shown as a space)");

		table = setup(4, 2);
		if (convert(assert, table, 1, 1, {type : 1, nested : false}))
			assert.strictEqual(describe(), "T[R1C1] P(R2C1) P(R2C2) T[R3C1,R4C1] P()", "paragraph marks");
	});

	QUnit.test("A cell merged across the end of the selection is merged again in the table after it", function (assert)
	{
		let table = setup(5, 3);
		AscTest.MergeTableCells(table, 1, 2, 1, 4); // column 2, rows 3-5
		AscTest.Recalculate();
		if (!convert(assert, table, 1, 2))
			return;

		// the merged cell starts in row 3, so its text (three paragraphs after merging) goes into the text
		assert.strictEqual(describe(), "T[R1C1] P(R2C1;R2C2;R2C3) P(R3C1;R3C2) P(R4C2) P(R5C2;R3C3) T[R4C1,R5C1] P()", "rows 2-3 converted");
		let after = logicDocument.GetElement(logicDocument.GetElementsCount() - 2);
		assert.strictEqual(after.GetRow(0).GetCell(1).GetVMerge(), vmerge_Restart, "the merge starts again in row 4");
		assert.strictEqual(after.GetRow(1).GetCell(1).GetVMerge(), vmerge_Continue, "row 5 is still merged with row 4");
	});

	QUnit.test("A nested table in the first cell is not changed", function (assert)
	{
		let table = setup(5, 2);
		let inner = AscTest.CreateTable(3, 2);
		table.GetRow(0).GetCell(0).GetContent().AddToContent(0, inner);
		AscTest.Recalculate();
		if (!convert(assert, table, 1, 2))
			return;

		let before = logicDocument.GetElement(0);
		assert.true(before.IsTable(), "a table before the text");
		assert.strictEqual(before.GetRowsCount(), 1, "with row 1");
		let nested = before.GetRow(0).GetCell(0).GetContent().GetElement(0);
		assert.true(nested.IsTable(), "the nested table is still there");
		assert.strictEqual(nested.GetRowsCount(), 3, "with all its rows");
	});

	QUnit.test("Undo brings the table back", function (assert)
	{
		let table = setup(5, 2);
		if (!convert(assert, table, 1, 2))
			return;

		logicDocument.Document_Undo();
		AscTest.Recalculate();
		assert.strictEqual(describe(), "T[R1C1,R2C1,R3C1,R4C1,R5C1] P()", "one table with five rows again");
	});
});
