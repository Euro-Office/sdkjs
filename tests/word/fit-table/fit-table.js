/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	function setup(ind)
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(210, 297);
		sectPr.SetPageMargins(20, 20, 20, 20);
		// three columns of 80 mm: 240 mm on a page with 170 mm between the margins
		let table = AscTest.CreateTable(2, 3, [80, 80, 80]);
		logicDocument.PushToContent(table);
		if (undefined !== ind)
			table.Set_TableInd(ind);
		AscTest.Recalculate();
		AscTest.MoveCursorToParagraph(table.GetRow(0).GetCell(0).GetContent().GetElement(0));
		return table;
	}

	QUnit.module("Fit table to margins");

	QUnit.test("A table that sticks out on both sides is put between the margins", function (assert)
	{
		let table = setup(-15);
		let before = table.GetPageBounds(0);
		assert.true(before.Left < 20 - 1 && before.Right > 190 + 1, "sticks out before (" + before.Left.toFixed(2) + "-" + before.Right.toFixed(2) + ", calc " + table.TableGridCalc.map(function (w) { return w.toFixed(1); }).join("/") + ")");

		assert.true(logicDocument.CanFitTableToMargins(), "can be fitted");
		assert.true(logicDocument.FitTableToMargins(), "fitted");

		let after = table.GetPageBounds(0);
		assert.true(Math.abs(after.Left - table.X) < 1, "starts at the left margin (" + after.Left.toFixed(2) + ")");
		assert.true(after.Right <= table.XLimit + 1, "ends at the right margin (" + after.Right.toFixed(2) + ")");
		assert.true(after.Right - after.Left > 165, "uses the width between the margins (" + (after.Right - after.Left).toFixed(2) + ", X " + table.X.toFixed(2) + ", XLimit " + table.XLimit.toFixed(2) + ", grid " + table.TableGrid.map(function (w) { return w.toFixed(1); }).join("/") + ")");

		let grid = table.TableGrid;
		assert.true(Math.abs(grid[0] - grid[1]) < 0.5 && Math.abs(grid[1] - grid[2]) < 0.5, "columns keep their proportions");
	});

	QUnit.test("Undo brings the old size back", function (assert)
	{
		let table = setup(-15);
		let before = table.GetPageBounds(0);
		logicDocument.FitTableToMargins();
		logicDocument.Document_Undo();
		AscTest.Recalculate();
		let after = table.GetPageBounds(0);
		assert.true(Math.abs(after.Left - before.Left) < 0.1 && Math.abs(after.Right - before.Right) < 0.1, "same bounds as before");
	});

	QUnit.test("Outside a table nothing happens", function (assert)
	{
		setup();
		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		AscTest.MoveCursorToParagraph(paragraph);
		assert.false(logicDocument.CanFitTableToMargins(), "cannot fit");
		assert.false(logicDocument.FitTableToMargins(), "nothing done");
	});
});
