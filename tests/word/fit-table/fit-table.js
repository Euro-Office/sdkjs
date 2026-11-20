/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();
	let compatibilityMode = logicDocument.GetCompatibilityMode();

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

	function fmt(n)
	{
		return n.toFixed(2);
	}

	function sticksOut(assert, table)
	{
		let b = table.GetPageBounds(0);
		assert.true(b.Left < table.X - 3 || b.Right > table.XLimit + 3, "sticks out before (" + fmt(b.Left) + "-" + fmt(b.Right) + ", margins " + fmt(table.X) + "-" + fmt(table.XLimit) + ")");
	}

	// Like Word, before compatibility mode 15 the cell margins (and borders) stand out of the text margins on both
	// sides, so that the text in the cells lines up with the text outside; from mode 15 the table edges are on the
	// margins.
	function checkFits(assert, table, label)
	{
		let b     = table.GetPageBounds(0);
		let left  = table.X - Math.abs(table.GetTableOffsetCorrection());
		let right = table.XLimit + Math.abs(table.GetRightTableOffsetCorrection());
		let info  = " (" + fmt(b.Left) + "-" + fmt(b.Right) + ", expected " + fmt(left) + "-" + fmt(right) + ", margins " + fmt(table.X) + "-" + fmt(table.XLimit) + ", grid " + table.TableGridCalc.map(fmt).join("/") + ")";
		assert.true(Math.abs(b.Left - left) < 0.1, label + ": starts at the left margin" + info);
		assert.true(Math.abs(b.Right - right) < 0.1, label + ": ends at the right margin" + info);
	}

	function sameColumns(assert, table)
	{
		let grid = table.TableGridCalc;
		assert.true(Math.abs(grid[0] - grid[1]) < 0.5 && Math.abs(grid[1] - grid[2]) < 0.5, "columns keep their proportions (" + grid.map(fmt).join("/") + ")");
	}

	function fit(assert)
	{
		assert.true(logicDocument.CanFitTableToMargins(), "can be fitted");
		assert.true(logicDocument.FitTableToMargins(), "fitted");
		AscTest.Recalculate();
	}

	QUnit.module("Fit table to margins");

	QUnit.test("A table that sticks out on both sides is put between the margins", function (assert)
	{
		let table = setup(-15);
		sticksOut(assert, table);
		fit(assert);
		checkFits(assert, table, "fitted");
		sameColumns(assert, table);

		let tablePr = table.Get_CompiledPr(false).TablePr;
		assert.strictEqual(tablePr.TableW.Type, tblwidth_Pct, "preferred width in percent");
		assert.strictEqual(tablePr.TableW.W, 100, "of 100%");
		assert.strictEqual(table.Get_TableInd(), 0, "indent 0");
	});

	QUnit.test("Fixed layout with fixed cell widths", function (assert)
	{
		let table = setup(-15);
		table.SetTableLayout(tbllayout_Fixed);
		for (let r = 0; r < table.GetRowsCount(); ++r)
		{
			for (let c = 0; c < 3; ++c)
				table.GetRow(r).GetCell(c).SetW(new CTableMeasurement(tblwidth_Mm, 80));
		}
		AscTest.Recalculate();
		sticksOut(assert, table);
		fit(assert);
		checkFits(assert, table, "fitted");
		sameColumns(assert, table);
	});

	QUnit.test("A table with a percent width", function (assert)
	{
		let table = setup();
		table.Set_TableW(tblwidth_Pct, 140);
		for (let r = 0; r < table.GetRowsCount(); ++r)
		{
			for (let c = 0; c < 3; ++c)
				table.GetRow(r).GetCell(c).SetW(new CTableMeasurement(tblwidth_Pct, 33.3));
		}
		AscTest.Recalculate();
		sticksOut(assert, table);
		fit(assert);
		checkFits(assert, table, "fitted");
		sameColumns(assert, table);
	});

	QUnit.test("A centred table", function (assert)
	{
		let table = setup();
		table.Set_TableAlign(AscCommon.align_Center);
		table.SetTableLayout(tbllayout_Fixed);
		AscTest.Recalculate();
		let b = table.GetPageBounds(0);
		assert.true(b.Left < table.X - 3 && b.Right > table.XLimit + 3, "sticks out on both sides before (" + fmt(b.Left) + "-" + fmt(b.Right) + ")");
		fit(assert);
		checkFits(assert, table, "fitted");
		sameColumns(assert, table);
	});

	QUnit.test("The table keeps fitting when the margins or the orientation change", function (assert)
	{
		let table = setup(-15);
		fit(assert);
		let sectPr = AscTest.GetFinalSection();

		sectPr.SetPageMargins(40, 20, 40, 20);
		AscTest.Recalculate();
		checkFits(assert, table, "wider margins");

		sectPr.SetPageMargins(10, 20, 10, 20);
		AscTest.Recalculate();
		checkFits(assert, table, "narrower margins");

		sectPr.SetPageSize(297, 210);
		AscTest.Recalculate();
		checkFits(assert, table, "landscape");
	});

	QUnit.test("In compatibility mode 15 the table edges are on the margins", function (assert)
	{
		AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Word15);
		let table = setup(-15);
		sticksOut(assert, table);
		fit(assert);
		let b = table.GetPageBounds(0);
		assert.true(Math.abs(b.Left - table.X) < 0.1 && Math.abs(b.Right - table.XLimit) < 0.1, "between the margins (" + fmt(b.Left) + "-" + fmt(b.Right) + ")");
		sameColumns(assert, table);
		AscTest.SetCompatibilityMode(compatibilityMode);
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

	QUnit.test("A table that is already fitted is recognised and left as it is", function (assert)
	{
		let table = setup(-15);
		assert.false(logicDocument.IsTableFittedToMargins(), "not fitted before");
		fit(assert);
		assert.true(logicDocument.IsTableFittedToMargins(), "fitted afterwards");
		assert.true(logicDocument.CanFitTableToMargins(), "the command itself is still possible");

		let pointsBefore = AscCommon.History.Points.length;
		assert.false(logicDocument.FitTableToMargins(), "a second fit changes nothing");
		assert.strictEqual(AscCommon.History.Points.length, pointsBefore, "and adds no undo step");
		checkFits(assert, table, "still fitted");

		table.Set_TableInd(5);
		AscTest.Recalculate();
		assert.false(logicDocument.IsTableFittedToMargins(), "an indent again: not fitted");
	});

	QUnit.test("Outside a table nothing happens", function (assert)
	{
		setup();
		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		AscTest.MoveCursorToParagraph(paragraph);
		assert.false(logicDocument.CanFitTableToMargins(), "cannot fit");
		assert.false(logicDocument.IsTableFittedToMargins(), "not a fitted table");
		assert.false(logicDocument.FitTableToMargins(), "nothing done");
	});
});
