/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	// columns of a DTP/PDF-converted table: one wide column and some narrow ones
	const GRID = [83.3, 7.5, 7.8, 37.1, 9.7, 14.2];

	function sum(arr)
	{
		return arr.reduce(function(a, b){return a + b;}, 0);
	}

	function fmt(arr)
	{
		return arr.map(function(n){return n.toFixed(2);}).join("/");
	}

	// rows 1.. have two cells, the second spans the narrow columns
	function createTable(rowsCount, fixed, pct)
	{
		let table = AscTest.CreateTable(rowsCount, GRID.length, GRID.slice());
		for (let r = 1; r < rowsCount; ++r)
		{
			let row = table.GetRow(r);
			for (let c = GRID.length - 1; c >= 2; --c)
				row.RemoveCell(c);
			row.GetCell(1).SetGridSpan(GRID.length - 1);
		}
		for (let r = 0; r < rowsCount; ++r)
		{
			let row = table.GetRow(r);
			for (let c = 0; c < row.GetCellsCount(); ++c)
			{
				row.GetCell(c).SetW(new CTableMeasurement(tblwidth_Auto, 0));
				AscTest.AddTextToParagraph(row.GetCell(c).GetContent().GetElement(0), "R" + r + "C" + c);
			}
		}
		if (fixed)
			table.SetTableLayout(tbllayout_Fixed);
		if (pct)
			table.Set_TableW(tblwidth_Pct, 100);
		else
			table.Set_TableW(tblwidth_Mm, sum(GRID));
		return table;
	}

	function setup(variant)
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(210, 297);
		sectPr.SetPageMargins(25, 20, 25, 20);

		let table;
		if (variant.simple)
		{
			// a table as Word writes it: widths of the cells in mm, autofit layout
			table = AscTest.CreateTable(2, 3, variant.grid || [50, 50, 50]);
			for (let r = 0; r < 2; ++r)
			{
				for (let c = 0; c < 3; ++c)
				{
					if (variant.cellW)
						table.GetRow(r).GetCell(c).SetW(new CTableMeasurement(tblwidth_Mm, variant.cellW));
					AscTest.AddTextToParagraph(table.GetRow(r).GetCell(c).GetContent().GetElement(0), "R" + r + "C" + c);
				}
			}
			if (variant.pct)
				table.Set_TableW(tblwidth_Pct, 100);
		}
		else
		{
			table = createTable(3, variant.fixed, variant.pct);
		}

		if (variant.nested)
		{
			let outer = AscTest.CreateTable(1, 1, [160]);
			outer.Set_TableW(tblwidth_Mm, 160);
			outer.SetTableLayout(variant.outerAutofit ? tbllayout_AutoFit : tbllayout_Fixed);
			outer.GetRow(0).GetCell(0).GetContent().AddToContent(0, table);
			logicDocument.PushToContent(outer);
		}
		else
		{
			logicDocument.PushToContent(table);
		}
		AscTest.Recalculate();
		return table;
	}

	function rowWidths(table, r)
	{
		let row = table.GetRow(r);
		let res = [];
		for (let c = 0; c < row.GetCellsCount(); ++c)
		{
			let info = row.GetCellInfo(c);
			res.push(info.X_grid_end - info.X_grid_start);
		}
		return res;
	}

	function cellText(table, r, c)
	{
		return AscTest.GetParagraphText(table.GetRow(r).GetCell(c).GetContent().GetElement(0));
	}

	// autofit: the columns are calculated from the contents and the preferred widths again, so the other columns may
	// move; the new part must still not be a narrow strip
	function checkSplit(assert, table, r, c, name, autofit)
	{
		let rowsBefore   = [];
		for (let i = 0; i < table.GetRowsCount(); ++i)
			rowsBefore.push(rowWidths(table, i));
		let tableW       = sum(table.TableGridCalc);
		let cellW        = rowsBefore[r][c];
		let cellsCount   = table.GetRow(r).GetCellsCount();
		let rowsCount    = table.GetRowsCount();
		let text         = cellText(table, r, c);

		AscTest.MoveCursorToParagraph(table.GetRow(r).GetCell(c).GetContent().GetElement(0));
		assert.true(logicDocument.CanSplitTableCells(), name + ": the cell can be split");

		logicDocument.StartAction(AscDFH.historydescription_Document_SplitTableCells);
		logicDocument.SplitTableCells(2, 1);
		logicDocument.FinalizeAction();
		AscTest.Recalculate();

		assert.strictEqual(table.GetRowsCount(), rowsCount, name + ": same number of rows");
		assert.strictEqual(table.GetRow(r).GetCellsCount(), cellsCount + 1, name + ": one cell more in the row");
		assert.strictEqual(cellText(table, r, c), text, name + ": the text stays in the first part");
		assert.strictEqual(cellText(table, r, c + 1), "", name + ": the second part is empty");

		let after = rowWidths(table, r);
		let parts = name + " (" + cellW.toFixed(2) + " -> " + after[c].toFixed(2) + " + " + after[c + 1].toFixed(2) + ")";
		if (autofit)
		{
			assert.ok(after[c] > cellW / 2 * 0.9 && after[c + 1] > cellW / 2 * 0.9, parts + ": both parts are about half as wide or wider, none is a narrow strip");
		}
		else
		{
			assert.ok(Math.abs(after[c] - cellW / 2) < 0.5 && Math.abs(after[c + 1] - cellW / 2) < 0.5, parts + ": both parts are half as wide");
		}
		assert.ok(Math.abs(sum(table.TableGridCalc) - tableW) < 0.5,
			name + ": the table keeps its width (" + tableW.toFixed(2) + " -> " + sum(table.TableGridCalc).toFixed(2) + ")");

		if (autofit)
			return;

		for (let i = 0; i < table.GetRowsCount(); ++i)
		{
			let widths = rowWidths(table, i);
			if (i === r)
				widths.splice(c, 2, widths[c] + widths[c + 1]);
			assert.strictEqual(fmt(widths), fmt(rowsBefore[i]), name + ": row " + i + " keeps its cell widths");
		}
	}

	QUnit.module("Split cell");

	const VARIANTS = [
		{simple : true, name : "simple table"},
		{simple : true, nested : true, outerAutofit : true, name : "simple table nested in an autofit table"},
		// "AutoFit to window": the columns in the file (gridCol) are narrower than the table is drawn
		{simple : true, pct : true, grid : [30, 30, 30], cellW : 54.6, name : "autofit table, width 100%, older column widths in the file"},
		{simple : true, pct : true, autofit : true, name : "autofit table, width 100%"},
		{fixed : true, pct : true, name : "fixed layout, width 100%, cell widths auto"},
		{fixed : false, pct : true, autofit : true, name : "autofit, width 100%, cell widths auto"},
		{fixed : true, pct : false, name : "fixed layout, width in mm, cell widths auto"},
		{fixed : true, pct : true, nested : true, outerAutofit : true, autofit : true, name : "fixed, 100%, nested in an autofit table"},
		{fixed : true, pct : true, nested : true, outerAutofit : false, name : "fixed, 100%, nested in a fixed table"}
	];

	VARIANTS.forEach(function(variant)
	{
		QUnit.test("Split a cell into 2 columns, 1 row: " + variant.name, function (assert)
		{
			let table = setup(variant);
			checkSplit(assert, table, 0, variant.simple ? 1 : 0, variant.name, variant.autofit);
		});

		if (!variant.simple)
		{
			QUnit.test("Split a cell with gridSpan into 2 columns, 1 row: " + variant.name, function (assert)
			{
				let table = setup(variant);
				checkSplit(assert, table, 1, 1, variant.name + ", gridSpan", variant.autofit);
			});
		}
	});

	QUnit.test("The parts get their width as preferred width, undo brings the cell back", function (assert)
	{
		let table = setup({fixed : false, pct : true});
		let cell  = table.GetRow(0).GetCell(0);
		let gridBefore = fmt(table.TableGrid);
		let cellW = rowWidths(table, 0)[0];
		AscTest.MoveCursorToParagraph(cell.GetContent().GetElement(0));
		logicDocument.StartAction(AscDFH.historydescription_Document_SplitTableCells);
		logicDocument.SplitTableCells(2, 1);
		logicDocument.FinalizeAction();
		AscTest.Recalculate();

		for (let c = 0; c < 2; ++c)
		{
			let w = table.GetRow(0).GetCell(c).GetW();
			assert.strictEqual(w.Type, tblwidth_Mm, "part " + (c + 1) + ": preferred width in mm");
			assert.ok(Math.abs(w.W - cellW / 2) < 0.01, "part " + (c + 1) + ": half of the cell");
		}

		logicDocument.Document_Undo();
		AscTest.Recalculate();
		assert.strictEqual(table.GetRow(0).GetCellsCount(), GRID.length, "undo: one cell again");
		assert.strictEqual(table.GetRow(0).GetCell(0).GetW().Type, tblwidth_Auto, "undo: automatic width again");
		assert.strictEqual(fmt(table.TableGrid), gridBefore, "undo: the old grid");
	});
});
