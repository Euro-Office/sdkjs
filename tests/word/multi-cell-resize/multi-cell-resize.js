/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	// the test drawing document has no ruler
	let drawingDocument = AscTest.DrawingDocument;
	["UpdateTableRuler", "LockCursorTypeCur", "UnlockCursorType", "LockCursorType", "CheckTableStyles"].forEach(function(name)
	{
		if (!drawingDocument[name])
			drawingDocument[name] = function(){};
	});
	if (!drawingDocument.CorrectRulerPosition)
		drawingDocument.CorrectRulerPosition = function(pos){return pos;};

	function sum(arr)
	{
		return arr.reduce(function(a, b){return a + b;}, 0);
	}

	function fmt(arr)
	{
		return arr.map(function(n){return n.toFixed(2);}).join("/");
	}

	function near(a, b)
	{
		return Math.abs(a - b) < 0.1;
	}

	// a table with 5 rows and 3 columns of 40 mm; a jagged table: first row with 4 cells, the other rows with 2 cells
	function setup(options)
	{
		options = options || {};
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(210, 297);
		sectPr.SetPageMargins(25, 20, 25, 20);

		let grid  = options.grid || [40, 40, 40];
		let rows  = options.rows || 5;
		let table = AscTest.CreateTable(rows, grid.length, grid.slice());
		if (options.jagged)
		{
			for (let r = 1; r < rows; ++r)
			{
				let row = table.GetRow(r);
				for (let c = grid.length - 1; c >= 2; --c)
					row.RemoveCell(c);
				row.GetCell(1).SetGridSpan(grid.length - 1);
			}
		}
		for (let r = 0; r < rows; ++r)
		{
			let row = table.GetRow(r);
			for (let c = 0; c < row.GetCellsCount(); ++c)
			{
				if (options.autoCells)
					row.GetCell(c).SetW(new CTableMeasurement(tblwidth_Auto, 0));
				AscTest.AddTextToParagraph(row.GetCell(c).GetContent().GetElement(0), "R" + r + "C" + c);
			}
		}
		table.SetTableLayout(options.autofit ? tbllayout_AutoFit : tbllayout_Fixed);
		if (options.pct)
			table.Set_TableW(tblwidth_Pct, 100);
		else
			table.Set_TableW(tblwidth_Mm, sum(grid));

		if (options.nested)
		{
			let outer = AscTest.CreateTable(1, 1, [160]);
			outer.Set_TableW(tblwidth_Mm, 160);
			outer.SetTableLayout(options.outerAutofit ? tbllayout_AutoFit : tbllayout_Fixed);
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

	function allWidths(table)
	{
		let res = [];
		for (let r = 0; r < table.GetRowsCount(); ++r)
			res.push(rowWidths(table, r));
		return res;
	}

	function fmtAll(widths)
	{
		return widths.map(fmt).join(" | ");
	}

	function select(table, startCell, startRow, endCell, endRow)
	{
		// like a selection with the mouse: the table (and its outer table) is the current element, then the cells
		// are selected by their place in the grid
		table.SelectRange(startCell, startRow, endCell, endRow);
		table.Document_SetThisElementCurrent(false);
		table.SelectRange(startCell, startRow, endCell, endRow);
	}

	// drags the border before cell nCell of row nRow by dx with the mouse (down, move, up)
	function dragBorder(table, nRow, nCell, dx)
	{
		let page  = table.Pages[0];
		let x     = page.X + table.GetRow(nRow).GetCellInfo(nCell).X_grid_start;
		let y     = table.RowsInfo[nRow].Y[0] + table.RowsInfo[nRow].H[0] / 2;
		let event = {ClickCount : 1, Type : AscCommon.g_mouse_event_type_down, CtrlKey : false};

		let hit = table.private_CheckHitInBorder(x, y, 0);
		table.Selection_SetStart(x, y, 0, event);
		event.Type = AscCommon.g_mouse_event_type_move;
		table.Selection_SetEnd(x + dx / 2, y, 0, event);
		event.Type = AscCommon.g_mouse_event_type_up;
		table.Selection_SetEnd(x + dx, y, 0, event);
		AscTest.Recalculate();
		return hit;
	}

	// what the Width spin of the right panel and the table advanced settings do: api.tblApply
	function applyProps(fill)
	{
		let props = new Asc.CTableProp();
		fill(props);
		logicDocument.StartAction(AscDFH.historydescription_Document_ApplyTablePr);
		logicDocument.SetTableProps(props);
		logicDocument.FinalizeAction();
		AscTest.Recalculate();
	}

	function undo()
	{
		AscTest.Undo ? AscTest.Undo() : logicDocument.Document_Undo();
		AscTest.Recalculate();
	}

	QUnit.module("Resize several selected table cells");

	QUnit.test("Dragging a border of a column of 3 selected cells moves it only in those 3 rows", function (assert)
	{
		let table  = setup();
		let before = allWidths(table);
		select(table, 0, 1, 0, 3);
		assert.strictEqual(table.GetSelectionArray().length, 3, "3 cells selected");

		let hit = dragBorder(table, 2, 1, 10);
		assert.strictEqual(hit.Border, 3, "the mouse hits the left border of the second cell");

		let after = allWidths(table);
		let info  = " (" + fmtAll(before) + " -> " + fmtAll(after) + ")";
		for (let r = 1; r <= 3; ++r)
		{
			assert.true(near(after[r][0], 50) && near(after[r][1], 30) && near(after[r][2], 40), "row " + r + ": the border moved by 10 mm" + info);
		}
		[0, 4].forEach(function(r)
		{
			assert.strictEqual(fmt(after[r]), fmt(before[r]), "row " + r + " (not selected) keeps its border" + info);
		});
		assert.true(near(sum(after[0]), sum(before[0])) && near(sum(after[2]), sum(before[2])), "all rows keep their width" + info);
		assert.true(table.IsCellSelection(), "the cells are still selected");

		undo();
		assert.strictEqual(fmtAll(allWidths(table)), fmtAll(before), "undo restores all rows");
	});

	QUnit.test("Dragging a border of a block of selected cells moves it only in those rows (nested in an autofit table, width 100%)", function (assert)
	{
		let table  = setup({nested : true, outerAutofit : true, pct : true});
		let before = allWidths(table);
		select(table, 0, 2, 1, 4);
		let hit = dragBorder(table, 3, 2, -10);
		assert.strictEqual(hit.Border, 3, "the mouse hits the left border of the third cell");
		let after = allWidths(table);
		let info  = " (" + fmtAll(before) + " -> " + fmtAll(after) + ")";
		for (let r = 2; r <= 4; ++r)
			assert.true(near(after[r][1], before[r][1] - 10) && near(after[r][2], before[r][2] + 10), "row " + r + ": the border moved 10 mm to the left" + info);
		[0, 1].forEach(function(r)
		{
			assert.strictEqual(fmt(after[r]), fmt(before[r]), "row " + r + " keeps its borders" + info);
		});
		assert.true(near(sum(table.TableGridCalc), sum(before[0])), "the table keeps its width" + info);
	});

	QUnit.test("Width field with 4 selected cells: all 4 get the value, the other rows stay", function (assert)
	{
		let table  = setup();
		let before = allWidths(table);
		select(table, 0, 1, 1, 2);
		assert.strictEqual(table.GetSelectionArray().length, 4, "4 cells selected");

		applyProps(function(props){props.put_ColumnWidth(30);});

		let after = allWidths(table);
		let info  = " (" + fmtAll(before) + " -> " + fmtAll(after) + ")";
		[1, 2].forEach(function(r)
		{
			assert.true(near(after[r][0], 30) && near(after[r][1], 30), "row " + r + ": both selected cells are 30 mm" + info);
			assert.true(near(after[r][2], 40), "row " + r + ": the cell right of the selection keeps its width" + info);
		});
		[0, 3, 4].forEach(function(r)
		{
			assert.strictEqual(fmt(after[r]), fmt(before[r]), "row " + r + " (not selected) is unchanged" + info);
		});
		assert.true(near(table.Get_Props().ColumnWidth, 30), "the Width field shows 30 mm");

		undo();
		assert.strictEqual(fmtAll(allWidths(table)), fmtAll(before), "undo restores all rows");
	});

	QUnit.test("Width field without a cell selection still sets the whole column", function (assert)
	{
		let table  = setup();
		table.RemoveSelection();
		table.CurCell = table.GetRow(2).GetCell(1);
		table.Document_SetThisElementCurrent(false);
		applyProps(function(props){props.put_ColumnWidth(25);});
		let after = allWidths(table);
		for (let r = 0; r < 5; ++r)
			assert.true(near(after[r][1], 25), "row " + r + ": the second cell is 25 mm (" + fmtAll(after) + ")");
	});

	QUnit.test("Width field in a jagged table only changes the selected spanned cells", function (assert)
	{
		let table  = setup({jagged : true, grid : [50, 10, 20, 40]});
		let before = allWidths(table);
		select(table, 1, 1, 1, 3);
		applyProps(function(props){props.put_ColumnWidth(60);});
		let after = allWidths(table);
		let info  = " (" + fmtAll(before) + " -> " + fmtAll(after) + ")";
		for (let r = 1; r <= 3; ++r)
			assert.true(near(after[r][1], 60), "row " + r + ": the selected cell is 60 mm" + info);
		assert.strictEqual(fmt(after[0]), fmt(before[0]), "the first row with 4 cells is unchanged" + info);
		assert.strictEqual(fmt(after[4]), fmt(before[4]), "the last row is unchanged" + info);
	});

	QUnit.test("Width field in a table nested in an autofit table (width 100%) is kept", function (assert)
	{
		let table  = setup({nested : true, outerAutofit : true, pct : true});
		let outer  = table.Parent.IsInTable(true);
		let before = allWidths(table);
		select(table, 1, 0, 1, 4);
		applyProps(function(props){props.put_ColumnWidth(30);});
		let after = allWidths(table);
		let info  = " (" + fmtAll(before) + " -> " + fmtAll(after) + ")";
		for (let r = 0; r < 5; ++r)
			assert.true(near(after[r][1], 30), "row " + r + ": the selected cell is 30 mm" + info);
		assert.true(sum(table.TableGridCalc) <= sum(outer.TableGridCalc) + 0.1, "the nested table fits into the outer table (" + fmt(table.TableGridCalc) + " in " + fmt(outer.TableGridCalc) + ")");
	});

	QUnit.test("Width field shows mixed widths as empty and equal widths as the value", function (assert)
	{
		let table = setup({grid : [30, 40, 50]});
		select(table, 0, 1, 1, 1);
		assert.strictEqual(table.Get_Props().ColumnWidth, undefined, "30 and 40 mm: the Width field is empty");
		select(table, 1, 0, 1, 4);
		assert.true(near(table.Get_Props().ColumnWidth, 40), "a column of 40 mm cells: the Width field shows 40");

		// a jagged column: the selected cells are equal, cells of other rows at the same index are not
		table = setup({jagged : true, grid : [50, 10, 20, 40]});
		select(table, 1, 2, 1, 4);
		assert.true(near(table.Get_Props().ColumnWidth, 70), "selected spanned cells of 70 mm: the Width field shows 70");
	});

	QUnit.test("Table advanced settings: cell width with 4 selected cells", function (assert)
	{
		let table = setup();
		select(table, 0, 1, 1, 2);
		let props = table.Get_Props();
		assert.false(props.CellsWidthNotEqual, "4 cells of 40 mm: the cell width is not shown as mixed");
		assert.true(near(props.CellsWidth, 40), "the cell width is 40 mm");

		let before = allWidths(table);
		applyProps(function(p){p.put_CellsWidth(25);});
		let after = allWidths(table);
		let info  = " (" + fmtAll(before) + " -> " + fmtAll(after) + ")";
		[1, 2].forEach(function(r)
		{
			assert.true(near(after[r][0], 25) && near(after[r][1], 25), "row " + r + ": both selected cells are 25 mm" + info);
		});
		[0, 3, 4].forEach(function(r)
		{
			assert.strictEqual(fmt(after[r]), fmt(before[r]), "row " + r + " (not selected) is unchanged" + info);
		});
		props = table.Get_Props();
		assert.false(props.CellsWidthNotEqual, "the cell width is not mixed");
		assert.true(near(props.CellsWidth, 25), "the dialog shows 25 mm");

		select(table, 0, 1, 2, 1);
		assert.true(table.Get_Props().CellsWidthNotEqual, "25, 25 and 40 mm: the cell width is mixed");

		undo();
		assert.strictEqual(fmtAll(allWidths(table)), fmtAll(before), "undo restores all rows");
	});

	QUnit.test("Table advanced settings: cell width in percent with 2 selected cells", function (assert)
	{
		let table = setup();
		select(table, 2, 1, 2, 2);
		applyProps(function(p){p.put_CellsWidth(-10);});
		let pctW  = table.private_RecalculatePercentWidth() / 10;
		let after = allWidths(table);
		[1, 2].forEach(function(r)
		{
			assert.true(near(after[r][2], pctW), "row " + r + ": the last cell is 10% (" + pctW.toFixed(2) + " mm) wide (" + fmtAll(after) + ")");
			assert.strictEqual(table.GetRow(r).GetCell(2).GetW().Type, tblwidth_Pct, "row " + r + ": the preferred width stays in percent");
		});
		assert.true(near(after[0][2], 40), "the first row keeps its last cell");
	});

	QUnit.test("Width field keeps the indent of a row that starts later in the grid", function (assert)
	{
		let table = setup();
		let row0  = table.GetRow(0);
		row0.RemoveCell(0);
		row0.SetBefore(1);
		AscTest.Recalculate();
		let start = row0.GetCellInfo(0).X_grid_start;
		select(table, 1, 1, 1, 2);
		applyProps(function(props){props.put_ColumnWidth(30);});
		let after = allWidths(table);
		assert.true(near(row0.GetCellInfo(0).X_grid_start, start), "the first row still starts at " + start.toFixed(2) + " mm (now " + row0.GetCellInfo(0).X_grid_start.toFixed(2) + ")");
		assert.strictEqual(fmt(after[0]), "40.00/40.00", "the first row keeps its cells (" + fmtAll(after) + ")");
		assert.true(near(after[1][1], 30) && near(after[2][1], 30), "the selected cells are 30 mm (" + fmtAll(after) + ")");
	});

	QUnit.test("Width field with a selected cell merged vertically changes all its rows", function (assert)
	{
		let table = setup();
		table.GetRow(1).GetCell(0).SetVMerge(vmerge_Restart);
		table.GetRow(2).GetCell(0).SetVMerge(vmerge_Continue);
		AscTest.Recalculate();
		select(table, 0, 1, 0, 1);
		applyProps(function(props){props.put_ColumnWidth(25);});
		let after = allWidths(table);
		assert.true(near(after[1][0], 25) && near(after[2][0], 25), "both rows of the merged cell are 25 mm (" + fmtAll(after) + ")");
		assert.true(near(after[0][0], 40) && near(after[3][0], 40), "the other rows stay (" + fmtAll(after) + ")");
	});

	QUnit.test("Height field with 3 selected rows sets all 3 rows", function (assert)
	{
		let table = setup();
		select(table, 0, 1, 2, 3);
		applyProps(function(p){p.put_RowHeight(35);});
		for (let r = 1; r <= 3; ++r)
			assert.true(near(table.GetRow(r).GetHeight().GetValue(), 35), "row " + r + " is 35 mm high");
		assert.true(table.GetRow(0).GetHeight().IsAuto() && table.GetRow(4).GetHeight().IsAuto(), "the other rows stay auto");
		assert.true(near(table.Get_Props().RowHeight, 35), "the Height field shows 35 mm");
		select(table, 0, 0, 0, 2);
		assert.strictEqual(table.Get_Props().RowHeight, undefined, "mixed heights: the Height field is empty");
		undo();
		assert.true(table.GetRow(2).GetHeight().IsAuto(), "undo restores the height");
	});
});
