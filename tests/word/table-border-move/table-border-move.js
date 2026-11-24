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
	["UpdateTableRuler", "LockCursorTypeCur", "UnlockCursorType", "LockCursorType"].forEach(function(name)
	{
		if (!drawingDocument[name])
			drawingDocument[name] = function(){};
	});
	if (!drawingDocument.CorrectRulerPosition)
		drawingDocument.CorrectRulerPosition = function(pos){return pos;};

	// the grid of the real table: a first row with 6 cells, the other rows with 2 cells (the second one spans 5 columns)
	const GRID = [83.3, 7.5, 7.8, 37.1, 9.7, 14.2];

	function sum(arr)
	{
		return arr.reduce(function(a, b){return a + b;}, 0);
	}

	function fmt(arr)
	{
		return arr.map(function(n){return n.toFixed(2);}).join("/");
	}

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

	let otherTable = null;

	function setup(nested, fixed, pct, outerAutofit)
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(210, 297);
		sectPr.SetPageMargins(25, 20, 25, 20);

		let table = createTable(3, fixed, pct);
		if (nested)
		{
			let outer = AscTest.CreateTable(1, 1, [160]);
			outer.Set_TableW(tblwidth_Mm, 160);
			outer.SetTableLayout(outerAutofit ? tbllayout_AutoFit : tbllayout_Fixed);
			let content = outer.GetRow(0).GetCell(0).GetContent();
			content.AddToContent(0, table);
			// another nested table in the same cell, whose columns are calculated from its contents
			otherTable = AscTest.CreateTable(1, 2, [100, 59.6]);
			AscTest.AddTextToParagraph(otherTable.GetRow(0).GetCell(0).GetContent().GetElement(0), "Other table");
			otherTable.Set_TableW(tblwidth_Pct, 100);
			content.AddToContent(content.GetElementsCount(), otherTable);
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

	function rowBorders(table, r)
	{
		let row = table.GetRow(r);
		let res = [];
		for (let c = 0; c < row.GetCellsCount(); ++c)
			res.push(row.GetCellInfo(c).X_grid_start);
		res.push(row.GetCellInfo(row.GetCellsCount() - 1).X_grid_end);
		return res;
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

	function checkMove(assert, table, dx, label)
	{
		let gridBefore   = table.TableGridCalc.slice();
		let row0Before   = rowBorders(table, 0);
		let row1Before   = rowWidths(table, 1);
		let totalBefore  = sum(gridBefore);
		let colsBefore   = table.TableGrid.length;

		let hit = dragBorder(table, 1, 1, dx);
		assert.strictEqual(hit.Border, 3, label + ": the mouse hits the left border of the second cell");

		let grid    = table.TableGridCalc;
		let row1    = rowWidths(table, 1);
		let row0    = rowBorders(table, 0);
		let info    = " (grid " + fmt(gridBefore) + " -> " + fmt(grid) + ")";

		assert.true(Math.abs(sum(grid) - totalBefore) < 0.1, label + ": the table keeps its width " + totalBefore.toFixed(2) + info);
		assert.true(Math.abs(row1[0] - (row1Before[0] + dx)) < 0.1, label + ": the left cell is " + dx + " mm wider" + info);
		assert.true(Math.abs(row1[1] - (row1Before[1] - dx)) < 0.1, label + ": the spanned right cell is " + dx + " mm narrower" + info);
		assert.true(Math.abs(row0[row0.length - 1] - row0Before[row0Before.length - 1]) < 0.1, label + ": the first row still ends at the same place" + info);
		assert.true(table.TableGrid.length <= colsBefore, label + ": no extra grid column" + info);
		return {grid : gridBefore, row0 : row0Before, row1 : row1Before};
	}

	function checkUndo(assert, table, before, label)
	{
		AscTest.Undo ? AscTest.Undo() : logicDocument.Document_Undo();
		AscTest.Recalculate();
		assert.strictEqual(fmt(table.TableGridCalc), fmt(before.grid), label + ": undo restores the grid");
	}

	QUnit.module("Move a table border");

	[
		{nested : true, fixed : true, pct : true, name : "nested, fixed layout, width 100%"},
		{nested : false, fixed : true, pct : true, name : "fixed layout, width 100%"},
		{nested : false, fixed : true, pct : false, name : "fixed layout, width in mm"},
		{nested : false, fixed : false, pct : false, name : "autofit, width in mm"},
		{nested : true, fixed : false, pct : true, name : "nested, autofit, width 100%"},
		{nested : true, fixed : true, pct : true, outerAutofit : true, name : "nested in an autofit table, fixed layout, width 100%"},
		{nested : true, fixed : false, pct : true, outerAutofit : true, name : "nested in an autofit table, autofit, width 100%"}
	].forEach(function(variant)
	{
		QUnit.test("Moving the border to a cell that spans several columns keeps the table width (" + variant.name + ")", function (assert)
		{
			let table  = setup(variant.nested, variant.fixed, variant.pct, variant.outerAutofit);
			let before = checkMove(assert, table, 15, "15 mm to the right");
			checkUndo(assert, table, before, variant.name);
		});
	});

	QUnit.test("Moving the border to the left makes the spanned cell wider", function (assert)
	{
		let table  = setup(true, true, true);
		let before = checkMove(assert, table, -20, "20 mm to the left");
		checkUndo(assert, table, before, "to the left");
	});

	QUnit.test("A border shared by all rows moves in all rows, the next border stays", function (assert)
	{
		let table = setup(false, true, false);
		let row0  = rowBorders(table, 0);
		let row2  = rowBorders(table, 2);
		dragBorder(table, 1, 1, 3);
		let newRow0 = rowBorders(table, 0);
		let newRow2 = rowBorders(table, 2);
		assert.true(Math.abs(newRow2[1] - (row2[1] + 3)) < 0.1, "the border in the third row moves too");
		assert.true(Math.abs(newRow0[1] - (row0[1] + 3)) < 0.1, "the border in the first row moves too");
		assert.true(Math.abs(newRow0[2] - row0[2]) < 0.1, "the next border of the first row stays (" + fmt(row0) + " -> " + fmt(newRow0) + ")");
		assert.true(Math.abs(newRow0[newRow0.length - 1] - row0[row0.length - 1]) < 0.1, "the first row ends at the same place");
	});

	QUnit.test("If the next cell of another row is too narrow, the following cells of that row give up the rest", function (assert)
	{
		let table = setup(false, true, false);
		let row0  = rowBorders(table, 0);
		dragBorder(table, 1, 1, 15);
		let newRow0 = rowBorders(table, 0);
		let widths  = rowWidths(table, 0);
		let info    = " (" + fmt(row0) + " -> " + fmt(newRow0) + ")";
		assert.true(Math.abs(newRow0[1] - (row0[1] + 15)) < 0.1, "the border moves by 15 mm" + info);
		assert.true(widths[1] > 3 && widths[2] > 3, "the squeezed cells keep their minimum width" + info);
		assert.true(Math.abs(newRow0[4] - row0[4]) < 0.1 && Math.abs(newRow0[5] - row0[5]) < 0.1, "the borders far right stay" + info);
		assert.true(Math.abs(newRow0[newRow0.length - 1] - row0[row0.length - 1]) < 0.1, "the first row ends at the same place" + info);
	});

	QUnit.test("Moving a border of selected cells keeps the table width", function (assert)
	{
		let table = setup(true, true, true);
		table.SelectAll();
		assert.true(table.IsCellSelection(), "all cells selected");
		let before = checkMove(assert, table, 15, "selected cells, 15 mm to the right");
		checkUndo(assert, table, before, "selected cells");
	});

	QUnit.test("In a table nested in an autofit table, the move is kept and the other nested table keeps its columns", function (assert)
	{
		let table      = setup(true, true, true, true);
		let outer      = table.Parent.IsInTable(true);
		let otherGrid  = fmt(otherTable.TableGridCalc);
		let outerGrid  = fmt(outer.TableGridCalc);
		let before     = checkMove(assert, table, 15, "nested in an autofit table");
		assert.strictEqual(fmt(otherTable.TableGridCalc), otherGrid, "the other nested table keeps its columns");
		assert.strictEqual(fmt(outer.TableGridCalc), outerGrid, "the outer table keeps its columns");
		assert.strictEqual(outer.Get_CompiledPr(false).TablePr.TableLayout, tbllayout_Fixed, "the outer table has fixed column widths now");
		checkUndo(assert, table, before, "nested in an autofit table");
		assert.strictEqual(outer.Get_CompiledPr(false).TablePr.TableLayout, tbllayout_AutoFit, "undo: the outer table fits its contents again");
	});

	QUnit.test("Moving a column marker on the ruler works in a table nested in an autofit table", function (assert)
	{
		let table = setup(true, true, true, true);
		let grid  = table.TableGridCalc.slice();
		table.private_UpdateTableMarkup(1, 0, 0);
		let markup = {X : table.Markup.X, Cols : table.Markup.Cols.slice(), Margins : table.Markup.Margins, Rows : table.Markup.Rows, Internal : table.Markup.Internal, Table : table};
		markup.Cols[0] += 5;
		logicDocument.StartAction(AscDFH.historydescription_Document_SetTableMarkup_Hor);
		table.Update_TableMarkupFromRuler(markup, true, 1);
		AscTest.Recalculate();
		logicDocument.FinalizeAction();
		assert.true(Math.abs(table.TableGridCalc[0] - (grid[0] + 5)) < 0.1, "the first column is 5 mm wider (" + fmt(grid) + " -> " + fmt(table.TableGridCalc) + ")");
	});

	// The outer table (autofit, width 100 %) has two rows. The first row has the dragged table and another table
	// (cells in percent), the second row has an unrelated table (cells in mm). The outer cells have widths in percent.
	function setupAround()
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(210, 297);
		sectPr.SetPageMargins(25, 20, 25, 20);

		let table = createTable(3, true, true);

		let outer = AscTest.CreateTable(2, 1, [160]);
		outer.Set_TableW(tblwidth_Pct, 100);
		outer.SetTableLayout(tbllayout_AutoFit);
		outer.GetRow(0).GetCell(0).SetW(new CTableMeasurement(tblwidth_Pct, 100));
		outer.GetRow(1).GetCell(0).SetW(new CTableMeasurement(tblwidth_Pct, 100));

		let sibling = AscTest.CreateTable(1, 2, [100, 59.6]);
		sibling.Set_TableW(tblwidth_Pct, 100);
		sibling.SetTableLayout(tbllayout_AutoFit);
		sibling.GetRow(0).GetCell(0).SetW(new CTableMeasurement(tblwidth_Pct, 60));
		sibling.GetRow(0).GetCell(1).SetW(new CTableMeasurement(tblwidth_Pct, 40));
		AscTest.AddTextToParagraph(sibling.GetRow(0).GetCell(0).GetContent().GetElement(0), "Sibling table");

		let other = AscTest.CreateTable(1, 3, [40, 60, 59.6]);
		other.SetTableLayout(tbllayout_AutoFit);
		AscTest.AddTextToParagraph(other.GetRow(0).GetCell(1).GetContent().GetElement(0), "Unrelated table");

		let content0 = outer.GetRow(0).GetCell(0).GetContent();
		content0.AddToContent(0, table);
		content0.AddToContent(content0.GetElementsCount(), sibling);
		let content1 = outer.GetRow(1).GetCell(0).GetContent();
		content1.AddToContent(0, other);

		logicDocument.PushToContent(outer);
		// the autofit widths of the nested tables settle only in the second pass
		AscTest.Recalculate();
		AscTest.Recalculate();
		return {table : table, outer : outer, sibling : sibling, other : other};
	}

	function cellWidths(table)
	{
		let res = [];
		for (let r = 0; r < table.GetRowsCount(); ++r)
		{
			let row = table.GetRow(r);
			for (let c = 0; c < row.GetCellsCount(); ++c)
			{
				let w = row.GetCell(c).GetW();
				res.push((tblwidth_Pct === w.Type ? "pct " : tblwidth_Mm === w.Type ? "mm " : "auto ") + w.W.toFixed(2));
			}
		}
		return res.join(", ");
	}

	function cellTypes(table)
	{
		return cellWidths(table).split(", ").map(function(w){return w.split(" ")[0];}).join(", ");
	}

	function tableState(table)
	{
		let pr = table.Get_CompiledPr(false).TablePr;
		return {
			grid   : fmt(table.TableGridCalc),
			layout : pr.TableLayout,
			tableW : (tblwidth_Pct === pr.TableW.Type ? "pct " : "") + pr.TableW.W.toFixed(2),
			cells  : cellWidths(table),
			types  : cellTypes(table)
		};
	}

	QUnit.test("Only the tables that contain the dragged table get fixed column widths, the other tables keep their layout and percent widths", function (assert)
	{
		let t       = setupAround();
		let outer   = tableState(t.outer);
		let sibling = tableState(t.sibling);
		let other   = tableState(t.other);

		let before = checkMove(assert, t.table, 15, "nested next to other tables");

		let newOuter = tableState(t.outer);
		assert.strictEqual(newOuter.layout, tbllayout_Fixed, "the outer table has fixed column widths now");
		assert.strictEqual(newOuter.grid, outer.grid, "the outer table keeps its columns");
		assert.strictEqual(newOuter.tableW, outer.tableW, "the outer table keeps its width in percent");
		assert.strictEqual(newOuter.types, outer.types, "the cells of the outer table keep their widths in percent (" + newOuter.cells + ")");

		let newSibling = tableState(t.sibling);
		assert.strictEqual(newSibling.layout, sibling.layout, "the other nested table keeps its layout");
		assert.strictEqual(newSibling.grid, sibling.grid, "the other nested table keeps its columns");
		assert.strictEqual(newSibling.tableW, sibling.tableW, "the other nested table keeps its width in percent");
		assert.strictEqual(newSibling.types, sibling.types, "the cells of the other nested table keep their widths in percent (" + sibling.cells + " -> " + newSibling.cells + ")");

		let newOther = tableState(t.other);
		assert.strictEqual(newOther.layout, other.layout, "the table in the other outer cell keeps its layout");
		assert.strictEqual(newOther.grid, other.grid, "the table in the other outer cell keeps its columns");
		assert.strictEqual(newOther.types, other.types, "the cells of the table in the other outer cell keep their width type");

		checkUndo(assert, t.table, before, "nested next to other tables");
		assert.deepEqual(tableState(t.outer), outer, "undo restores the outer table");
		assert.strictEqual(JSON.stringify(tableState(t.sibling)), JSON.stringify(sibling), "undo restores the other nested table");
		assert.deepEqual(tableState(t.other), other, "undo restores the table in the other outer cell");
	});

	QUnit.test("The lock check covers every table that changes", function (assert)
	{
		let t = setupAround();
		assert.deepEqual(t.table.GetTablesChangedByColumnResize(), [t.table, t.outer, t.sibling, t.other], "in an autofit table: the dragged table and all tables of the top table");

		let lockData = null;
		let isSelectionLocked = logicDocument.Document_Is_SelectionLocked;
		logicDocument.Document_Is_SelectionLocked = function(checkType, additionalData)
		{
			lockData = additionalData;
			return true;
		};

		let outer = tableState(t.outer);
		let grid  = fmt(t.table.TableGridCalc);
		try
		{
			dragBorder(t.table, 1, 1, 15);
		}
		finally
		{
			logicDocument.Document_Is_SelectionLocked = isSelectionLocked;
		}

		assert.true(!!lockData && AscCommon.changestype_2_ElementsArray_and_Type === lockData.Type, "the tables are checked as a list");
		assert.true(!!lockData && Array.isArray(lockData.Elements) && -1 !== lockData.Elements.indexOf(t.table) && -1 !== lockData.Elements.indexOf(t.outer), "the dragged table and the outer table are checked");
		assert.strictEqual(fmt(t.table.TableGridCalc), grid, "locked: the dragged table doesn't change");
		assert.deepEqual(tableState(t.outer), outer, "locked: the outer table doesn't change");

		let table = setup(true, true, true, false);
		assert.deepEqual(table.GetTablesChangedByColumnResize(), [table], "in a fixed table: only the dragged table");
		table = setup(false, true, true);
		assert.deepEqual(table.GetTablesChangedByColumnResize(), [table], "a table on its own: only the dragged table");
	});
});
