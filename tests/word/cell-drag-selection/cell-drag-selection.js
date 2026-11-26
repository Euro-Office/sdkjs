/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	let drawingDocument = AscTest.DrawingDocument;
	["UpdateTableRuler", "LockCursorTypeCur", "UnlockCursorType", "LockCursorType", "IsMobileVersion"].forEach(function(name)
	{
		if (!drawingDocument[name])
			drawingDocument[name] = function(){};
	});

	function fillTable(table)
	{
		for (let r = 0; r < table.GetRowsCount(); ++r)
		{
			let row = table.GetRow(r);
			for (let c = 0; c < row.GetCellsCount(); ++c)
				AscTest.AddTextToParagraph(row.GetCell(c).GetContent().GetElement(0), "Some text in R" + r + "C" + c);
		}
	}

	// returns the table in which the mouse drags
	function setup(nested)
	{
		AscTest.ClearDocument();
		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(210, 297);
		sectPr.SetPageMargins(25, 20, 25, 20);

		let table = AscTest.CreateTable(2, 3, [50, 50, 50]);
		fillTable(table);
		if (nested)
		{
			let outer = AscTest.CreateTable(1, 1, [160]);
			let content = outer.GetRow(0).GetCell(0).GetContent();
			content.AddToContent(0, table);
			logicDocument.PushToContent(outer);
		}
		else
		{
			logicDocument.PushToContent(table);
		}
		AscTest.Recalculate();
		return table;
	}

	// a point in the text of a cell (offset in mm from the start of the cell content)
	function point(table, nRow, nCell, offset)
	{
		let page = table.Pages[0];
		let info = table.GetRow(nRow).GetCellInfo(nCell);
		return {
			x : page.X + info.X_content_start + (undefined === offset ? 15 : offset),
			y : table.RowsInfo[nRow].Y[0] + table.RowsInfo[nRow].H[0] / 2
		};
	}

	// the document part of the mouse handlers (OnMouseDown/Move/Up need the whole editor UI)
	function mouse(type, p)
	{
		let e = {Button : AscCommon.g_mouse_button_left, ClickCount : 1, Type : type, CtrlKey : false, ShiftKey : false};
		logicDocument.CurPage = 0;
		if (AscCommon.g_mouse_event_type_down === type)
			logicDocument.Selection_SetStart(p.x, p.y, e);
		else
			logicDocument.Selection_SetEnd(p.x, p.y, e);
	}

	// mouse down on the first point, moves over all points, mouse up on the last one
	function drag(points, checkBeforeUp)
	{
		mouse(AscCommon.g_mouse_event_type_down, points[0]);
		for (let i = 1; i < points.length; ++i)
			mouse(AscCommon.g_mouse_event_type_move, points[i]);
		if (checkBeforeUp)
			checkBeforeUp();
		mouse(AscCommon.g_mouse_event_type_up, points[points.length - 1]);
	}

	function cells(table)
	{
		if (!table.IsCellSelection())
			return "text";
		return table.GetSelectionArray().map(function(pos){return pos.Row + ":" + pos.Cell;}).join(",");
	}

	QUnit.module("Drag the mouse across table cells");

	[false, true].forEach(function(nested)
	{
		let name = nested ? " (nested table)" : "";

		QUnit.test("Drag into the next cell and back selects the whole start cell" + name, function (assert)
		{
			let table = setup(nested);
			let a     = point(table, 0, 0, 1);
			let a2    = point(table, 0, 0, 20);
			let b     = point(table, 0, 1);
			drag([a, b, a2], function()
			{
				assert.strictEqual(cells(table), "0:0", "while the mouse is held: whole cell A selected");
			});
			assert.strictEqual(cells(table), "0:0", "after mouse up: cell selection of cell A only");
			assert.true(logicDocument.IsSelectionUse(), "the document has a selection");
			assert.false(logicDocument.IsSelectionEmpty(), "the selection is not empty");

			// a new drag inside a cell selects text again
			drag([a, a2]);
			assert.strictEqual(cells(table), "text", "a new drag inside cell A selects text");
		});

		QUnit.test("Drag inside one cell selects text" + name, function (assert)
		{
			let table = setup(nested);
			drag([point(table, 0, 0, 1), point(table, 0, 0, 10), point(table, 0, 0, 20)]);
			assert.strictEqual(cells(table), "text", "text selection inside the cell");
			assert.true(table.GetRow(0).GetCell(0).GetContent().IsSelectionUse(), "the cell content has a selection");
			assert.false(logicDocument.IsSelectionEmpty(), "some text is selected");
		});

		QUnit.test("Drag into the next cell selects both cells" + name, function (assert)
		{
			let table = setup(nested);
			drag([point(table, 0, 0, 1), point(table, 0, 1)]);
			assert.strictEqual(cells(table), "0:0,0:1", "two cells selected");
		});

		QUnit.test("Drag into the cell below and back selects the whole start cell" + name, function (assert)
		{
			let table = setup(nested);
			drag([point(table, 0, 1, 1), point(table, 1, 1), point(table, 1, 2), point(table, 0, 1, 5)]);
			assert.strictEqual(cells(table), "0:1", "cell selection of the start cell only");
		});
	});

	QUnit.test("Formatting after drag A -> B -> A applies to the whole cell A", function (assert)
	{
		let table = setup(false);
		drag([point(table, 0, 0, 1), point(table, 0, 1), point(table, 0, 0, 5)]);
		assert.strictEqual(cells(table), "0:0", "cell A selected");
		logicDocument.AddToParagraph(new AscCommonWord.ParaTextPr({Bold : true}));
		let runs = table.GetRow(0).GetCell(0).GetContent().GetElement(0).Content.filter(function(run)
		{
			return run instanceof AscCommonWord.ParaRun && run.Content.length > 0 && !(run.Content[0] instanceof AscWord.CRunParagraphMark);
		});
		assert.true(runs.length > 0 && runs.every(function(run){return true === run.Get_CompiledPr(false).Bold;}), "the whole text of cell A is bold");
		let runB = table.GetRow(0).GetCell(1).GetContent().GetElement(0).Content[0];
		assert.strictEqual(runB.Get_CompiledPr(false).Bold, false, "cell B is not bold");
	});
});
