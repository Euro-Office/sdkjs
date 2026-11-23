/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	window["NATIVE_EDITOR_ENJINE_SYNC_RECALC"] = true;
	let logicDocument = AscTest.CreateLogicDocument();

	// A graphics object that only records what the table draws
	function recordingGraphics(log)
	{
		return new Proxy({}, {
			get : function(target, name)
			{
				if ("SetBorder" === name || "then" === name || typeof name !== "string")
					return undefined;
				return function()
				{
					log.push({name : name, args : Array.prototype.slice.call(arguments)});
				};
			}
		});
	}

	function round(x)
	{
		return Math.round(x * 10) / 10;
	}

	/**
	 * The x positions of the vertical lines the table draws on its first page
	 * @returns {{gridlines : number[], lines : number[]}} gridlines (shown for invisible borders) and visible lines
	 */
	function getVerticalLines(table, page)
	{
		let log = [];
		table.Draw(page || 0, recordingGraphics(log));
		let result = {gridlines : [], lines : []};
		log.forEach(function(item)
		{
			if ("DrawEmptyTableLine" === item.name && Math.abs(item.args[0] - item.args[2]) < 0.001)
				result.gridlines.push(round(item.args[0]));
			else if ("drawVerLine" === item.name)
				result.lines.push(round(item.args[1]));
		});
		result.gridlines.sort(function(a, b){return a - b;});
		result.lines.sort(function(a, b){return a - b;});
		return result;
	}

	// The x positions of all the cell edges of the first row
	function getCellEdges(table)
	{
		let row   = table.GetRow(0);
		let x     = table.Pages[0].X;
		let edges = [round(x + row.GetCellInfo(0).X_cell_start)];
		for (let cell = 0; cell < row.GetCellsCount(); ++cell)
			edges.push(round(x + row.GetCellInfo(cell).X_cell_end));
		return edges;
	}

	// "Convert text to table" of the paragraph, as Insert > Table > Convert text to table
	function convertTextToTable(docContent, paragraph, text)
	{
		AscTest.AddTextToParagraph(paragraph, text);
		AscTest.Recalculate();
		docContent.SelectAll();
		logicDocument.ConvertTextToTable(logicDocument.GetConvertTextToTableProps());
		AscTest.Recalculate();
		for (let i = 0; i < docContent.GetElementsCount(); ++i)
		{
			if (docContent.GetElement(i).IsTable())
				return docContent.GetElement(i);
		}
		return null;
	}

	// What the "No borders" button of the toolbar sends for the selected cells: borders with the value 0
	function applyNoBordersFromToolbar(table, onlyLeftCell)
	{
		if (onlyLeftCell)
			AscTest.SelectTableCells(table, 0, 0, 0, 0);
		else
			table.SelectAll();

		let borders = new Asc.CBorders();
		["Left", "Top", "Right", "Bottom", "InsideH", "InsideV"].forEach(function(side)
		{
			let border = new Asc.asc_CTextBorder();
			border.put_Color(new Asc.asc_CColor());
			border.put_Value(0);
			borders["put_" + side](border);
		});

		let props = new Asc.CTableProp();
		props.put_CellBorders(borders);
		props.put_CellSelect(true);
		table.Set_Props(props);
		table.RemoveSelection();
		AscTest.Recalculate();
	}

	function createFooter()
	{
		let footer = new AscCommonWord.CHeaderFooter(logicDocument.HdrFtr, logicDocument, logicDocument.DrawingDocument, AscCommon.hdrftr_Footer);
		logicDocument.GetFinalSectPr().Set_Footer_Default(footer);
		AscTest.Recalculate();
		logicDocument.SetDocPosType(AscCommonWord.docpostype_HdrFtr);
		logicDocument.HdrFtr.Set_CurHdrFtr(footer);
		footer.Set_Page(0);
		return footer;
	}

	function setup()
	{
		AscTest.ClearDocument();
		logicDocument.SetDocPosType(AscCommonWord.docpostype_Content);
		logicDocument.GetFinalSectPr().Set_Footer_Default(null);
		logicDocument.PushToContent(AscTest.CreateParagraph());
	}

	QUnit.module("Table gridlines");

	QUnit.test("Convert text to table gives a table like Insert table", function(assert)
	{
		setup();
		let table = convertTextToTable(logicDocument, logicDocument.GetElement(0), "IDnow\t06.10.2026 15:38:36\t6/7");
		assert.strictEqual(table.GetRow(0).GetCellsCount(), 3, "three cells");
		assert.strictEqual(table.Get_TableStyle(), logicDocument.GetStyles().Get_Default_TableGrid(), "the style Table Grid, as an inserted table");
		assert.deepEqual(getVerticalLines(table).lines, getCellEdges(table), "visible lines on all cell edges");
	});

	QUnit.test("Gridlines between the cells after No borders", function(assert)
	{
		setup();
		let table = convertTextToTable(logicDocument, logicDocument.GetElement(0), "IDnow\t06.10.2026 15:38:36\t6/7");
		applyNoBordersFromToolbar(table);
		let lines = getVerticalLines(table);
		assert.deepEqual(lines.lines, [], "no visible lines");
		assert.deepEqual(lines.gridlines, getCellEdges(table), "gridlines on the outer and the inner cell edges");

		setup();
		table = convertTextToTable(logicDocument, logicDocument.GetElement(0), "A\tB\tC");
		AscTest.RemoveTableBorders(table);
		AscTest.Recalculate();
		assert.deepEqual(getVerticalLines(table).gridlines, getCellEdges(table), "the same for table borders set to none");
	});

	QUnit.test("Gridlines between the cells of a table in a footer", function(assert)
	{
		setup();
		let footer = createFooter();
		let table  = convertTextToTable(footer.GetContent(), footer.GetContent().GetElement(0), "IDnow\t06.10.2026 15:38:36\t6/7");
		assert.ok(!!table, "the table is in the footer");
		applyNoBordersFromToolbar(table);
		let lines = getVerticalLines(table);
		assert.deepEqual(lines.lines, [], "no visible lines");
		assert.deepEqual(lines.gridlines, getCellEdges(table), "gridlines on the outer and the inner cell edges");
		logicDocument.SetDocPosType(AscCommonWord.docpostype_Content);
	});

	QUnit.test("Gridlines between the cells of a nested table", function(assert)
	{
		setup();
		let outer = convertTextToTable(logicDocument, logicDocument.GetElement(0), "X\tY");
		let cellContent = outer.GetRow(0).GetCell(0).GetContent();
		let table = AscTest.CreateTable(1, 3, [20, 20, 20]);
		cellContent.AddToContent(0, table);
		AscTest.Recalculate();
		applyNoBordersFromToolbar(table);
		assert.ok(table.GetPagesCount() > 0, "the nested table is calculated");
		assert.deepEqual(getVerticalLines(table).gridlines, getCellEdges(table), "gridlines on the outer and the inner cell edges");
	});

	QUnit.test("No borders for one cell", function(assert)
	{
		setup();
		let table = convertTextToTable(logicDocument, logicDocument.GetElement(0), "A\tB\tC");
		applyNoBordersFromToolbar(table, true);
		let edges = getCellEdges(table);
		let lines = getVerticalLines(table);
		assert.deepEqual(lines.lines, edges.slice(2), "visible lines around the other cells");
		assert.deepEqual(lines.gridlines, edges.slice(0, 2), "gridlines on both edges of the first cell");
		assert.strictEqual(table.GetRow(0).GetCell(0).Get_Borders().Left.GetWidth(), 0, "an invisible border has no width");
	});
});
