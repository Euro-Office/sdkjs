/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	window["NATIVE_EDITOR_ENJINE_SYNC_RECALC"] = true;
	let logicDocument = AscTest.CreateLogicDocument();
	// The table properties ask the drawing document to update the style previews; the test has no previews
	if (!AscTest.DrawingDocument.CheckTableStyles)
		AscTest.DrawingDocument.CheckTableStyles = function(){};

	const GRID = [40, 40, 40];

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

	// The free lines (_m to _l) the table draws on the page, as [x0, y0, x1, y1]
	function getDiagonals(table, page)
	{
		let log = [];
		table.Draw(page || 0, recordingGraphics(log));
		let lines = [];
		for (let i = 0; i < log.length; ++i)
		{
			if ("_m" === log[i].name && log[i + 1] && "_l" === log[i + 1].name)
			{
				lines.push([round(log[i].args[0]), round(log[i].args[1]), round(log[i + 1].args[0]), round(log[i + 1].args[1])]);
			}
		}
		return lines;
	}

	// The corners of the cell (merged cells included) on the first page
	function getCellRect(table, rowIndex, cellIndex, rowsCount)
	{
		let row      = table.GetRow(rowIndex);
		let info     = row.GetCellInfo(cellIndex);
		let x        = table.Pages[0].X;
		let lastRow  = table.RowsInfo[rowIndex + (rowsCount || 1) - 1];
		return {
			x0 : round(x + info.X_cell_start),
			x1 : round(x + info.X_cell_end),
			y0 : round(table.RowsInfo[rowIndex].Y[0]),
			y1 : round(lastRow.Y[0] + lastRow.H[0])
		};
	}

	function createTable(rowsCount)
	{
		AscTest.ClearDocument();
		let table = AscTest.CreateTable(rowsCount, GRID.length, GRID.slice());
		logicDocument.PushToContent(table);
		for (let r = 0; r < rowsCount; ++r)
		{
			for (let c = 0; c < GRID.length; ++c)
				AscTest.AddTextToParagraph(table.GetRow(r).GetCell(c).GetContent().GetElement(0), "R" + r + "C" + c);
		}
		AscTest.Recalculate();
		return table;
	}

	function makeBorder(r, g, b, sizePt)
	{
		let border = new Asc.asc_CTextBorder();
		border.put_Value(1);
		border.put_Size(sizePt * g_dKoef_pt_to_mm);
		border.put_Color(AscCommon.CreateAscColorCustom(r, g, b));
		return border;
	}

	function noBorder()
	{
		let border = new Asc.asc_CTextBorder();
		border.put_Value(0);
		return border;
	}

	// Applies diagonals as the table settings panel does (only the diagonals are given)
	function applyDiagonals(down, up)
	{
		let props   = new Asc.CTableProp();
		let borders = new Asc.CBorders();
		if (down)
			borders.put_DiagonalDown(down);
		if (up)
			borders.put_DiagonalUp(up);
		props.put_CellBorders(borders);
		props.put_CellSelect(true);
		// the same call as the panel: asc_docs_api.tblApply on the test document
		AscTest.Editor.WordControl.m_oLogicDocument = logicDocument;
		window["Asc"]["asc_docs_api"].prototype.tblApply.call(AscTest.Editor, props);
		AscTest.Recalculate();
	}

	function selectCell(table, row, cell)
	{
		AscTest.SelectTableCells(table, cell, row, cell, row);
	}

	QUnit.module("Diagonal borders in table cells");

	QUnit.test("The cell properties keep both diagonals (copy, compare, binary for undo and co-editing)", function (assert)
	{
		let pr = new CTableCellPr();
		assert.strictEqual(pr.TableCellBorders.TL2BR, undefined, "A new cell has no diagonal from top left");
		assert.strictEqual(pr.TableCellBorders.TR2BL, undefined, "A new cell has no diagonal from top right");
		assert.ok(pr.Is_Empty(), "Empty without diagonals");

		let down = new CDocumentBorder();
		down.Value = border_Single;
		down.Size  = 1;
		down.Color = new CDocumentColor(255, 0, 0);
		pr.TableCellBorders.TL2BR = down;
		assert.notOk(pr.Is_Empty(), "Not empty with a diagonal");

		let copy = pr.Copy();
		assert.ok(copy.TableCellBorders.TL2BR && copy.TableCellBorders.TL2BR.IsEqual(down), "Copy keeps the diagonal");
		assert.ok(copy.Is_Equal(pr), "The copy is equal");

		let other = pr.Copy();
		other.TableCellBorders.TR2BL = down.Copy();
		assert.notOk(other.Is_Equal(pr), "A second diagonal makes the properties different");

		let writer = AscTest.GetBinaryWriter();
		other.Write_ToBinary(writer);
		let read = new CTableCellPr();
		read.Read_FromBinary(AscTest.GetBinaryReader(writer));
		assert.ok(read.Is_Equal(other), "Binary write and read keep both diagonals");
	});

	QUnit.test("Setting a diagonal in the panel draws one line from corner to corner; no diagonal, no line", function (assert)
	{
		let table = createTable(2);
		assert.deepEqual(getDiagonals(table), [], "Without diagonals the table draws no free lines");

		selectCell(table, 0, 1);
		applyDiagonals(makeBorder(255, 0, 0, 1), null);

		let cell = table.GetRow(0).GetCell(1);
		let diag = cell.GetDiagonalBorder(false);
		assert.ok(diag, "The cell has a diagonal from top left");
		assert.strictEqual(diag && diag.Color.r, 255, "with the chosen colour");
		assert.ok(diag && Math.abs(diag.Size - g_dKoef_pt_to_mm) < 0.001, "and the chosen width");
		assert.strictEqual(cell.GetDiagonalBorder(true), undefined, "No diagonal from top right");
		assert.strictEqual(table.GetRow(0).GetCell(0).GetDiagonalBorder(false), undefined, "The other cells have none");

		let rect = getCellRect(table, 0, 1);
		assert.deepEqual(getDiagonals(table), [[rect.x0, rect.y0, rect.x1, rect.y1]], "One line from the top left to the bottom right corner");

		applyDiagonals(null, makeBorder(0, 0, 255, 0.5));
		let lines = getDiagonals(table);
		assert.strictEqual(lines.length, 2, "Both diagonals are drawn");
		assert.deepEqual(lines[1], [rect.x1, rect.y0, rect.x0, rect.y1], "The second goes from the top right to the bottom left corner");

		let tableProps = new Asc.CTableProp(table.Get_Props());
		assert.strictEqual(tableProps.get_CellBorders().get_DiagonalDown().get_Value(), 1, "The panel sees the diagonal from top left");
		assert.strictEqual(tableProps.get_CellBorders().get_DiagonalUp().get_Value(), 1, "The panel sees the diagonal from top right");

		applyDiagonals(noBorder(), noBorder());
		assert.strictEqual(cell.GetDiagonalBorder(false), undefined, "No diagonal removes the first");
		assert.strictEqual(cell.GetDiagonalBorder(true), undefined, "No diagonal removes the second");
		assert.deepEqual(getDiagonals(table), [], "Nothing is drawn any more");
	});

	QUnit.test("Undo and redo, and the change for co-editing", function (assert)
	{
		let table = createTable(2);
		let cell  = table.GetRow(1).GetCell(0);
		selectCell(table, 1, 0);

		AscCommon.History.Clear();
		applyDiagonals(makeBorder(0, 128, 0, 1.5), null);
		assert.ok(cell.GetDiagonalBorder(false), "Set");

		let change = null;
		AscCommon.History.Points[AscCommon.History.Index].Items.forEach(function(item)
		{
			if (AscDFH.historyitem_TableCell_Border_TL2BR === item.Data.Type)
				change = item.Data;
		});
		assert.ok(change, "The change is recorded in the history");
		if (change)
		{
			let writer = AscTest.GetBinaryWriter();
			change.WriteToBinary(writer);
			let read = new AscDFH.changesFactory[change.Type](cell);
			read.ReadFromBinary(AscTest.GetBinaryReader(writer));
			assert.ok(read.New && border_Single === read.New.Value && 128 === read.New.Color.g
				&& Math.abs(read.New.Size - change.New.Size) < 0.001, "The co-editing change carries the new diagonal");
			assert.strictEqual(read.Old, undefined, "and that there was none before");
		}

		logicDocument.Document_Undo();
		AscTest.Recalculate();
		assert.strictEqual(cell.GetDiagonalBorder(false), undefined, "Undo removes the diagonal");
		assert.deepEqual(getDiagonals(table), [], "and nothing is drawn");

		logicDocument.Document_Redo();
		AscTest.Recalculate();
		assert.ok(cell.GetDiagonalBorder(false), "Redo brings it back");
		assert.strictEqual(getDiagonals(table).length, 1, "and it is drawn again");
	});

	QUnit.test("Merged cells get one line from corner to corner of the merged cell", function (assert)
	{
		let table = createTable(3);
		AscTest.MergeTableCells(table, 0, 0, 1, 1);
		AscTest.Recalculate();

		// After merging, the merged cell is the first cell of the first row and goes down over two rows
		selectCell(table, 0, 0);
		applyDiagonals(makeBorder(0, 0, 0, 0.5), makeBorder(0, 0, 0, 0.5));

		let rect  = getCellRect(table, 0, 0, 2);
		let lines = getDiagonals(table);
		assert.strictEqual(lines.length, 2, "Two lines for the merged cell, not one per row or column");
		assert.deepEqual(lines[0], [rect.x0, rect.y0, rect.x1, rect.y1], "From the top left of the merged cell to its bottom right");
		assert.deepEqual(lines[1], [rect.x1, rect.y0, rect.x0, rect.y1], "From its top right to its bottom left");
		assert.ok(rect.x1 - rect.x0 > GRID[0] + GRID[1] - 1, "The line goes across both columns");

		// Selecting the lower part of the merged cell still changes the merged cell
		let lowerRow = table.GetRow(1);
		assert.strictEqual(lowerRow.GetCell(0).GetVMerge(), vmerge_Continue, "The second row continues the merged cell");
		selectCell(table, 1, 0);
		applyDiagonals(noBorder(), null);
		assert.strictEqual(table.GetRow(0).GetCell(0).GetDiagonalBorder(false), undefined, "The diagonal of the merged cell is removed");
		assert.strictEqual(lowerRow.GetCell(0).GetDiagonalBorder(false), undefined, "The hidden part never gets one");
	});

	QUnit.test("Several selected cells: each gets its own diagonal; different cells show as mixed", function (assert)
	{
		let table = createTable(2);
		AscTest.SelectTableCells(table, 0, 0, 1, 0);
		applyDiagonals(makeBorder(0, 0, 0, 0.5), null);
		assert.ok(table.GetRow(0).GetCell(0).GetDiagonalBorder(false), "first cell");
		assert.ok(table.GetRow(0).GetCell(1).GetDiagonalBorder(false), "second cell");
		assert.strictEqual(getDiagonals(table).length, 2, "two lines, one per cell (Word has no diagonal across several cells)");

		AscTest.SelectTableCells(table, 1, 0, 2, 0);
		let mixed = new Asc.CTableProp(table.Get_Props());
		assert.strictEqual(mixed.get_CellBorders().get_DiagonalDown(), null, "A cell with and a cell without: mixed");
	});

	QUnit.test("The file keeps the diagonals (save and open in the editor's format)", function (assert)
	{
		let table = createTable(1);
		selectCell(table, 0, 2);
		applyDiagonals(makeBorder(255, 0, 0, 1), makeBorder(0, 0, 255, 0.5));

		// The test document has a forms part, which needs the zip engine of the app; the table does not
		let getOForm = logicDocument.GetOFormDocument;
		logicDocument.GetOFormDocument = function(){return null;};
		// and the writer asks the app for macros
		let api = AscTest.DrawingDocument.m_oWordControl ? AscTest.DrawingDocument.m_oWordControl.m_oApi : null;
		if (api && !api.macros)
			api.macros = {GetData : function(){return null;}};
		let data = new AscCommonWord.BinaryFileWriter(logicDocument).Write();
		logicDocument.GetOFormDocument = getOForm;
		assert.ok(data, "The document is written");

		let reread = AscFormat.ExecuteNoHistory(function()
		{
			let doc    = new CDocument(AscTest.DrawingDocument, false);
			let reader = new AscCommonWord.BinaryFileReader(doc, {disableRevisions : true, noSendComments : true, noGenerateSmartArts : true});
			return reader.Read(data) ? doc : null;
		}, this, []);
		assert.ok(reread, "and read again");

		let readTable = null;
		for (let i = 0; reread && i < reread.Content.length; ++i)
		{
			if (reread.Content[i].IsTable())
				readTable = reread.Content[i];
		}
		assert.ok(readTable, "The table is in the file");
		if (!readTable)
			return;

		let cellPr = readTable.GetRow(0).GetCell(2).Pr;
		let down   = cellPr.TableCellBorders.TL2BR;
		let up     = cellPr.TableCellBorders.TR2BL;
		assert.ok(down && border_Single === down.Value, "The diagonal from top left is in the file");
		assert.strictEqual(down && down.Color.r, 255, "with its colour");
		assert.ok(up && border_Single === up.Value, "The diagonal from top right is in the file");
		assert.strictEqual(up && up.Color.b, 255, "with its colour");
		assert.strictEqual(readTable.GetRow(0).GetCell(0).Pr.TableCellBorders.TL2BR, undefined, "The other cells have none");
	});
});
