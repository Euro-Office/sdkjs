/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	// A 4×3 table with text in every cell
	function setup(autofit)
	{
		AscTest.ClearDocument();
		let table = AscTest.CreateTable(4, 3, [40, 40, 40]);
		for (let r = 0; r < 4; ++r)
		{
			for (let c = 0; c < 3; ++c)
				AscTest.AddTextToParagraph(table.GetRow(r).GetCell(c).GetContent().GetElement(0), "Row " + r + " col " + c);
		}
		table.SetTableLayout(autofit ? tbllayout_AutoFit : tbllayout_Fixed);
		logicDocument.PushToContent(table);
		AscTest.Recalculate();
		return table;
	}

	// Rows 2 and 3 as whole cells, like a selection with the mouse
	function selectRows(table)
	{
		table.SelectRange(0, 1, 2, 2);
		table.Document_SetThisElementCurrent(false);
		table.SelectRange(0, 1, 2, 2);
	}

	function round(v)
	{
		return Math.round(v * 100) / 100;
	}

	// Column widths, position and indent of the table
	function tableState(table)
	{
		return {
			grid : table.TableGridCalc.map(round),
			x    : round(table.Pages[0].X),
			ind  : round(table.Get_CompiledPr(false).TablePr.TableInd || 0)
		};
	}

	// The left indents of the paragraphs, row by row
	function indents(table)
	{
		let res = [];
		for (let r = 0; r < table.GetRowsCount(); ++r)
		{
			let row = [];
			for (let c = 0; c < 3; ++c)
				row.push(round(table.GetRow(r).GetCell(c).GetContent().GetElement(0).Get_CompiledPr2(false).ParaPr.Ind.Left));
			res.push(row);
		}
		return res;
	}

	function createRuler()
	{
		let ruler = new CHorRuler();
		ruler.m_oWordControl = {m_oLogicDocument : logicDocument, m_oApi : {}};
		return ruler;
	}

	QUnit.module("Indents in selected table cells");

	[false, true].forEach(function (autofit)
	{
		let layout = autofit ? " (autofit)" : " (fixed layout)";

		QUnit.test("Increase indent moves the text in the selected cells, not the table" + layout, function (assert)
		{
			let table  = setup(autofit);
			let before = tableState(table);
			selectRows(table);
			assert.true(table.IsCellSelection(), "rows 2 and 3 are selected as cells");

			logicDocument.IncreaseIndent();
			assert.deepEqual(indents(table), [[0, 0, 0], [12.5, 12.5, 12.5], [12.5, 12.5, 12.5], [0, 0, 0]], "the paragraphs of the selected cells are indented");
			assert.deepEqual(tableState(table), before, "column borders and table position stay");

			logicDocument.IncreaseIndent();
			assert.deepEqual(indents(table)[1], [25, 25, 25], "a second click indents one more step");

			logicDocument.DecreaseIndent();
			logicDocument.DecreaseIndent();
			assert.deepEqual(indents(table)[2], [0, 0, 0], "decrease indent takes the text back");
			assert.deepEqual(tableState(table), before, "the table still has not moved");

			logicDocument.Document_Undo();
			assert.deepEqual(indents(table)[1], [12.5, 12.5, 12.5], "one undo step per click");
		});
	});

	QUnit.test("Decrease indent does not move a table that has its own indent", function (assert)
	{
		let table = setup(false);
		table.Set_TableInd(10);
		AscTest.Recalculate();
		let before = tableState(table);
		selectRows(table);

		logicDocument.DecreaseIndent();
		assert.deepEqual(tableState(table), before, "the table indent stays 10 mm");
		assert.deepEqual(indents(table)[1], [0, 0, 0], "the text has no indent to decrease");
	});

	QUnit.test("Increase indent with the whole table selected", function (assert)
	{
		let table  = setup(false);
		let before = tableState(table);
		table.SelectAll();
		table.Document_SetThisElementCurrent(false);
		table.SelectAll();

		logicDocument.IncreaseIndent();
		assert.deepEqual(indents(table), [[12.5, 12.5, 12.5], [12.5, 12.5, 12.5], [12.5, 12.5, 12.5], [12.5, 12.5, 12.5]], "all paragraphs are indented");
		assert.deepEqual(tableState(table), before, "the table stays");
	});

	QUnit.test("Increase indent with the cursor in one cell works as before", function (assert)
	{
		let table  = setup(false);
		let before = tableState(table);
		AscTest.MoveCursorToParagraph(table.GetRow(1).GetCell(1).GetContent().GetElement(0));

		logicDocument.IncreaseIndent();
		assert.deepEqual(indents(table)[1], [0, 12.5, 0], "only this paragraph");
		assert.deepEqual(tableState(table), before, "the table stays");
	});

	[false, true].forEach(function (autofit)
	{
		let layout = autofit ? " (autofit)" : " (fixed layout)";

		QUnit.test("Ruler indent drag with selected rows" + layout, function (assert)
		{
			let table  = setup(autofit);
			let before = tableState(table);
			selectRows(table);

			// the left indent marker (with the first line) from 0 to 10 mm, with two preview steps
			let ruler = createRuler();
			ruler.m_dIndentRight       = 0;
			[3, 6, 10].forEach(function (left, i, steps)
			{
				ruler.m_dIndentLeft      = left;
				ruler.m_dIndentLeftFirst = left;
				ruler.SetPrProperties(i < steps.length - 1);
			});
			assert.deepEqual(indents(table), [[0, 0, 0], [10, 10, 10], [10, 10, 10], [0, 0, 0]], "the selected cells get the indent");
			assert.deepEqual(tableState(table), before, "column borders and table position stay");
		});

		QUnit.test("Indent fields of the Layout tab with selected rows" + layout, function (assert)
		{
			let table  = setup(autofit);
			let before = tableState(table);
			selectRows(table);

			// as api.paraApply with a left and a right indent
			logicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphPr);
			logicDocument.SetParagraphIndent({Left : 10});
			logicDocument.SetParagraphIndent({Right : 5});
			logicDocument.Recalculate();
			logicDocument.FinalizeAction();
			assert.deepEqual(indents(table)[2], [10, 10, 10], "the selected cells get the indent");
			assert.deepEqual(tableState(table), before, "column borders and table position stay");
		});
	});
});
