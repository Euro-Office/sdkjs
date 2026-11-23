/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	// A table whose column widths the user has just changed (its own undo step), the cursor in the right cell. The
	// text there has its own indents and a tab stop, as in most documents.
	function setup()
	{
		AscTest.ClearDocument();
		let table = AscTest.CreateTable(1, 2, [50, 50]);
		logicDocument.PushToContent(table);

		let para = table.GetRow(0).GetCell(1).GetContent().GetElement(0);
		para.Set_Ind({Left : 0, Right : 0, FirstLine : 0}, false);
		let tabs = new CParaTabs();
		tabs.Add(new CParaTab(tab_Left, 10));
		para.Set_Tabs(tabs);
		AscTest.AddTextToParagraph(para, "Text");
		AscTest.Recalculate();

		logicDocument.StartAction(AscDFH.historydescription_Document_SetTableMarkup_Hor);
		table.SetTableGrid([70, 30]);
		logicDocument.Recalculate();
		logicDocument.FinalizeAction();

		AscTest.MoveCursorToParagraph(para);
		return {table : table, para : para};
	}

	function createRuler()
	{
		let ruler = new CHorRuler();
		ruler.m_oWordControl = {m_oLogicDocument : logicDocument, m_oApi : {}};
		return ruler;
	}

	// Drags the left indent marker: every value but the last is a preview step, the last one is the drop
	function dragLeftIndent(ruler, steps)
	{
		ruler.m_dIndentRight = 0;
		steps.forEach(function (left, i)
		{
			ruler.m_dIndentLeft      = left;
			ruler.m_dIndentLeftFirst = left;
			ruler.SetPrProperties(i < steps.length - 1);
		});
	}

	function grid(table)
	{
		return table.TableGrid.slice();
	}

	function leftInd(para)
	{
		return para.Get_CompiledPr2(false).ParaPr.Ind.Left;
	}

	QUnit.module("Ruler preview while dragging");

	QUnit.test("Preview steps replace each other, one undo step at the end", function (assert)
	{
		let doc = setup();
		dragLeftIndent(createRuler(), [2, 3, 4]);
		assert.strictEqual(leftInd(doc.para), 4, "the indent is set");
		assert.deepEqual(grid(doc.table), [70, 30], "the column widths stay");
		logicDocument.Document_Undo();
		assert.strictEqual(leftInd(doc.para), 0, "one undo removes the indent");
		assert.deepEqual(grid(doc.table), [70, 30], "and keeps the column widths");
	});

	QUnit.test("A first preview step without a change does not undo the step before", function (assert)
	{
		let doc = setup();
		// the mouse has not moved far enough yet: the first step still has the old indent
		dragLeftIndent(createRuler(), [0, 0, 4, 5]);
		assert.deepEqual(grid(doc.table), [70, 30], "the new column widths stay");
		assert.strictEqual(leftInd(doc.para), 5, "the indent is set");

		logicDocument.Document_Undo();
		assert.strictEqual(leftInd(doc.para), 0, "undo removes the indent");
		assert.deepEqual(grid(doc.table), [70, 30], "and keeps the column widths");
		logicDocument.Document_Undo();
		assert.deepEqual(grid(doc.table), [50, 50], "the next undo restores the column widths");
	});

	QUnit.test("A marker dragged back to its old place does not undo the step before", function (assert)
	{
		let doc = setup();
		dragLeftIndent(createRuler(), [3, 0, 0]);
		assert.deepEqual(grid(doc.table), [70, 30], "the new column widths stay");

		// the next action of the user
		logicDocument.StartAction(AscDFH.historydescription_Document_AddLetter);
		logicDocument.AddToParagraph(new AscWord.CRunText(0x41));
		logicDocument.FinalizeAction();
		assert.deepEqual(grid(doc.table), [70, 30], "still there after the next action");
	});

	QUnit.test("Tab stop preview does not undo the step before", function (assert)
	{
		let doc = setup();
		let ruler = createRuler();
		ruler.SetTabsProperties(true, [{type : tab_Left, pos : 10, leader : undefined}]);
		ruler.SetTabsProperties(true, [{type : tab_Left, pos : 15, leader : undefined}]);
		assert.deepEqual(grid(doc.table), [70, 30], "the new column widths stay");
	});
});
