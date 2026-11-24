/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	function bottom()
	{
		return {Bottom : {Value : border_Single, Size : 0.5, Space : 0, Color : {r : 0, g : 0, b : 0}}};
	}

	function hasBottom(paragraph)
	{
		return !paragraph.Get_CompiledPr2(false).ParaPr.Brd.Bottom.IsNone();
	}

	function setup(withExistingGroup)
	{
		AscTest.ClearDocument();
		let paras = [];
		["Beilage", "", "FIRMA | Erika Beispiel", "", "", ""].forEach(function (text)
		{
			let p = AscTest.CreateParagraph();
			text && AscTest.AddTextToParagraph(p, text);
			logicDocument.PushToContent(p);
			paras.push(p);
		});
		logicDocument.PushToContent(AscTest.CreateTable(2, 2));
		if (withExistingGroup)
		{
			paras[2].Set_Borders(bottom());
			paras[3].Set_Borders(bottom());
		}
		AscTest.Recalculate();
		return paras;
	}

	QUnit.module("Paragraph borders");

	QUnit.test("A border on a paragraph without borders goes on that paragraph only", function (assert)
	{
		let paras = setup(false);
		AscTest.MoveCursorToParagraph(paras[2], false);
		logicDocument.SetParagraphBorders(bottom());
		assert.deepEqual(paras.map(hasBottom), [false, false, true, false, false, false], "only the text line has the bottom border, not the empty lines up to the table");
	});

	QUnit.test("Paragraphs that already share a border still change together", function (assert)
	{
		let paras = setup(true);
		AscTest.MoveCursorToParagraph(paras[2], false);
		logicDocument.SetParagraphBorders({Bottom : {Value : border_None}});
		assert.deepEqual(paras.map(hasBottom), [false, false, false, false, false, false], "the whole bordered group loses the border, as in Word");
	});

	QUnit.test("Inside a table cell too", function (assert)
	{
		setup(false);
		let table = logicDocument.GetElement(6);
		let content = table.GetRow(0).GetCell(0).GetContent();
		AscTest.AddTextToParagraph(content.GetElement(0), "Text");
		let p2 = AscTest.CreateParagraph();
		content.AddToContent(1, p2);
		AscTest.Recalculate();
		AscTest.MoveCursorToParagraph(content.GetElement(0), false);
		logicDocument.SetParagraphBorders(bottom());
		assert.true(hasBottom(content.GetElement(0)), "the line in the cell has the border");
		assert.false(hasBottom(p2), "the empty line below it has none");
	});
});
