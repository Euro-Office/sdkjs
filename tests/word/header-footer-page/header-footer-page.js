/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	// Calculate all pages at once, without the timer
	window["NATIVE_EDITOR_ENJINE_SYNC_RECALC"] = true;
	let logicDocument = AscTest.CreateLogicDocument();

	// Six pages, one section, one header and one footer for all pages
	function setup()
	{
		AscTest.ClearDocument();
		let sectPr = logicDocument.GetFinalSectPr();

		let headerContent = AscTest.CreateDefaultHeader(sectPr);
		AscTest.AddTextToParagraph(headerContent.GetElement(0), "Header");

		let footer = new AscCommonWord.CHeaderFooter(logicDocument.HdrFtr, logicDocument, logicDocument.DrawingDocument, AscCommon.hdrftr_Footer);
		sectPr.Set_Footer_Default(footer);
		AscTest.AddTextToParagraph(footer.GetContent().GetElement(0), "Footer");

		for (let i = 0; i < 6; ++i)
		{
			let p = AscTest.CreateParagraph();
			AscTest.AddTextToParagraph(p, "Page " + (i + 1));
			if (i > 0)
				p.Set_PageBreakBefore(true);
			logicDocument.PushToContent(p);
		}
		AscTest.Recalculate();
		return {header : headerContent.Parent, footer : footer};
	}

	function editOnPage(hdrFtr, pageAbs)
	{
		logicDocument.RemoveSelection();
		hdrFtr.Set_CurrentElement(true, pageAbs);
		hdrFtr.GetContent().MoveCursorToEndPos(false);
	}

	QUnit.module("Header and footer stay on their page");

	QUnit.test("The document has six pages with the same header", function (assert)
	{
		let hf = setup();
		assert.strictEqual(logicDocument.GetPagesCount(), 6, "six pages");
		editOnPage(hf.header, 4);
		assert.strictEqual(hf.header.GetPage(), 4, "header of page 5 is current");
	});

	QUnit.test("Typing in the header of page 5 keeps page 5", function (assert)
	{
		let hf = setup();
		editOnPage(hf.header, 4);
		AscTest.EnterText("x");
		assert.strictEqual(hf.header.GetContent().GetElement(0).GetText({ParaSeparator : ""}), "Headerx", "text added");
		assert.strictEqual(hf.header.GetPage(), 4, "header still on page 5");
		assert.strictEqual(logicDocument.HdrFtr.Get_CurPage(), 4, "current page of the document is page 5");
		assert.strictEqual(hf.header.GetContent().GetAbsoluteStartPage(), 4, "cursor computed on page 5");
	});

	QUnit.test("Typing in the footer of page 3 keeps page 3", function (assert)
	{
		let hf = setup();
		editOnPage(hf.footer, 2);
		AscTest.EnterText("ab");
		assert.strictEqual(hf.footer.GetPage(), 2, "footer still on page 3");
		assert.strictEqual(hf.footer.GetContent().GetAbsoluteStartPage(), 2, "cursor computed on page 3");
	});

	QUnit.test("A new line in the header of the last page keeps the last page", function (assert)
	{
		let hf = setup();
		editOnPage(hf.header, 5);
		AscTest.PressKey(AscTest.Key.enter);
		AscTest.EnterText("y");
		assert.strictEqual(hf.header.GetContent().GetElementsCount(), 2, "two paragraphs");
		assert.strictEqual(hf.header.GetPage(), 5, "header still on page 6");
	});

	QUnit.test("Undo keeps the page", function (assert)
	{
		let hf = setup();
		editOnPage(hf.header, 3);
		AscTest.EnterText("z");
		logicDocument.Document_Undo();
		assert.strictEqual(hf.header.GetContent().GetElement(0).GetText({ParaSeparator : ""}), "Header", "undone");
		assert.strictEqual(hf.header.GetPage(), 3, "header still on page 4");
	});
});
