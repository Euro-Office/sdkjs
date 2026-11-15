/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

// w:compat/w:doNotUseHTMLParagraphAutoSpacing: the space after a paragraph and the space before the next one add up
// instead of collapsing to the larger of the two.
$(function ()
{
	const charWidth = AscTest.CharWidth * AscTest.FontSize;

	let logicDocument = AscTest.CreateLogicDocument();

	function initDocument()
	{
		AscTest.ClearDocument();

		let sectPr = AscTest.GetFinalSection();
		sectPr.SetPageSize(80 * charWidth, 1000);
		sectPr.SetPageMargins(10, 10, 10, 10);

		for (let i = 0; i < 2; ++i)
		{
			let p = AscTest.CreateParagraph();
			logicDocument.AddToContent(i, p);
			p.SetParagraphIndent({FirstLine : 0, Left : 0, Right : 0});
			AscTest.AddTextToParagraph(p, "Paragraph " + i);
		}
	}

	function getBottom(after, before)
	{
		let p1 = logicDocument.GetElement(0);
		let p2 = logicDocument.GetElement(1);
		p1.SetParagraphSpacing({Before : 0, After : after, LineRule : linerule_Auto, Line : 1});
		p2.SetParagraphSpacing({Before : before, After : 0, LineRule : linerule_Auto, Line : 1});
		AscTest.Recalculate();
		return p2.GetPageBounds(0).Bottom;
	}

	// The vertical space between the two paragraphs: how much lower the second one ends than without any spacing.
	function getGap(after, before)
	{
		return getBottom(after, before) - getBottom(0, 0);
	}

	QUnit.module("Paragraph spacing: doNotUseHTMLParagraphAutoSpacing", {
		beforeEach : function()
		{
			initDocument();
			AscTest.SetCompatibilityMode(AscCommon.document_compatibility_mode_Current);
		},
		afterEach : function()
		{
			logicDocument.GetDocumentSettings().DoNotUseHTMLParagraphAutoSpacing = false;
		}
	});

	QUnit.test("Without the setting the larger spacing counts", function(assert)
	{
		logicDocument.GetDocumentSettings().DoNotUseHTMLParagraphAutoSpacing = false;
		assert.close(getGap(20, 30), 30, 0.001, "space after 20, space before 30");
		assert.close(getGap(30, 20), 30, 0.001, "space after 30, space before 20");
	});

	QUnit.test("With the setting the spacings add up", function(assert)
	{
		logicDocument.GetDocumentSettings().DoNotUseHTMLParagraphAutoSpacing = true;
		assert.close(getGap(20, 30), 50, 0.001, "space after 20, space before 30");
		assert.close(getGap(30, 20), 50, 0.001, "space after 30, space before 20");
		assert.close(getGap(0, 15), 15, 0.001, "no space after, space before 15");
	});
});
