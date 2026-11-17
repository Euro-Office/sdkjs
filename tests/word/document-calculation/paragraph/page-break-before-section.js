/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

// w:pageBreakBefore on the first paragraph of a continuous section. A continuous section starts on a new page by
// itself when the page before has footnotes; the paragraph must not break the page a second time then.
$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	// Three paragraphs: the second one ends the first section, the third one is the first of the final section.
	function initDocument(withFootnote, pageBreakBefore, sectionType)
	{
		AscTest.ClearDocument();

		let finalSectPr = AscTest.GetFinalSection();
		finalSectPr.SetPageSize(400, 400);
		finalSectPr.SetPageMargins(50, 50, 50, 50);
		finalSectPr.SetType(sectionType);

		let paragraphs = [];
		for (let i = 0; i < 3; ++i)
		{
			let p = AscTest.CreateParagraph();
			logicDocument.AddToContent(i, p);
			AscTest.AddTextToParagraph(p, "Paragraph " + i);
			paragraphs.push(p);
		}

		let sectPr = new AscWord.SectPr(logicDocument);
		sectPr.SetPageSize(400, 400);
		sectPr.SetPageMargins(50, 50, 50, 50);
		paragraphs[1].SetSectionPr(sectPr);
		paragraphs[2].Set_PageBreakBefore(pageBreakBefore);

		if (withFootnote)
		{
			logicDocument.SetDocPosType(docpostype_Content);
			AscTest.MoveCursorToParagraph(paragraphs[0], false);
			let footnote = logicDocument.AddFootnote();
			AscTest.AddTextToParagraph(footnote.GetElement(0), "Footnote");
			logicDocument.SetDocPosType(docpostype_Content);
		}

		AscTest.Recalculate();
		return paragraphs;
	}

	// A footnote finds its document only in the document editor
	let isDocumentEditor;

	QUnit.module("Page break before the first paragraph of a section", {
		beforeEach : function()
		{
			isDocumentEditor = editor.isDocumentEditor;
			editor.isDocumentEditor = true;
		},
		afterEach : function()
		{
			AscTest.ClearDocument();
			AscTest.GetFinalSection().SetType(Asc.c_oAscSectionBreakType.NextPage);
			editor.isDocumentEditor = isDocumentEditor;
		}
	});

	QUnit.test("Continuous section after a page with footnotes", function (assert)
	{
		let p = initDocument(true, true, Asc.c_oAscSectionBreakType.Continuous);
		assert.strictEqual(logicDocument.Footnotes.IsEmptyPage(0), false, "The first page has a footnote");
		assert.strictEqual(p[1].GetAbsoluteStartPage(), 0, "The first section ends on the first page");
		assert.strictEqual(p[2].GetPagesCount(), 1, "The paragraph has no empty page of its own");
		assert.strictEqual(p[2].GetAbsoluteStartPage(), 1, "The paragraph starts on the second page");
		assert.strictEqual(logicDocument.GetPagesCount(), 2, "No empty page between the sections");
	});

	QUnit.test("Continuous section after a page without footnotes", function (assert)
	{
		let p = initDocument(false, true, Asc.c_oAscSectionBreakType.Continuous);
		assert.strictEqual(logicDocument.Footnotes.IsEmptyPage(0), true, "The first page has no footnote");
		assert.strictEqual(p[2].GetPagesCount(), 2, "The paragraph breaks the page");
		assert.strictEqual(p[2].IsEmptyPage(0) && !p[2].IsEmptyPage(1), true, "Its text is on its second page");
		assert.strictEqual(p[2].GetAbsoluteStartPage(), 0, "The section starts on the first page");
		assert.strictEqual(logicDocument.GetPagesCount(), 2, "The page break gives the second page");
	});

	QUnit.test("Continuous section after a page with footnotes, no page break before", function (assert)
	{
		let p = initDocument(true, false, Asc.c_oAscSectionBreakType.Continuous);
		assert.strictEqual(p[2].GetPagesCount(), 1, "The paragraph has one page");
		assert.strictEqual(p[2].GetAbsoluteStartPage(), 1, "The section starts on the second page");
		assert.strictEqual(logicDocument.GetPagesCount(), 2, "Two pages");
	});

	QUnit.test("Next page section after a page with footnotes", function (assert)
	{
		let p = initDocument(true, true, Asc.c_oAscSectionBreakType.NextPage);
		assert.strictEqual(p[2].GetPagesCount(), 1, "The paragraph has no empty page of its own");
		assert.strictEqual(p[2].GetAbsoluteStartPage(), 1, "The paragraph starts on the second page");
		assert.strictEqual(logicDocument.GetPagesCount(), 2, "Two pages");
	});
});
