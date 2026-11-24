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

	QUnit.test("A second change before the recalculation reaches the page keeps the page", function (assert)
	{
		// In the editor the recalculation runs in timer steps. The first step calculates the first pages and makes the
		// first page that uses this header the current one; a change that comes before the step that reaches the
		// edited page must not take that first page as the edited one.
		let hf = setup();
		editOnPage(hf.header, 4);
		let header = hf.header;

		header.Refresh_RecalcData2();
		assert.strictEqual(header.RecalcInfo.RequestedPage, 4, "page 5 waits after the first change");

		// the first timer step: page 1 is calculated and becomes the current page of the header
		header.Recalculate(0, logicDocument.GetFinalSectPr());
		assert.strictEqual(header.GetPage(), 0, "page 1 is current for the moment");

		// the second change of the same action (or the next key) comes now
		header.Refresh_RecalcData2();
		assert.strictEqual(header.RecalcInfo.RequestedPage, 4, "page 5 still waits");

		AscTest.Recalculate();
		assert.strictEqual(header.GetPage(), 4, "header on page 5 after the recalculation");
		assert.strictEqual(header.RecalcInfo.RequestedPage, -1, "nothing waits any more");
	});

	// While the recalculation runs in timer steps, the document only has the pages up to the step that is being
	// calculated. These helpers cut the pages off and put them back.
	function withFewerPages(count, callback)
	{
		let pages = logicDocument.Pages;
		logicDocument.Pages = pages.slice(0, count);
		try
		{
			callback();
		}
		finally
		{
			logicDocument.Pages = pages;
		}
	}

	QUnit.test("Measuring the footer on an earlier page keeps the edited page", function (assert)
	{
		// The recalculation of page 2 asks for the top of the footer on page 2; page 3, where the footer is edited,
		// is not there yet at that moment
		let hf = setup();
		editOnPage(hf.footer, 2);
		let expected = hf.footer.RecalcInfo.RecalcObj[1].Get_PageBounds(0).Top;

		let top;
		withFewerPages(2, function()
		{
			top = hf.footer.Get_DividingLine(1);
		});
		assert.strictEqual(top, expected, "top of the footer on page 2");
		assert.strictEqual(hf.footer.GetPage(), 2, "footer still on page 3");

		AscTest.Recalculate();
		assert.strictEqual(hf.footer.GetPage(), 2, "footer on page 3 after the recalculation");
	});

	QUnit.test("Measuring the header on its own page gives its bottom", function (assert)
	{
		let hf = setup();
		editOnPage(hf.header, 1);
		assert.strictEqual(hf.header.Get_DividingLine(1), hf.header.GetContent().Get_PageBounds(0).Bottom, "bottom of the header on page 2");
		assert.strictEqual(hf.header.Get_DividingLine(3), hf.header.RecalcInfo.RecalcObj[3].Get_PageBounds(0).Bottom, "bottom of the header on page 4");
		assert.strictEqual(hf.header.GetPage(), 1, "header still on page 2");
	});

	QUnit.test("A change while the later pages are not there yet keeps the page", function (assert)
	{
		// Fast typing: the next key comes while the recalculation has only reached page 2
		let hf = setup();
		editOnPage(hf.header, 4);
		let header = hf.header;

		header.Refresh_RecalcData2();
		header.Recalculate(0, logicDocument.GetFinalSectPr());
		withFewerPages(2, function()
		{
			header.Refresh_RecalcData2();
		});
		assert.strictEqual(header.RecalcInfo.RequestedPage, 4, "page 5 still waits");

		AscTest.Recalculate();
		assert.strictEqual(header.GetPage(), 4, "header on page 5 after the recalculation");
	});

	QUnit.test("Undo goes back to the page that is waiting", function (assert)
	{
		// The selection state for Undo is taken while the recalculation has only reached page 1
		let hf = setup();
		editOnPage(hf.header, 4);
		let header = hf.header;

		header.Refresh_RecalcData2();
		header.Recalculate(0, logicDocument.GetFinalSectPr());
		assert.strictEqual(header.GetPage(), 0, "page 1 is current for the moment");

		let state = logicDocument.HdrFtr.GetSelectionState();
		assert.strictEqual(state[state.length - 1].CurPage, 4, "the state remembers page 5");

		AscTest.Recalculate();
		editOnPage(header, 4);
		AscTest.EnterText("a");
		AscTest.EnterText("b");
		logicDocument.Document_Undo();
		assert.strictEqual(header.GetPage(), 4, "header on page 5 after Undo");
	});

	QUnit.test("Every page keeps its header and footer after editing on page 4", function (assert)
	{
		let hf = setup();
		editOnPage(hf.footer, 3);
		AscTest.EnterText("x");
		AscTest.PressKey(AscTest.Key.enter);
		AscTest.EnterText("y");
		for (let page = 0; page < 6; ++page)
		{
			let pageHdrFtr = logicDocument.HdrFtr.Pages[page];
			assert.ok(pageHdrFtr && pageHdrFtr.Header === hf.header && hf.header.RecalcInfo.RecalcObj[page], "header calculated on page " + (page + 1));
			assert.ok(pageHdrFtr && pageHdrFtr.Footer === hf.footer && hf.footer.RecalcInfo.RecalcObj[page], "footer calculated on page " + (page + 1));
		}
		assert.strictEqual(hf.footer.GetPage(), 3, "footer still on page 4");
	});

	// Three sections with three pages each, every section with its own header and footer
	function setupSections(titlePage, evenAndOdd)
	{
		AscTest.ClearDocument();
		EvenAndOddHeaders = !!evenAndOdd;
		// ClearDocument can leave one empty paragraph
		let emptyCount = logicDocument.GetElementsCount();
		let result = {sectPr : [], header : [], footer : []};
		for (let section = 0; section < 3; ++section)
		{
			let sectPr = 2 === section ? logicDocument.GetFinalSectPr() : new AscWord.SectPr(logicDocument);
			for (let i = 0; i < 3; ++i)
			{
				let p = AscTest.CreateParagraph();
				AscTest.AddTextToParagraph(p, "Page " + (section * 3 + i + 1));
				if (i > 0)
					p.Set_PageBreakBefore(true);
				logicDocument.PushToContent(p);
				if (2 === i && 2 !== section)
					p.SetSectionPr(sectPr);
			}
			sectPr.Set_TitlePage(!!titlePage);

			let header = new AscCommonWord.CHeaderFooter(logicDocument.HdrFtr, logicDocument, logicDocument.DrawingDocument, AscCommon.hdrftr_Header);
			AscTest.AddTextToParagraph(header.GetContent().GetElement(0), "Header " + (section + 1));
			let footer = new AscCommonWord.CHeaderFooter(logicDocument.HdrFtr, logicDocument, logicDocument.DrawingDocument, AscCommon.hdrftr_Footer);
			AscTest.AddTextToParagraph(footer.GetContent().GetElement(0), "Footer " + (section + 1));
			if (titlePage)
			{
				sectPr.Set_Header_First(header);
				sectPr.Set_Footer_First(footer);
			}
			else if (evenAndOdd)
			{
				sectPr.Set_Header_Even(header);
				sectPr.Set_Footer_Even(footer);
			}
			else
			{
				sectPr.Set_Header_Default(header);
				sectPr.Set_Footer_Default(footer);
			}
			result.sectPr.push(sectPr);
			result.header.push(header);
			result.footer.push(footer);
		}
		if (emptyCount > 0)
			logicDocument.RemoveFromContent(0, emptyCount);
		AscTest.Recalculate();
		return result;
	}

	// The same calls as asc_docs_api.HeadersAndFooters_LinkToPrevious (toolbar and header/footer settings)
	function setLinkToPrevious(isOn)
	{
		if (logicDocument.Document_Is_SelectionLocked(AscCommon.changestype_HdrFtr))
			return;

		logicDocument.StartAction(AscDFH.historydescription_Document_SetHdrFtrLink);
		logicDocument.Document_SetHdrFtrLink(isOn);
		logicDocument.FinalizeAction();
	}

	function getPageHdrFtr(pageAbs, isHeader)
	{
		return logicDocument.HdrFtr.GetHdrFtr(pageAbs, isHeader);
	}

	function checkStaysOnPage(assert, pageAbs, isHeader, text)
	{
		let hdrFtr = logicDocument.HdrFtr.Get_CurHdrFtr();
		assert.strictEqual(logicDocument.GetDocPosType(), AscCommonWord.docpostype_HdrFtr, text + ": still in the header/footer");
		assert.strictEqual(hdrFtr, getPageHdrFtr(pageAbs, isHeader), text + ": current header/footer is the one shown on page " + (pageAbs + 1));
		assert.strictEqual(logicDocument.HdrFtr.Get_CurPage(), pageAbs, text + ": current page is page " + (pageAbs + 1));
		assert.strictEqual(hdrFtr.GetPage(), pageAbs, text + ": header/footer is calculated for page " + (pageAbs + 1));
		assert.strictEqual(hdrFtr.GetContent().GetAbsoluteStartPage(), pageAbs, text + ": cursor computed on page " + (pageAbs + 1));
	}

	function linkTest(name, isHeader, pageAbs, titlePage, evenAndOdd)
	{
		QUnit.test(name, function (assert)
		{
			let hf = setupSections(titlePage, evenAndOdd);
			assert.strictEqual(logicDocument.GetPagesCount(), 9, "nine pages");
			let own = isHeader ? hf.header[2] : hf.footer[2];
			let prev = isHeader ? hf.header[1] : hf.footer[1];
			assert.strictEqual(getPageHdrFtr(pageAbs, isHeader), own, "page " + (pageAbs + 1) + " shows the header/footer of section 3");

			editOnPage(own, pageAbs);
			checkStaysOnPage(assert, pageAbs, isHeader, "before");

			setLinkToPrevious(true);
			assert.strictEqual(getPageHdrFtr(pageAbs, isHeader), prev, "linked: page " + (pageAbs + 1) + " shows the header/footer of section 2");
			checkStaysOnPage(assert, pageAbs, isHeader, "linked");

			setLinkToPrevious(false);
			let copy = getPageHdrFtr(pageAbs, isHeader);
			assert.ok(copy !== prev && copy !== own, "unlinked: page " + (pageAbs + 1) + " shows a new copy");
			checkStaysOnPage(assert, pageAbs, isHeader, "unlinked");

			logicDocument.Document_Undo();
			assert.strictEqual(getPageHdrFtr(pageAbs, isHeader), prev, "undo of unlink: linked again");
			checkStaysOnPage(assert, pageAbs, isHeader, "undo of unlink");

			logicDocument.Document_Undo();
			assert.strictEqual(getPageHdrFtr(pageAbs, isHeader), own, "undo of link: own header/footer again");
			checkStaysOnPage(assert, pageAbs, isHeader, "undo of link");

			logicDocument.Document_Redo();
			assert.strictEqual(getPageHdrFtr(pageAbs, isHeader), prev, "redo of link: linked again");
			checkStaysOnPage(assert, pageAbs, isHeader, "redo of link");
			EvenAndOddHeaders = false;
			logicDocument.GetFinalSectPr().Set_TitlePage(false);
		});
	}

	QUnit.module("Link to previous keeps the page");

	linkTest("Header of page 7 (first page of section 3)", true, 6);
	linkTest("Header of page 8 (middle of section 3)", true, 7);
	linkTest("Footer of page 9 (last page of section 3)", false, 8);
	linkTest("First-page header of page 7", true, 6, true, false);
	linkTest("First-page footer of page 7", false, 6, true, false);
	linkTest("Even-page header of page 8", true, 7, false, true);

});
