"use strict";

// Regression test for the "invisible first column" bug in continuous multi-column
// sections that start mid-page (Euro-Office / ONLYOFFICE bug #73801): when such a
// section overflows onto the next page, recalculation re-enters it via
// recalcresult_PrevPage, which re-Init()s the section (rebuilding every column) and
// then resumed mid-section, leaving column 0 permanently empty -- its text remained
// selectable but was never painted.

$(function ()
{
	const logicDocument = AscTest.CreateLogicDocument();

	// The shared test measurer reports a fixed line height regardless of font size,
	// which flattens the vertical geometry and hides this layout bug (it only triggers
	// at certain mid-page start positions). Use a measurer whose metrics scale with the
	// font size so the column flow matches a real render.
	function useProportionalMeasurer()
	{
		const tm = window.g_oTextMeasurer || AscCommon.g_oTextMeasurer;
		let size = 10;
		tm.SetTextPr       = function(tp){ if (tp && (tp.FontSize || tp.FontSizeCS)) size = tp.FontSize || tp.FontSizeCS; };
		tm.SetFontSlot     = function(){};
		tm.SetFont         = function(f){ if (f && f.FontSize) size = f.FontSize; };
		tm.SetFontInternal = function(){};
		tm.GetHeight       = function(){ return size * 0.6; };
		tm.GetAscender     = function(){ return size * 0.46; };
		tm.GetDescender    = function(){ return size * 0.14; };
		tm.MeasureCode     = function(){ return { Width : size * 0.5 }; };
		tm.Measure         = function(){ return { Width : size * 0.5 }; };
		tm.Measure2Code    = function(){ return { Width : size * 0.5 }; };
	}

	function setupPage(sectPr)
	{
		sectPr.SetPageSize(210, 297);
		sectPr.SetPageMargins(25, 25, 25, 25);
	}

	function getTwoColumnSectionOnPage(nPage)
	{
		const page = logicDocument.Pages[nPage];
		if (!page) return null;
		for (let i = 0; i < page.Sections.length; ++i)
		{
			const sec = page.Sections[i];
			if (sec.Columns && sec.Columns.length === 2 && sec.YLimit - sec.Y > 1)
				return sec;
		}
		return null;
	}

	QUnit.test("Column 0 of a mid-page 2-column section that overflows to the next page is not empty", function (assert)
	{
		AscTest.ClearDocument();
		useProportionalMeasurer();

		// Final (body) section: continuous, 2 equal columns.
		let body = AscTest.GetFinalSection();
		setupPage(body);
		body.Set_Type(Asc.c_oAscSectionBreakType.Continuous);
		body.Set_Columns_EqualWidth(true);
		body.Set_Columns_Num(2);
		body.Set_Columns_Space(8);

		// First (top) section: full-width single column ending mid-page. Its height is
		// what makes the 2-column section begin part-way down page 1 (the bug trigger).
		let topSect = new AscWord.SectPr(logicDocument);
		setupPage(topSect);
		topSect.Set_Columns_EqualWidth(true);
		topSect.Set_Columns_Num(1);

		for (let i = 0; i < 4; ++i)
		{
			let p = AscTest.CreateParagraph();
			AscTest.AddTextToParagraph(p, "Top line " + (i + 1));
			if (i === 3)
				p.Set_SectionPr(topSect);
			logicDocument.PushToContent(p);
		}

		// Body: a long first paragraph (an "abstract") + more, so the 2-column section
		// fills page 1 and overflows onto page 2 -- which triggers the PrevPage re-entry.
		let abstract = AscTest.CreateParagraph();
		AscTest.AddTextToParagraph(abstract, "Abstract - " + ("the abstract text should be a single paragraph with many words so that it wraps across many lines and naturally flows from the first column into the second column of the section. ").repeat(3));
		logicDocument.PushToContent(abstract);
		for (let i = 0; i < 12; ++i)
		{
			let p = AscTest.CreateParagraph();
			AscTest.AddTextToParagraph(p, "Body paragraph " + (i + 1) + " with several words to add height to the column flow.");
			logicDocument.PushToContent(p);
		}

		AscTest.Recalculate();

		assert.ok(logicDocument.GetPagesCount() >= 2, "the section spans more than one page (got " + logicDocument.GetPagesCount() + ")");

		const sec = getTwoColumnSectionOnPage(0);
		assert.ok(!!sec, "a 2-column section exists on page 0");
		if (!sec) return;

		assert.strictEqual(sec.Columns[1].Empty, false, "column 1 has content");
		assert.strictEqual(sec.Columns[0].Empty, false, "column 0 has content (regression: it was left empty)");
		assert.ok(sec.Columns[0].EndPos >= sec.Columns[0].Pos, "column 0 draw range is non-empty");
	});
});
