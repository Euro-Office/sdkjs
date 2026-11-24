/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

$(function () {
	
	let logicDocument = AscTest.CreateLogicDocument();
	// Выставим стандартные настройки для параграфа
	logicDocument.GetStyles().Set_DefaultParaPr(AscWord.CParaPr.fromObject({
		Spacing : {
			After : 10 * g_dKoef_pt_to_mm,
			LineRule : linerule_Auto,
			Line : 259 / 240
		}
	}));

	QUnit.module("Test api for the document editor");


	QUnit.test("Test AddText/RemoveSelection", function (assert)
	{
		AscTest.ClearDocument();

		let p = new AscWord.Paragraph();
		logicDocument.AddToContent(0, p);

		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(), "\r\n", "Check empty selection");

		logicDocument.AddTextWithPr("Hello World!");

		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(false), "Hello World!\r\n", "Add text 'Hello World!'");

		let p2 = new AscWord.Paragraph();
		logicDocument.AddToContent(1, p2);

		logicDocument.RemoveSelection();
		p2.SetThisElementCurrent();
		p2.MoveCursorToStartPos();

		logicDocument.AddTextWithPr("Second paragraph");

		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(false), "Hello World!\r\nSecond paragraph\r\n", "Add text to the second paragraph");

		logicDocument.AddTextWithPr("Test");
		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectedText(false), "Test\r\n", "Replace all with adding text 'Test'");

		function StartTest(text)
		{
			AscTest.ClearDocument();
			let p = new AscWord.Paragraph();
			logicDocument.AddToContent(0, p);
			logicDocument.AddTextWithPr(text);
			logicDocument.MoveCursorToEndPos();
		}

		function GetText()
		{
			logicDocument.SelectAll();
			let text = logicDocument.GetSelectedText(false, {ParaSeparator : ""});
			logicDocument.RemoveSelection();
			return text;
		}

		let settings = new AscCommon.CAddTextSettings();

		settings.SetWrapWithSpaces(false);

		StartTest("Text");
		AscTest.MoveCursorLeft();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Tex123t", "Check text Tex123t");

		StartTest("Tex t");
		AscTest.MoveCursorLeft();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Tex 123t", "Check text Tex 123t");

		StartTest("Tex t");
		AscTest.MoveCursorLeft();
		AscTest.MoveCursorLeft();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Tex123 t", "Check text Tex123 t");

		StartTest("Text");
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Text123", "Check text Text123");

		StartTest("Text");
		logicDocument.MoveCursorToStartPos();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "123Text", "Check text 123Text");

		settings.SetWrapWithSpaces(true);

		StartTest("Text");
		AscTest.MoveCursorLeft();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Tex 123 t", "Check wrap spaces Tex 123 t");

		StartTest("Tex t");
		AscTest.MoveCursorLeft();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Tex 123 t", "Check wrap spaces Tex 123 t");

		StartTest("Tex t");
		AscTest.MoveCursorLeft();
		AscTest.MoveCursorLeft();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Tex 123 t", "Check wrap spaces Tex 123 t");

		StartTest("Text");
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "Text 123", "Check wrap spaces Text 123");

		StartTest("Text");
		logicDocument.MoveCursorToStartPos();
		logicDocument.AddTextWithPr("123", settings);
		assert.strictEqual(GetText(), "123 Text", "Check wrap spaces 123 Text");


	});
	
	QUnit.test("Change numbering level", function(assert)
	{
		AscTest.ClearDocument();
		
		let p1 = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p1);
		
		let p2 = AscTest.CreateParagraph();
		logicDocument.AddToContent(1, p2);
		
		AscTest.MoveCursorToParagraph(p1, true);
		AscTest.AddNumbering(1, 1);
		AscTest.MoveCursorToParagraph(p2, true);
		AscTest.AddNumbering(1, 1);
		
		let numPr1 = p1.GetNumPr();
		let numPr2 = p2.GetNumPr();
		
		assert.strictEqual(!!numPr1, true, "Check paragraph 1 numbering");
		assert.strictEqual(!!numPr2, true, "Check paragraph 2 numbering");
		assert.deepEqual(numPr1, numPr2, "Check paragraphs have the same paragraph 2 numbering");
		assert.strictEqual(numPr2.Lvl, 0, "Check numbering lvl of paragraph 2");
		
		// При двойном Enter на нулевом уровне в пустом параграфе нумерация должна убираться
		assert.strictEqual(p2.IsEmpty(), true, "Check if paragraph 2 is empty");
		
		AscTest.MoveCursorToParagraph(p2, true);
		AscTest.PressKey(AscTest.Key.enter);
		
		assert.strictEqual(!!p2.GetNumPr(), false, "Check of removing numbering from paragraph 2");
		
		
		AscTest.MoveCursorToParagraph(p2, true);
		AscTest.AddNumbering(1, 1);
		assert.strictEqual(!!p2.GetNumPr(), true, "Check paragraph 2 numbering");
		assert.deepEqual(p1.GetNumPr(), p2.GetNumPr(), "Check paragraphs have the same paragraph 2 numbering");
		
		AscTest.SetParagraphNumberingLvl(p2, 1);
		assert.deepEqual(p2.GetNumPr().Lvl, 1, "Check level of second paragraph");
		
		// При двойном Enter на ненулевом уровне в пустом параграфе уровень нумераации должен уменьшаться
		AscTest.MoveCursorToParagraph(p2, true);
		AscTest.PressKey(AscTest.Key.enter);
		
		assert.strictEqual(!!p2.GetNumPr(), true, "Check paragraph 2 numbering");
		assert.deepEqual(p2.GetNumPr().Lvl, 0, "Check level of second paragraph");
	});
	
	QUnit.test("Test add/remove space before/after paragraph", function(assert)
	{
		const epsilon = 0.001;
		
		AscTest.ClearDocument();
		let p = [];
		
		for (let i = 0; i < 3; ++i)
		{
			p[i] = AscTest.CreateParagraph();
			logicDocument.AddToContent(i, p[i]);
		}
		
		function checkSpacing(spacings)
		{
			for (let i = 0; i < 3; ++i)
			{
				let paraPr = p[i].GetCompiledParaPr(false);
				
				assert.close(paraPr.Spacing.Before, spacings[i][0], epsilon, "Check paragraph " + (i + 1) + " spacing before");
				assert.close(paraPr.Spacing.After, spacings[i][1], epsilon, "Check paragraph " + (i + 1) + " spacing after");
			}
		}
		function checkHaveSpace(haveSpaceBefore, haveSpaceAfter)
		{
			assert.strictEqual(AscTest.Editor.asc_haveSpaceBeforeParagraph(), haveSpaceBefore, "Check have space before");
			assert.strictEqual(AscTest.Editor.asc_haveSpaceAfterParagraph(), haveSpaceAfter, "Check have space after");
		}
		
		checkSpacing([
			[0, 10 * g_dKoef_pt_to_mm],
			[0, 10 * g_dKoef_pt_to_mm],
			[0, 10 * g_dKoef_pt_to_mm]
		]);
		
		AscTest.MoveCursorToParagraph(p[0]);
		checkHaveSpace(false, true);
		
		AscTest.Editor.asc_addSpaceBeforeParagraph();
		
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[0, 10 * g_dKoef_pt_to_mm],
			[0, 10 * g_dKoef_pt_to_mm]
		]);
		checkHaveSpace(true, true);
		
		AscTest.MoveCursorToParagraph(p[1]);
		AscTest.Editor.asc_removeSpaceAfterParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[0, 0],
			[0, 10 * g_dKoef_pt_to_mm]
		]);
		checkHaveSpace(false, false);
		
		AscTest.SelectDocumentRange(0, 2);
		AscTest.Editor.asc_removeSpaceAfterParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 0],
			[0, 0],
			[0, 0]
		]);
		checkHaveSpace(true, false);
		
		AscTest.Editor.asc_removeSpaceBeforeParagraph();
		checkSpacing([
			[0, 0],
			[0, 0],
			[0, 0]
		]);
		checkHaveSpace(false, false);
		
		AscTest.Editor.asc_addSpaceBeforeParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 0],
			[12 * g_dKoef_pt_to_mm, 0],
			[12 * g_dKoef_pt_to_mm, 0]
		]);
		checkHaveSpace(true, false);
		
		AscTest.Editor.asc_addSpaceAfterParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm]
		]);
		checkHaveSpace(true, true);
		
		let paraStyle = AscTest.CreateParagraphStyle("SpacingStyle");
		paraStyle.SetParaPr(AscWord.CParaPr.fromObject({
			Spacing : {
				Before : 25 * g_dKoef_pt_to_mm,
				After  : 0
			}
		}));
		p[1].SetPStyle(paraStyle.GetId());
		
		AscTest.SelectDocumentRange(1, 2);
		AscTest.Editor.asc_removeSpaceBeforeParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[0, 0],
			[0, 10 * g_dKoef_pt_to_mm]
		]);
		checkHaveSpace(false, false);
		
		AscTest.Editor.asc_addSpaceBeforeParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[25 * g_dKoef_pt_to_mm, 0],
			[25 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm]
		]);
		checkHaveSpace(true, false);
		
		AscTest.Editor.asc_addSpaceAfterParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[25 * g_dKoef_pt_to_mm, 12 * g_dKoef_pt_to_mm],
			[25 * g_dKoef_pt_to_mm, 12 * g_dKoef_pt_to_mm]
		]);
		checkHaveSpace(true, true);
		
		AscTest.Editor.asc_removeSpaceAfterParagraph();
		checkSpacing([
			[12 * g_dKoef_pt_to_mm, 10 * g_dKoef_pt_to_mm],
			[25 * g_dKoef_pt_to_mm, 0],
			[25 * g_dKoef_pt_to_mm, 0]
		]);
		checkHaveSpace(true, false);
		
		
	});
	
	QUnit.test("Get text/selected text", function(assert)
	{
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		
		let run = AscTest.CreateRun();
		logicDocument.AddToContent(0, p);
		p.AddToContentToEnd(run);
		run.AddText("The quick brown fox jumps over the lazy dog");
		AscTest.MoveCursorToParagraph(p, true);
		AscTest.MoveCursorRight(false, false, 4);
		assert.strictEqual(logicDocument.GetSelectedText(), "", "Check no selection");
		AscTest.MoveCursorRight(true, false, 5);
		assert.strictEqual(logicDocument.GetSelectedText(), "quick", "Select word 'quick'");
		
		// TODO: Переделать заполнение формулы по-нормальному
		let math = AscTest.CreateMath();
		p.AddToContentToEnd(math);
		math.SetThisElementCurrent();
		let text = "abcd";
		for (let i = 0; i < text.length; ++i)
		{
			math.Add(new AscWord.CRunText(text.charCodeAt(i)));
		}
		
		AscTest.MoveCursorToParagraph(p, false);
		AscTest.MoveCursorLeft(false, false, 4); // 1 сдвиг для захода в формулу
		AscTest.MoveCursorRight(true, false, 2);
		assert.strictEqual(logicDocument.GetSelectedText(), "bc", "Add math with text 'abcd' and partially select the text");
		
	});

	function CreateParagraphWithText(text)
	{
		let p = AscTest.CreateParagraph();
		AscTest.AddTextToParagraph(p, text);
		return p;
	}

	function GetDocumentTotal()
	{
		return AscBuilder.Word.Api.GetDocument.call(AscTest.Editor).GetStatistics();
	}

	function CheckStatistics(assert, expected, message)
	{
		let stats = logicDocument.GetSelectionStatistics();
		assert.strictEqual(stats.WordsCount, expected.WordsCount, message + ": words");
		assert.strictEqual(stats.ParagraphCount, expected.ParagraphCount, message + ": paragraphs");
		assert.strictEqual(stats.SymbolsCount, expected.SymbolsCount, message + ": symbols");
		assert.strictEqual(stats.SymbolsWSCount, expected.SymbolsWSCount, message + ": symbols with spaces");
	}

	QUnit.module("Test selection statistics");

	QUnit.test("No selection", function (assert)
	{
		AscTest.ClearDocument();
		let p = CreateParagraphWithText("The quick brown fox");
		logicDocument.AddToContent(0, p);
		AscTest.MoveCursorToParagraph(p, true);

		assert.strictEqual(logicDocument.GetSelectionStatistics(), null, "Collapsed cursor returns null");
	});

	QUnit.test("Part of a paragraph", function (assert)
	{
		AscTest.ClearDocument();
		let p = CreateParagraphWithText("The quick brown fox");
		logicDocument.AddToContent(0, p);
		AscTest.SelectParagraphRange(p, 4, 15);
		assert.strictEqual(logicDocument.GetSelectedText(), "quick brown", "Selected text");

		CheckStatistics(assert, {
			WordsCount     : 2,
			ParagraphCount : 1,
			SymbolsCount   : 10,
			SymbolsWSCount : 11
		}, "Partial paragraph");
	});

	QUnit.test("Several paragraphs", function (assert)
	{
		AscTest.ClearDocument();
		logicDocument.AddToContent(0, CreateParagraphWithText("Hello world"));
		logicDocument.AddToContent(1, CreateParagraphWithText("Second paragraph here"));
		logicDocument.AddToContent(2, CreateParagraphWithText("Not selected"));
		AscTest.SelectDocumentRange(0, 1);

		CheckStatistics(assert, {
			WordsCount     : 5,
			ParagraphCount : 2,
			SymbolsCount   : 29,
			SymbolsWSCount : 32
		}, "Two of three paragraphs");
	});

	QUnit.test("Table cells", function (assert)
	{
		AscTest.ClearDocument();
		let table = AscTest.CreateTable(2, 2);
		logicDocument.AddToContent(0, table);
		AscTest.AddTextToParagraph(table.GetRow(0).GetCell(0).GetContent().GetElement(0), "a b");
		AscTest.AddTextToParagraph(table.GetRow(0).GetCell(1).GetContent().GetElement(0), "c");
		AscTest.AddTextToParagraph(table.GetRow(1).GetCell(0).GetContent().GetElement(0), "d e f");
		AscTest.AddTextToParagraph(table.GetRow(1).GetCell(1).GetContent().GetElement(0), "g");
		AscTest.SelectTableCells(table, 0, 0, 1, 0);

		CheckStatistics(assert, {
			WordsCount     : 3,
			ParagraphCount : 2,
			SymbolsCount   : 3,
			SymbolsWSCount : 4
		}, "First table row");
	});

	QUnit.test("Select all matches the document total", function (assert)
	{
		AscTest.ClearDocument();
		let p1 = CreateParagraphWithText("Numbered item");
		let p2 = CreateParagraphWithText("Plain paragraph with nbsp");
		let table = AscTest.CreateTable(1, 2);
		logicDocument.AddToContent(0, p1);
		logicDocument.AddToContent(1, p2);
		logicDocument.AddToContent(2, table);
		AscTest.AddTextToParagraph(table.GetRow(0).GetCell(0).GetContent().GetElement(0), "cell one");
		AscTest.AddTextToParagraph(table.GetRow(0).GetCell(1).GetContent().GetElement(0), "cell two");
		AscTest.SelectParagraph(p1);
		AscTest.AddNumbering(1, 1);

		let total = GetDocumentTotal();
		assert.strictEqual(total.WordsCount, 11, "Document total includes the list number as a word");
		logicDocument.SelectAll();

		CheckStatistics(assert, total, "Select all");
	});

	QUnit.test("A bullet is counted as a symbol, not as a word", function (assert)
	{
		AscTest.ClearDocument();
		let p = CreateParagraphWithText("Bulleted item");
		logicDocument.AddToContent(0, p);
		AscTest.SelectParagraph(p);
		AscTest.AddNumbering(0, 1);
		assert.strictEqual(p.GetNumPr() && logicDocument.GetNumbering().GetNum(p.GetNumPr().NumId).GetLvl(0).IsBulleted(), true, "Paragraph has a bullet");

		logicDocument.SelectAll();
		let stats = logicDocument.GetSelectionStatistics();
		assert.strictEqual(stats.WordsCount, 2, "Only the text words are counted");
		assert.strictEqual(stats.SymbolsCount, "Bulleteditem".length + 1, "The bullet is counted as a symbol");
	});

	QUnit.test("List number counts only when the paragraph start is selected", function (assert)
	{
		AscTest.ClearDocument();
		let p = CreateParagraphWithText("Numbered item");
		logicDocument.AddToContent(0, p);
		AscTest.SelectParagraph(p);
		AscTest.AddNumbering(1, 1);

		AscTest.SelectParagraphRange(p, 0, 8);
		assert.strictEqual(logicDocument.GetSelectionStatistics().WordsCount, 2, "From the start: number and one word");

		AscTest.SelectParagraphRange(p, 9, 13);
		assert.strictEqual(logicDocument.GetSelectionStatistics().WordsCount, 1, "From the middle: one word");
	});

	QUnit.test("Selection statistics leave the track changes queue empty", function (assert)
	{
		AscTest.ClearDocument();
		logicDocument.AddToContent(0, CreateParagraphWithText("Hello world"));
		logicDocument.AddToContent(1, CreateParagraphWithText("Second paragraph"));
		let trackManager = logicDocument.GetTrackRevisionsManager();
		trackManager.ContinueTrackRevisions(true);
		logicDocument.SelectAll();

		logicDocument.GetSelectionStatistics();
		assert.strictEqual(trackManager.IsAllChecked(), true, "Nothing is queued for the track changes check");
	});

	QUnit.test("Several footnotes or endnotes", function (assert)
	{
		function AddTwoNotes(isEndnote)
		{
			AscTest.ClearDocument();
			let p1 = CreateParagraphWithText("Body one");
			let p2 = CreateParagraphWithText("Body two");
			logicDocument.AddToContent(0, p1);
			logicDocument.AddToContent(1, p2);

			let notes = [[p1, "alpha beta"], [p2, "gamma delta epsilon"]].map(function(item)
			{
				logicDocument.SetDocPosType(docpostype_Content);
				AscTest.MoveCursorToParagraph(item[0], false);
				let note = isEndnote ? logicDocument.AddEndnote() : logicDocument.AddFootnote();
				AscTest.AddTextToParagraph(note.GetElement(0), item[1]);
				return note;
			});
			AscTest.Recalculate();
			return notes;
		}

		let isDocumentEditor = editor.isDocumentEditor;
		editor.isDocumentEditor = true;

		try
		{
			AddTwoNotes(false);
			logicDocument.SetDocPosType(docpostype_Footnotes);
			logicDocument.Footnotes.SelectAll(1);
			assert.strictEqual(logicDocument.GetSelectionStatistics().WordsCount, 5, "Two footnotes");

			// The endnotes controller has no SelectAll, so select from the first endnote to the second one directly
			let endnotes = AddTwoNotes(true);
			let selection = logicDocument.Endnotes.Selection;
			logicDocument.SetDocPosType(docpostype_Endnotes);
			selection.Use              = true;
			selection.Direction        = 1;
			selection.Start.Endnote    = endnotes[0];
			selection.End.Endnote      = endnotes[1];
			selection.Endnotes         = {};
			endnotes.forEach(function(endnote)
			{
				selection.Endnotes[endnote.GetId()] = endnote;
				endnote.SelectAll(1);
			});
			assert.strictEqual(logicDocument.GetSelectionStatistics().WordsCount, 5, "Two endnotes");
		}
		finally
		{
			logicDocument.SetDocPosType(docpostype_Content);
			editor.isDocumentEditor = isDocumentEditor;
		}
	});

	QUnit.test("Pages spanned by the selection", function (assert)
	{
		AscTest.ClearDocument();
		let p1 = CreateParagraphWithText("First page");
		let p2 = CreateParagraphWithText("Second page");
		let p3 = CreateParagraphWithText("Third page");
		p2.Set_PageBreakBefore(true);
		p3.Set_PageBreakBefore(true);
		logicDocument.AddToContent(0, p1);
		logicDocument.AddToContent(1, p2);
		logicDocument.AddToContent(2, p3);
		AscTest.Recalculate();
		assert.strictEqual(logicDocument.GetPagesCount(), 3, "Document has three pages");

		AscTest.SelectParagraphRange(p1, 0, 5);
		assert.strictEqual(logicDocument.GetSelectionStatistics().PageCount, 1, "Part of one page");

		AscTest.SelectDocumentRange(0, 1);
		assert.strictEqual(logicDocument.GetSelectionStatistics().PageCount, 2, "First two pages");

		logicDocument.SelectAll();
		assert.strictEqual(logicDocument.GetSelectionStatistics().PageCount, 3, "Select all");
	});

	QUnit.test("Selection in a header has no page count", function (assert)
	{
		AscTest.ClearDocument();
		logicDocument.AddToContent(0, CreateParagraphWithText("Body"));
		let header = AscTest.CreateDefaultHeader(logicDocument.GetFinalSectPr());
		let p = header.GetElement(0);
		AscTest.AddTextToParagraph(p, "Header text");
		AscTest.Recalculate();
		AscTest.SelectParagraph(p);

		let stats = logicDocument.GetSelectionStatistics();
		assert.strictEqual(stats.WordsCount, 2, "Header words are counted");
		assert.strictEqual(stats.PageCount, -1, "Pages are not counted");
	});

	QUnit.module("Test spellcheck navigation");

	// Marks a word of a single-run paragraph as misspelled, the way a spellcheck response does
	function MarkMisspelled(paragraph, word, occurrence)
	{
		let run = null;
		paragraph.CheckRunContent(function(item)
		{
			if (!run && -1 !== item.GetText().indexOf(word))
				run = item;
		});
		let text  = run.GetText();
		let start = -1;
		for (let i = 0; i <= (occurrence || 0); ++i)
		{
			start = text.indexOf(word, start + 1);
		}
		let element = new AscWord.CParagraphSpellCheckerElement(run, start, run, start + word.length, word, 1033);
		run.AddSpellCheckerElement(new AscWord.SpellMarkStart(element));
		run.AddSpellCheckerElement(new AscWord.SpellMarkEnd(element));
		paragraph.SpellChecker.Elements.push(element);
		element.SetWrong(["fixed"]);
		return element;
	}

	// The test editor is not the document editor api, so lend it the spellcheck methods under test
	function CallSpellApi(name)
	{
		["asc_nextWord", "private_SendNextWordVariants", "private_GetSelectedMisspelledElement",
			"asc_replaceMisspelledWord", "asc_setSpellCheckErrorsLimit", "SpellCheck_CallBack",
			"sync_SpellCheckVariantsFound"].forEach(function(method)
		{
			AscTest.Editor[method] = Asc.asc_docs_api.prototype[method];
		});
		return AscTest.Editor[name].apply(AscTest.Editor, Array.prototype.slice.call(arguments, 1));
	}

	// Records the events sent with the given name while func runs
	function CatchEvents(name, func)
	{
		let events = [];
		let callback = function()
		{
			events.push(Array.prototype.slice.call(arguments));
		};
		AscTest.Editor.asc_registerCallback(name, callback);
		try
		{
			func();
		}
		finally
		{
			AscTest.Editor.asc_unregisterCallback(name, callback);
		}
		return events;
	}

	// Replaces the spellchecker of the editor: requests are recorded, dictionaries are given by haveDictionary
	function MockSpellCheckApi(haveDictionary, func)
	{
		let spellCheckApi = AscTest.Editor.SpellCheckApi;
		let requests = [];
		AscTest.Editor.SpellCheckApi = {
			spellCheck : function(request)
			{
				requests.push(request);
			},
			checkDictionary : function(lang)
			{
				return haveDictionary(lang);
			}
		};
		try
		{
			func(requests);
		}
		finally
		{
			AscTest.Editor.SpellCheckApi = spellCheckApi;
		}
		return requests;
	}

	function CreateMisspelledDocument()
	{
		AscTest.ClearDocument();
		let p1 = CreateParagraphWithText("Helo wrld good");
		let p2 = CreateParagraphWithText("Fine tset");
		logicDocument.AddToContent(0, p1);
		logicDocument.AddToContent(1, p2);
		MarkMisspelled(p1, "Helo");
		MarkMisspelled(p1, "wrld");
		MarkMisspelled(p2, "tset");
		return [p1, p2];
	}

	QUnit.test("Select the next misspelled word", function (assert)
	{
		let paragraphs = CreateMisspelledDocument();
		AscTest.MoveCursorToParagraph(paragraphs[0], true);

		let found;
		let events = CatchEvents("asc_onSpellCheckVariantsFound", function()
		{
			found = CallSpellApi("asc_nextWord");
		});
		assert.strictEqual(found, true, "A misspelled word is found");
		assert.strictEqual(logicDocument.GetSelectedText(), "Helo", "First word");
		assert.strictEqual(events.length, 1, "The known suggestions are sent at once");
		assert.deepEqual([events[0][0].get_Word(), events[0][0].get_Variants(), events[0][0].get_Checked()],
			["Helo", ["fixed"], false], "The event has the word and its suggestions");
		let spellCheck = null;
		let syncSpellCheck = window.editor.sync_SpellCheckCallback;
		window.editor.sync_SpellCheckCallback = function(word, checked, variants)
		{
			spellCheck = {word : word, variants : variants};
		};
		paragraphs[0].Document_UpdateInterfaceState();
		window.editor.sync_SpellCheckCallback = syncSpellCheck;
		assert.deepEqual(spellCheck, {word : "Helo", variants : ["fixed"]}, "Interface gets the misspelled word and the suggestions");
		logicDocument.SelectNextMisspelledWord();
		assert.strictEqual(logicDocument.GetSelectedText(), "wrld", "Second word");
		logicDocument.SelectNextMisspelledWord();
		assert.strictEqual(logicDocument.GetSelectedText(), "tset", "Word in the next paragraph");
		logicDocument.SelectNextMisspelledWord();
		assert.strictEqual(logicDocument.GetSelectedText(), "Helo", "Continues from the start of the document");
	});

	QUnit.test("The view moves to the misspelled word", function (assert)
	{
		AscTest.ClearDocument();
		let p1 = CreateParagraphWithText("First page");
		let p2 = CreateParagraphWithText("Secnd page");
		p2.Set_PageBreakBefore(true);
		logicDocument.AddToContent(0, p1);
		logicDocument.AddToContent(1, p2);
		MarkMisspelled(p2, "Secnd");
		AscTest.Recalculate();
		AscTest.MoveCursorToParagraph(p1, true);

		// The test editor has no drawing document for the cursor, so give it one that records the cursor page
		let targetPage = -1;
		let drawingDocument = Object.create(logicDocument.GetDrawingDocument());
		drawingDocument.UpdateTarget = function(x, y, pageIndex)
		{
			targetPage = pageIndex;
		};
		let getDrawingDocument = Asc.editor.getDrawingDocument;
		Asc.editor.getDrawingDocument = function()
		{
			return drawingDocument;
		};
		try
		{
			logicDocument.SelectNextMisspelledWord();
		}
		finally
		{
			Asc.editor.getDrawingDocument = getDrawingDocument;
		}
		assert.strictEqual(logicDocument.GetSelectedText(), "Secnd", "Word on the second page is selected");
		assert.strictEqual(targetPage, 1, "The view is moved to the second page");
	});

	QUnit.test("The misspelled word is scrolled to the middle of the screen", function (assert)
	{
		AscTest.ClearDocument();
		let p1 = CreateParagraphWithText("First page");
		let p2 = CreateParagraphWithText("Secnd page");
		p2.Set_PageBreakBefore(true);
		logicDocument.AddToContent(0, p1);
		logicDocument.AddToContent(1, p2);
		MarkMisspelled(p2, "Secnd");
		AscTest.Recalculate();
		AscTest.MoveCursorToParagraph(p1, true);

		let wordControl = logicDocument.GetDrawingDocument().m_oWordControl;
		let centered = null;
		wordControl.ScrollToAbsolutePosition = function(x, y, pageIndex, isBottom, isCenter)
		{
			centered = {Y : y, Page : pageIndex, IsCenter : isCenter};
		};
		try
		{
			logicDocument.SelectNextMisspelledWord();
		}
		finally
		{
			delete wordControl.ScrollToAbsolutePosition;
		}
		assert.strictEqual(logicDocument.GetSelectedText(), "Secnd", "Word on the second page is selected");
		assert.ok(centered, "The view is asked to centre a position");
		assert.strictEqual(centered && centered.Page, 1, "The centred position is on the page of the word");
		assert.ok(centered && centered.Y > 0, "The centred position is inside the page");
		assert.strictEqual(centered && centered.IsCenter, true, "The position is put in the middle of the screen");
	});

	QUnit.test("Word with the cursor inside is found", function (assert)
	{
		let paragraphs = CreateMisspelledDocument();
		AscTest.MoveCursorToParagraph(paragraphs[0], true);
		paragraphs[0].MoveCursorRight(false, false);
		paragraphs[0].MoveCursorRight(false, false);

		logicDocument.SelectNextMisspelledWord();
		assert.strictEqual(logicDocument.GetSelectedText(), "Helo", "Word with the cursor");
	});

	QUnit.test("Ignored and correct words are skipped", function (assert)
	{
		let paragraphs = CreateMisspelledDocument();
		paragraphs[0].SpellChecker.Elements[0].SetCorrect();
		paragraphs[0].SpellChecker.Elements[1].Checked = true;
		AscTest.MoveCursorToParagraph(paragraphs[0], true);

		logicDocument.SelectNextMisspelledWord();
		assert.strictEqual(logicDocument.GetSelectedText(), "tset", "Only the misspelled word is found");

		paragraphs[1].SpellChecker.Elements[0].SetCorrect();
		assert.strictEqual(logicDocument.SelectNextMisspelledWord(), null, "No misspelled words left");
	});

	QUnit.test("Headers and footers are checked after the main text", function (assert)
	{
		let paragraphs = CreateMisspelledDocument();
		let sectPr = logicDocument.GetFinalSectPr();
		let header = AscTest.CreateDefaultHeader(sectPr);
		let headerParagraph = header.GetElement(0);
		AscTest.AddTextToParagraph(headerParagraph, "Hedaer text");
		MarkMisspelled(headerParagraph, "Hedaer");

		let footer = new AscCommonWord.CHeaderFooter(logicDocument.HdrFtr, logicDocument, logicDocument.DrawingDocument, AscCommon.hdrftr_Footer);
		sectPr.Set_Footer_Default(footer);
		let footerParagraph = footer.GetContent().GetElement(0);
		AscTest.AddTextToParagraph(footerParagraph, "Footre text");
		MarkMisspelled(footerParagraph, "Footre");
		AscTest.Recalculate();

		AscTest.MoveCursorToParagraph(paragraphs[1], false);
		logicDocument.SelectNextMisspelledWord();
		assert.strictEqual(logicDocument.GetSelectedText(), "Hedaer", "Header after the main text");
		assert.strictEqual(logicDocument.GetDocPosType(), AscCommonWord.docpostype_HdrFtr, "Cursor is in the header");
		logicDocument.SelectNextMisspelledWord();
		assert.strictEqual(logicDocument.GetSelectedText(), "Footre", "Footer after the header");
		logicDocument.SelectNextMisspelledWord();
		assert.strictEqual(logicDocument.GetSelectedText(), "Helo", "Continues from the start of the main text");
		assert.strictEqual(logicDocument.GetDocPosType(), AscCommonWord.docpostype_Content, "Cursor is back in the main text");

		sectPr.Set_Header_Default(null);
		sectPr.Set_Footer_Default(null);
	});

	QUnit.test("Footnotes and endnotes are checked after the main text", function (assert)
	{
		let isDocumentEditor = editor.isDocumentEditor;
		editor.isDocumentEditor = true;
		try
		{
			AscTest.ClearDocument();
			let p = CreateParagraphWithText("Helo body");
			logicDocument.AddToContent(0, p);
			MarkMisspelled(p, "Helo");

			AscTest.MoveCursorToParagraph(p, false);
			let footnote = logicDocument.AddFootnote();
			AscTest.AddTextToParagraph(footnote.GetElement(0), "Fotnote text");
			MarkMisspelled(footnote.GetElement(0), "Fotnote");

			logicDocument.SetDocPosType(AscCommonWord.docpostype_Content);
			AscTest.MoveCursorToParagraph(p, false);
			let endnote = logicDocument.AddEndnote();
			AscTest.AddTextToParagraph(endnote.GetElement(0), "Endnte text");
			MarkMisspelled(endnote.GetElement(0), "Endnte");
			AscTest.Recalculate();

			logicDocument.SetDocPosType(AscCommonWord.docpostype_Content);
			AscTest.MoveCursorToParagraph(p, true);
			logicDocument.SelectNextMisspelledWord();
			assert.strictEqual(logicDocument.GetSelectedText(), "Helo", "Main text first");
			logicDocument.SelectNextMisspelledWord();
			assert.strictEqual(logicDocument.GetSelectedText(), "Fotnote", "Then the footnote");
			assert.strictEqual(logicDocument.GetDocPosType(), AscCommonWord.docpostype_Footnotes, "Cursor is in the footnote");
			logicDocument.SelectNextMisspelledWord();
			assert.strictEqual(logicDocument.GetSelectedText(), "Endnte", "Then the endnote");
			assert.strictEqual(logicDocument.GetDocPosType(), AscCommonWord.docpostype_Endnotes, "Cursor is in the endnote");
			logicDocument.SelectNextMisspelledWord();
			assert.strictEqual(logicDocument.GetSelectedText(), "Helo", "Continues from the start of the main text");
			assert.strictEqual(logicDocument.GetDocPosType(), AscCommonWord.docpostype_Content, "Cursor is back in the main text");
		}
		finally
		{
			logicDocument.SetDocPosType(AscCommonWord.docpostype_Content);
			editor.isDocumentEditor = isDocumentEditor;
		}
	});

	QUnit.test("Text boxes are checked with the paragraph they are anchored to", function (assert)
	{
		AscTest.ClearDocument();
		let p = CreateParagraphWithText("Helo body");
		logicDocument.AddToContent(0, p);
		MarkMisspelled(p, "Helo");

		AscTest.MoveCursorToParagraph(p, false);
		let drawing = new ParaDrawing(100, 100, null, logicDocument.GetDrawingDocument(), logicDocument, null);
		let shapeTrack = new AscFormat.NewShapeTrack('rect', 0, 0, logicDocument.theme, null, null, null, 0);
		shapeTrack.track({}, 100, 100);
		let shape = shapeTrack.getShape(true, logicDocument.GetDrawingDocument(), null);
		shape.setBDeleted(false);
		shape.setParent(drawing);
		drawing.Set_GraphicObject(shape);
		drawing.Set_DrawingType(drawing_Anchor);
		drawing.Set_WrappingType(WRAPPING_TYPE_NONE);
		drawing.Set_Distance(0, 0, 0, 0);
		p.AddToParagraph(drawing);
		shape.createTextBoxContent();
		let boxParagraph = shape.getDocContent().GetElement(0);
		AscTest.AddTextToParagraph(boxParagraph, "Txet box");
		MarkMisspelled(boxParagraph, "Txet");
		AscTest.Recalculate();

		logicDocument.RemoveSelection();
		AscTest.MoveCursorToParagraph(p, false);
		logicDocument.SelectNextMisspelledWord();
		assert.strictEqual(logicDocument.GetSelectedText(), "Txet", "Word in the text box is found");
		logicDocument.SelectNextMisspelledWord();
		assert.strictEqual(logicDocument.GetSelectedText(), "Helo", "Then the main text again");
	});

	QUnit.test("From a text box only the group of its anchor paragraph is collected", function (assert)
	{
		AscTest.ClearDocument();
		let p1 = CreateParagraphWithText("Helo body");
		let p2 = CreateParagraphWithText("Fine tset");
		logicDocument.AddToContent(0, p1);
		logicDocument.AddToContent(1, p2);
		MarkMisspelled(p1, "Helo");
		MarkMisspelled(p2, "tset");

		AscTest.MoveCursorToParagraph(p2, false);
		let drawing = new ParaDrawing(100, 100, null, logicDocument.GetDrawingDocument(), logicDocument, null);
		let shapeTrack = new AscFormat.NewShapeTrack('rect', 0, 0, logicDocument.theme, null, null, null, 0);
		shapeTrack.track({}, 100, 100);
		let shape = shapeTrack.getShape(true, logicDocument.GetDrawingDocument(), null);
		shape.setBDeleted(false);
		shape.setParent(drawing);
		drawing.Set_GraphicObject(shape);
		drawing.Set_DrawingType(drawing_Anchor);
		drawing.Set_WrappingType(WRAPPING_TYPE_NONE);
		drawing.Set_Distance(0, 0, 0, 0);
		p2.AddToParagraph(drawing);
		shape.createTextBoxContent();
		let boxParagraph = shape.getDocContent().GetElement(0);
		AscTest.AddTextToParagraph(boxParagraph, "Txet box");
		MarkMisspelled(boxParagraph, "Txet");
		AscTest.Recalculate();

		let requested = [];
		function getGroup(groupIndex)
		{
			requested.push(groupIndex);
			return logicDocument.private_GetSpellCheckParagraphs(groupIndex);
		}
		assert.strictEqual(logicDocument.private_GetSpellCheckGroup(boxParagraph, getGroup), 1, "The text box is in the group of the second paragraph");
		assert.deepEqual(requested, [1], "Only that group is collected");

		let getParagraphs = logicDocument.private_GetSpellCheckParagraphs;
		requested = [];
		logicDocument.private_GetSpellCheckParagraphs = function(groupIndex)
		{
			requested.push(groupIndex);
			return getParagraphs.call(this, groupIndex);
		};
		try
		{
			logicDocument.RemoveSelection();
			AscTest.MoveCursorToParagraph(p1, false);
			logicDocument.SelectNextMisspelledWord();
			assert.strictEqual(logicDocument.GetSelectedText(), "Txet", "Word in the text box is found");

			requested = [];
			logicDocument.SelectNextMisspelledWord();
			assert.strictEqual(logicDocument.GetSelectedText(), "tset", "From the text box the search goes on in its anchor paragraph");
			assert.deepEqual(requested, [1], "Only the group of the anchor paragraph is collected");
		}
		finally
		{
			logicDocument.private_GetSpellCheckParagraphs = getParagraphs;
		}
	});

	QUnit.test("Spellcheck progress", function (assert)
	{
		let paragraphs = CreateMisspelledDocument();
		let spelling = logicDocument.Spelling;
		spelling.Reset();

		let state = logicDocument.GetSpellCheckState();
		assert.deepEqual([state.Checked, state.Total, state.IsComplete], [0, 0, true], "Nothing left to check");

		spelling.AddParagraphToCheck(paragraphs[0]);
		spelling.AddWaitingParagraph(paragraphs[1], 0, ["tset"], [1033]);
		state = logicDocument.GetSpellCheckState();
		assert.deepEqual([state.Checked, state.Total, state.IsComplete], [0, 2, false], "Queued and waiting paragraphs are not checked");

		spelling.RemoveWaitingParagraph(paragraphs[1]);
		state = logicDocument.GetSpellCheckState();
		assert.deepEqual([state.Checked, state.Total, state.IsComplete], [1, 2, false], "One of the two paragraphs is checked");

		spelling.AddParagraphToCheck(paragraphs[1]);
		spelling.AddCurrentParagraph(paragraphs[1]);
		state = logicDocument.GetSpellCheckState();
		assert.deepEqual([state.Checked, state.Total], [0, 2], "A paragraph in two lists counts once");

		spelling.Reset();
		state = logicDocument.GetSpellCheckState();
		assert.deepEqual([state.Checked, state.Total, state.IsComplete, state.IsStopped], [0, 0, true, false], "Check complete");

		spelling.Use = false;
		assert.strictEqual(logicDocument.GetSpellCheckState().IsStopped, true, "Spellcheck turned off");
		spelling.Use = true;
	});

	QUnit.test("Spellcheck progress is sent as an event when it changes", function (assert)
	{
		let paragraphs = CreateMisspelledDocument();
		let spelling = logicDocument.Spelling;
		spelling.Reset();
		spelling.Progress = null;

		let events = CatchEvents("asc_onSpellCheckProgress", function()
		{
			spelling.AddParagraphToCheck(paragraphs[0]);
			spelling.AddWaitingParagraph(paragraphs[1], 0, ["tset"], [1033]);
			spelling.private_SendProgress();
			spelling.private_SendProgress();
			spelling.RemoveWaitingParagraph(paragraphs[1]);
			spelling.private_SendProgress();
			spelling.Reset();
			spelling.private_SendProgress();
		});
		assert.deepEqual(events, [[0, 2, false, false], [1, 2, false, false], [0, 0, true, false]],
			"Checked, total, complete and stopped are sent once per change");

		spelling.Progress = null;
		events = CatchEvents("asc_onSpellCheckProgress", function()
		{
			spelling.ContinueSpellCheck();
		});
		assert.deepEqual(events, [[0, 0, true, false]], "The spellcheck timer sends the progress");
	});

	QUnit.test("The limit of 2000 errors is lifted on request", function (assert)
	{
		let spelling = logicDocument.Spelling;
		spelling.Reset();
		let getErrorsCount = spelling.GetErrorsCount;
		spelling.GetErrorsCount = function()
		{
			return 2001;
		};
		try
		{
			spelling.private_CheckErrorsExceed();
			assert.strictEqual(spelling.IsErrorsExceed(), true, "By default the check stops after 2000 errors");

			CallSpellApi("asc_setSpellCheckErrorsLimit", false);
			assert.strictEqual(spelling.IsErrorsExceed(), false, "The check goes on without the limit");
			spelling.private_CheckErrorsExceed();
			assert.strictEqual(spelling.IsErrorsExceed(), false, "And does not stop again");

			CallSpellApi("asc_setSpellCheckErrorsLimit", true);
			spelling.private_CheckErrorsExceed();
			assert.strictEqual(spelling.IsErrorsExceed(), true, "With the limit back the check stops again");
		}
		finally
		{
			spelling.GetErrorsCount = getErrorsCount;
			spelling.SetErrorsLimit(true);
			spelling.Reset();
		}
	});

	QUnit.test("Past the limit of 2000 errors the spellcheck timer does nothing", function (assert)
	{
		let paragraphs = CreateMisspelledDocument();
		let spelling = logicDocument.Spelling;
		spelling.Reset();
		spelling.Progress = null;
		let getErrorsCount = spelling.GetErrorsCount;
		spelling.GetErrorsCount = function()
		{
			return 2001;
		};
		let getPendingCount = spelling.GetPendingParagraphsCount;
		try
		{
			let events = CatchEvents("asc_onSpellCheckProgress", function()
			{
				spelling.ContinueSpellCheck();
			});
			assert.strictEqual(spelling.IsErrorsExceed(), true, "This tick hits the limit");
			assert.deepEqual(events, [[0, 0, true, false]], "The tick that hits the limit sends the progress");

			spelling.AddParagraphToCheck(paragraphs[0]);
			spelling.AddParagraphToCheck(paragraphs[1]);
			let counted = 0;
			spelling.GetPendingParagraphsCount = function()
			{
				++counted;
				return getPendingCount.call(this);
			};
			events = CatchEvents("asc_onSpellCheckProgress", function()
			{
				spelling.ContinueSpellCheck();
				spelling.ContinueSpellCheck();
			});
			assert.deepEqual(events, [], "A tick past the limit sends no progress");
			assert.strictEqual(counted, 0, "And does not count the pending paragraphs");
		}
		finally
		{
			spelling.GetErrorsCount = getErrorsCount;
			delete spelling.GetPendingParagraphsCount;
			spelling.Reset();
			spelling.Progress = null;
		}
	});

	// Runs the background check the way the editor timer does: a tick, then the spellchecker answers. The words in
	// wrongWords are misspelled.
	function RunSpellCheckTicks(requests, wrongWords, count)
	{
		let spelling = logicDocument.Spelling;
		for (let tick = 0; tick < count; ++tick)
		{
			spelling.ContinueSpellCheck();
			while (requests.length)
			{
				let request   = requests.shift();
				let paragraph = null;
				logicDocument.private_GetSpellCheckParagraphs().forEach(function(item)
				{
					if (item.GetId() === request["ParagraphId"])
						paragraph = item;
				});
				if (!paragraph)
					continue;

				if ("spell" === request["type"])
				{
					paragraph.SpellChecker.SpellCheckResponse(request["RecalcId"], request["usrWords"].map(function(word)
					{
						return !wrongWords[word];
					}));
				}
				else
				{
					paragraph.SpellChecker.SuggestResponse(request["RecalcId"], request["ElementId"], ["fixed"]);
				}
			}
		}
	}

	QUnit.test("The word changed last does not keep the spellcheck from completing", function (assert)
	{
		AscTest.ClearDocument();
		let p1 = CreateParagraphWithText("Helo there");
		let p2 = CreateParagraphWithText("Fine text wrld");
		logicDocument.AddToContent(0, p1);
		logicDocument.AddToContent(1, p2);
		AscTest.Recalculate();

		let spelling = logicDocument.Spelling;
		spelling.Reset();
		let isSpellCheckCurrentWord = AscTest.Editor.asc_IsSpellCheckCurrentWord;
		AscTest.Editor.asc_IsSpellCheckCurrentWord = function()
		{
			return false;
		};
		let wrongWords = {"Helo" : true, "wrld" : true};
		try
		{
			MockSpellCheckApi(function()
			{
				return true;
			}, function(requests)
			{
				logicDocument.RemoveSelection();
				AscTest.MoveCursorToParagraph(p2, true);
				logicDocument.RestartSpellCheck();
				RunSpellCheckTicks(requests, wrongWords, 5);
				assert.strictEqual(spelling.GetErrorsCount(), 2, "The two misspelled words are marked");

				// What the panel does: Next, then Change, as asc_replaceMisspelledWord does it, for every word
				let changed = [];
				for (let step = 0; step < 5 && CallSpellApi("asc_nextWord"); ++step)
				{
					let paragraph = logicDocument.GetCurrentParagraph();
					let element   = CallSpellApi("private_GetSelectedMisspelledElement", paragraph);
					changed.push(element.GetWord());
					logicDocument.StartAction(AscDFH.historydescription_Document_ReplaceMisspelledWord);
					paragraph.ReplaceMisspelledWord("fixed", element);
					paragraph.Document_SetThisElementCurrent(true);
					logicDocument.Recalculate();
					logicDocument.FinalizeAction();
					RunSpellCheckTicks(requests, wrongWords, 1);
				}
				assert.deepEqual(changed, ["wrld", "Helo"], "Both words are changed");
				assert.strictEqual(AscTest.GetParagraphText(p1), "fixed there", "The cursor is at the last changed word");

				RunSpellCheckTicks(requests, wrongWords, 5);
				assert.strictEqual(logicDocument.GetSpellCheckState().IsComplete, true, "After the last Next the check completes");
				assert.strictEqual(CallSpellApi("asc_nextWord"), false, "No more words");
				assert.deepEqual(p1.SpellChecker.Elements.filter(function(element)
				{
					return element.IsUndefined() || element.IsCurrent();
				}), [], "The changed word is checked");
			});
		}
		finally
		{
			if (isSpellCheckCurrentWord)
				AscTest.Editor.asc_IsSpellCheckCurrentWord = isSpellCheckCurrentWord;
			else
				delete AscTest.Editor.asc_IsSpellCheckCurrentWord;
			spelling.Reset();
		}
	});

	QUnit.test("Spellcheck counts as stopped when the spellchecker is disconnected", function (assert)
	{
		let spelling = logicDocument.Spelling;
		spelling.Reset();
		assert.strictEqual(logicDocument.GetSpellCheckState().IsStopped, false, "Spellcheck runs");
		spelling.TurnOff();
		try
		{
			assert.strictEqual(logicDocument.GetSpellCheckState().IsStopped, true, "No check without the spellchecker");
		}
		finally
		{
			spelling.TurnOn();
		}
	});

	QUnit.test("Suggestions for the next word are sent when the spellchecker answers", function (assert)
	{
		let paragraphs = CreateMisspelledDocument();
		paragraphs[0].SpellChecker.Elements[0].SetVariants(null);
		logicDocument.Spelling.Reset();
		AscTest.MoveCursorToParagraph(paragraphs[0], true);

		let events = [];
		let requests = MockSpellCheckApi(function()
		{
			return true;
		}, function(requests)
		{
			events = CatchEvents("asc_onSpellCheckVariantsFound", function()
			{
				CallSpellApi("asc_nextWord");
			});
			assert.strictEqual(events.length, 0, "No suggestions yet");
			let suggest = requests.filter(function(request)
			{
				return "suggest" === request["type"];
			});
			assert.ok(suggest.length >= 1, "The suggestions are requested");
			assert.deepEqual(suggest[0]["usrWords"], ["Helo"], "For the selected word");

			// What SpellCheck_CallBack does with the answer (its paragraph lookup does not work in the test editor)
			events = CatchEvents("asc_onSpellCheckVariantsFound", function()
			{
				paragraphs[0].SpellChecker.SuggestResponse(suggest[0]["RecalcId"], suggest[0]["ElementId"], ["Hello", "Help"]);
				CallSpellApi("private_SendNextWordVariants", paragraphs[0]);
			});
		});
		assert.ok(requests.length >= 1, "The spellchecker was asked");
		let withWord = events.filter(function(args)
		{
			return args[0];
		});
		assert.strictEqual(withWord.length, 1, "One event with the word");
		assert.deepEqual([withWord[0][0].get_Word(), withWord[0][0].get_Variants(), withWord[0][0].ParaId],
			["Helo", ["Hello", "Help"], paragraphs[0].GetId()], "The event has the word and the answer");
	});

	QUnit.test("A word without a dictionary gets an empty list of suggestions", function (assert)
	{
		let paragraphs = CreateMisspelledDocument();
		paragraphs[0].SpellChecker.Elements[0].SetVariants(null);
		AscTest.MoveCursorToParagraph(paragraphs[0], true);

		let events = [];
		let requests = MockSpellCheckApi(function()
		{
			return false;
		}, function()
		{
			events = CatchEvents("asc_onSpellCheckVariantsFound", function()
			{
				CallSpellApi("asc_nextWord");
			});
		});
		assert.strictEqual(requests.length, 0, "Nothing is requested");
		assert.strictEqual(events.length, 1, "The event is sent at once");
		assert.deepEqual([events[0][0].get_Word(), events[0][0].get_Variants()], ["Helo", []], "With no suggestions");
	});

	QUnit.test("Spellcheck lists languages without a dictionary", function (assert)
	{
		let paragraphs = CreateMisspelledDocument();
		let checkDictionary = editor.SpellCheckApi && editor.SpellCheckApi.checkDictionary;
		editor.SpellCheckApi = editor.SpellCheckApi || {};
		editor.SpellCheckApi.checkDictionary = function(lang)
		{
			return 1033 === lang;
		};
		try
		{
			let textPr = new AscCommonWord.CTextPr();
			textPr.Lang.Val = 3079;
			paragraphs[1].GetElement(0).Apply_TextPr(textPr, undefined, true);
			let defaultLang = logicDocument.GetStyles().Default.TextPr.Lang.Val;

			let state = logicDocument.GetSpellCheckState(true);
			assert.deepEqual(state.MissingLanguages.sort(), [defaultLang, 3079].filter(function(lang)
			{
				return 1033 !== lang;
			}).sort(), "German (Austria) has no dictionary");
			assert.deepEqual(logicDocument.GetSpellCheckState().MissingLanguages, [], "Languages are listed only on request");
		}
		finally
		{
			editor.SpellCheckApi.checkDictionary = checkDictionary;
		}
	});

	QUnit.test("Change all occurrences in the header too", function (assert)
	{
		AscTest.ClearDocument();
		let p = CreateParagraphWithText("tset one");
		logicDocument.AddToContent(0, p);
		MarkMisspelled(p, "tset");
		let sectPr = logicDocument.GetFinalSectPr();
		let headerParagraph = AscTest.CreateDefaultHeader(sectPr).GetElement(0);
		AscTest.AddTextToParagraph(headerParagraph, "tset header");
		MarkMisspelled(headerParagraph, "tset");

		logicDocument.ReplaceAllMisspelledWords("tset", "test");
		assert.strictEqual(AscTest.GetParagraphText(p), "test one", "Main text is replaced");
		assert.strictEqual(AscTest.GetParagraphText(headerParagraph), "test header", "Header is replaced");
		sectPr.Set_Header_Default(null);
	});

	QUnit.test("Change all occurrences of a misspelled word", function (assert)
	{
		AscTest.ClearDocument();
		let p1 = CreateParagraphWithText("tset one tset");
		let p2 = CreateParagraphWithText("two tset tsets");
		logicDocument.AddToContent(0, p1);
		logicDocument.AddToContent(1, p2);
		MarkMisspelled(p1, "tset");
		MarkMisspelled(p1, "tset", 1);
		MarkMisspelled(p2, "tset");
		MarkMisspelled(p2, "tsets");

		logicDocument.ReplaceAllMisspelledWords("tset", "test");
		assert.strictEqual(AscTest.GetParagraphText(p1), "test one test", "Both words in a paragraph are replaced");
		assert.strictEqual(AscTest.GetParagraphText(p2), "two test tsets", "Other misspelled words are kept");
	});

	QUnit.test("One Undo reverts the whole Change all", function (assert)
	{
		AscTest.ClearDocument();
		let p1 = CreateParagraphWithText("tset one tset");
		let p2 = CreateParagraphWithText("two tset");
		logicDocument.AddToContent(0, p1);
		logicDocument.AddToContent(1, p2);
		MarkMisspelled(p1, "tset");
		MarkMisspelled(p1, "tset", 1);
		MarkMisspelled(p2, "tset");

		let result = logicDocument.ReplaceAllMisspelledWords("tset", "test");
		assert.deepEqual(result, {Replaced : 3, Skipped : 0}, "Three words replaced");
		assert.strictEqual(AscTest.GetParagraphText(p1) + "|" + AscTest.GetParagraphText(p2), "test one test|two test", "Replaced");

		logicDocument.Document_Undo();
		assert.strictEqual(AscTest.GetParagraphText(p1) + "|" + AscTest.GetParagraphText(p2), "tset one tset|two tset", "One Undo restores every word");
	});

	QUnit.test("Change all through asc_replaceMisspelledWord", function (assert)
	{
		AscTest.ClearDocument();
		let p1 = CreateParagraphWithText("tset one");
		let p2 = CreateParagraphWithText("two tset");
		logicDocument.AddToContent(0, p1);
		logicDocument.AddToContent(1, p2);
		let element = MarkMisspelled(p1, "tset");
		MarkMisspelled(p2, "tset");

		let property = new AscCommon.asc_CSpellCheckProperty("tset", false, ["test"], p1.GetId(), element);
		let result = CallSpellApi("asc_replaceMisspelledWord", "test", property, true);
		assert.deepEqual(result, {"Replaced" : 2, "Skipped" : 0}, "The api says how many words were replaced");
		assert.strictEqual(AscTest.GetParagraphText(p1) + "|" + AscTest.GetParagraphText(p2), "test one|two test", "Both paragraphs are replaced");
	});

	QUnit.test("Change all skips words in locked content controls", function (assert)
	{
		AscTest.ClearDocument();
		let p = CreateParagraphWithText("tset one");
		logicDocument.AddToContent(0, p);
		MarkMisspelled(p, "tset");

		let cc = AscTest.CreateBlockLvlSdt();
		logicDocument.AddToContent(1, cc);
		let ccParagraph = cc.GetContent().GetElement(0);
		AscTest.AddTextToParagraph(ccParagraph, "locked tset");
		MarkMisspelled(ccParagraph, "tset");
		cc.SetContentControlLock(Asc.c_oAscSdtLockType.ContentLocked);

		AscTest.MoveCursorToParagraph(p, true);
		let result = logicDocument.ReplaceAllMisspelledWords("tset", "test");
		assert.strictEqual(AscTest.GetParagraphText(p), "test one", "The free word is replaced");
		assert.strictEqual(AscTest.GetParagraphText(ccParagraph), "locked tset", "The word in the locked control is kept");
		assert.deepEqual(result, {Replaced : 1, Skipped : 1}, "One word skipped");
	});

	QUnit.test("Change all skips words in a locked inline content control without the cursor", function (assert)
	{
		AscTest.ClearDocument();
		let p1 = CreateParagraphWithText("tset one");
		let p2 = CreateParagraphWithText("two ");
		logicDocument.AddToContent(0, p1);
		logicDocument.AddToContent(1, p2);
		let cc = new AscWord.CInlineLevelSdt();
		let ccRun = new AscWord.CRun();
		ccRun.AddText("tset");
		cc.AddToContent(0, ccRun);
		p2.AddToContent(1, cc);
		cc.SetContentControlLock(Asc.c_oAscSdtLockType.SdtContentLocked);
		MarkMisspelled(p1, "tset");
		MarkMisspelled(p2, "tset");

		AscTest.MoveCursorToParagraph(p1, true);
		let result = logicDocument.ReplaceAllMisspelledWords("tset", "test");
		assert.strictEqual(AscTest.GetParagraphText(p1), "test one", "The free word is replaced");
		assert.strictEqual(AscTest.GetParagraphText(p2), "two tset", "The word in the locked control is kept");
		assert.deepEqual(result, {Replaced : 1, Skipped : 1}, "One word skipped");
	});

	QUnit.test("Change all skips paragraphs locked by another user", function (assert)
	{
		AscTest.ClearDocument();
		let p1 = CreateParagraphWithText("tset one");
		let p2 = CreateParagraphWithText("two tset");
		logicDocument.AddToContent(0, p1);
		logicDocument.AddToContent(1, p2);
		MarkMisspelled(p1, "tset");
		MarkMisspelled(p2, "tset");
		p2.Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeOther, false);
		try
		{
			let result = logicDocument.ReplaceAllMisspelledWords("tset", "test");
			assert.strictEqual(AscTest.GetParagraphText(p1), "test one", "The free paragraph is replaced");
			assert.strictEqual(AscTest.GetParagraphText(p2), "two tset", "The locked paragraph is kept");
			assert.deepEqual(result, {Replaced : 1, Skipped : 1}, "One word skipped");
		}
		finally
		{
			p2.Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeNone, false);
		}
	});

	QUnit.test("asc_GetSelectionDocInfo", function (assert)
	{
		let getSelectionDocInfo = Asc.asc_docs_api.prototype.asc_GetSelectionDocInfo;

		AscTest.ClearDocument();
		let p = CreateParagraphWithText("The quick brown fox");
		logicDocument.AddToContent(0, p);
		AscTest.MoveCursorToParagraph(p, true);
		assert.strictEqual(getSelectionDocInfo.call(AscTest.Editor), null, "No selection returns null");

		AscTest.SelectParagraphRange(p, 4, 15);
		let info  = getSelectionDocInfo.call(AscTest.Editor);
		let stats = logicDocument.GetSelectionStatistics();
		assert.deepEqual(Object.assign({}, info), stats, "Returns the selection statistics");
	});
});
