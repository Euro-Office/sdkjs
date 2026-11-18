/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	function createRun(text, reviewType)
	{
		let run = AscTest.CreateRun();
		run.AddText(text);
		if (undefined !== reviewType)
		{
			let info = new AscWord.ReviewInfo();
			info.UserName = "Alice";
			info.DateTime = 1000;
			run.SetReviewTypeWithInfo(reviewType, info);
		}
		return run;
	}

	function createParagraph()
	{
		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		for (let index = 0; index < arguments.length; ++index)
		{
			paragraph.AddToContentToEnd(arguments[index]);
		}
		return paragraph;
	}

	function count()
	{
		let stats = logicDocument.Statistics;
		stats.Reset();
		logicDocument.Content.forEach(function(element)
		{
			element.CollectDocumentStatistics(stats);
		});
		return stats;
	}

	QUnit.module("Word count as in LibreOffice");

	QUnit.test("En and em dashes separate words", function (assert)
	{
		AscTest.ClearDocument();
		createParagraph(createRun("Haus–Boot Rot—Blau gelb – grün Bindestrich-Wort"));

		let stats = count();
		assert.strictEqual(stats.Words, 7, "Haus, Boot, Rot, Blau, gelb, grün and the hyphenated word");
		assert.strictEqual(stats.SymbolsWOSpaces, 42, "Dashes count as symbols");
	});

	QUnit.test("Deleted and hidden text is not counted", function (assert)
	{
		AscTest.ClearDocument();
		let hidden = createRun("hidden words ");
		hidden.Pr.Vanish = true;
		hidden.Recalc_CompiledPr(true);
		createParagraph(createRun("one two "), createRun("gone away ", reviewtype_Remove), hidden, createRun("three"));
		createParagraph(createRun("only deleted", reviewtype_Remove));

		let stats = count();
		assert.strictEqual(stats.Words, 3, "Only the visible words count");
		assert.strictEqual(stats.Paragraphs, 1, "A paragraph with only deleted text is not counted");
	});

	QUnit.test("Words added and deleted with tracked changes have their own counters", function (assert)
	{
		AscTest.ClearDocument();
		createParagraph(
			createRun("Das "),
			createRun("alte ", reviewtype_Remove),
			createRun("neue grüne ", reviewtype_Add),
			createRun("Haus"),
			createRun("es", reviewtype_Add),
			createRun(" steht"),
			createRun(" nicht mehr da", reviewtype_Remove));

		let stats = count();
		assert.strictEqual(stats.AddedWords, 3, "neue, grüne and the added ending");
		assert.strictEqual(stats.DeletedWords, 4, "alte, nicht, mehr, da");
		assert.strictEqual(stats.Words, 5, "Das neue grüne Hauses steht");
	});

	QUnit.test("Footnotes are counted", function (assert)
	{
		AscTest.ClearDocument();
		let footnote = logicDocument.Footnotes.CreateFootnote();
		footnote.GetElement(0).AddToContentToEnd(createRun("Quelle siehe oben"));

		let run = createRun("Text");
		run.AddToContent(run.GetElementsCount(), new AscWord.CRunFootnoteReference(footnote));
		createParagraph(run);

		assert.strictEqual(count().Words, 4, "One word of the text and three of the footnote");
	});

	QUnit.test("Comments are counted with their replies", function (assert)
	{
		AscTest.ClearDocument();
		let data = new AscCommon.CCommentData();
		data.m_sText = "Bitte prüfen–danke";
		let reply = new AscCommon.CCommentData();
		reply.m_sText = "Erledigt.\nGruß";
		data.Add_Reply(reply);
		logicDocument.Comments.Add(new AscCommon.CComment(logicDocument.Comments, data));

		let stats = logicDocument.Statistics;
		stats.Reset();
		logicDocument.CollectCommentStatistics(stats);
		assert.strictEqual(stats.Words, 5, "Bitte, prüfen, danke, Erledigt., Gruß");
		assert.strictEqual(stats.Paragraphs, 3, "Each line counts as a paragraph");
	});

	QUnit.test("Every Chinese or Japanese character is a word, Korean is separated by spaces", function (assert)
	{
		AscTest.ClearDocument();
		// "Hello world." in Chinese, "Euro-Office" plus Hiragana and Katakana, and Korean
		createParagraph(createRun("\u4F60\u597D\u4E16\u754C\u3002 Euro\u3067\u30C6\u30B9\u30C8 \uC548\uB155\uD558\uC138\uC694 \uC138\uACC4"));

		let stats = count();
		assert.strictEqual(stats.Words, 12, "4 Chinese and the full stop as in LibreOffice, Euro, 4 Japanese, 2 Korean words");
	});

	QUnit.test("Chinese characters inserted and deleted with tracked changes are counted one by one", function (assert)
	{
		AscTest.ClearDocument();
		createParagraph(createRun("\u4F60\u597D", reviewtype_Add), createRun("\u4E16\u754C\u4EBA", reviewtype_Remove));

		let stats = count();
		assert.strictEqual(stats.AddedWords, 2, "Two inserted characters");
		assert.strictEqual(stats.DeletedWords, 3, "Three deleted characters");
	});

	QUnit.test("Plain text is counted like document text", function (assert)
	{
		let stats = logicDocument.Statistics;
		stats.Reset();
		stats.AddText("  zwei Wörter — ");
		assert.strictEqual(stats.Words, 2, "Spaces, no-break spaces and dashes separate words");
		assert.strictEqual(stats.SymbolsWhSpaces, 16, "All symbols with spaces");
		assert.strictEqual(stats.SymbolsWOSpaces, 11, "Symbols without spaces");
	});

	QUnit.test("A selection uses the same rules and counts only the notes it selects itself", function (assert)
	{
		AscTest.ClearDocument();
		let footnote = logicDocument.Footnotes.CreateFootnote();
		footnote.GetElement(0).AddToContentToEnd(createRun("Quelle siehe oben"));

		let run = createRun("drei");
		run.AddToContent(run.GetElementsCount(), new AscWord.CRunFootnoteReference(footnote));
		createParagraph(createRun("eins–zwei "), createRun("weg ", reviewtype_Remove), run);

		logicDocument.SelectAll();
		let stats = logicDocument.GetSelectionStatistics();
		assert.strictEqual(stats.WordsCount, 3, "eins, zwei, drei: the dash separates, deleted text and the footnote text are not counted");
		assert.strictEqual(count().Words, 6, "The document total still includes the footnote");
	});

	// the full count as the status bar does it: text, then the drawings per page, then headers and footers
	function countWithLayout(includeHdrFtr)
	{
		AscTest.Recalculate();
		let stats = logicDocument.Statistics;
		stats.Reset();
		stats.IncludeHdrFtr = !!includeHdrFtr;
		logicDocument.Content.forEach(function(element)
		{
			element.CollectDocumentStatistics(stats);
		});
		for (let page = 0; page < logicDocument.Pages.length; ++page)
		{
			logicDocument.DrawingObjects.documentStatistics(page, stats);
		}
		if (includeHdrFtr)
			logicDocument.CollectHdrFtrStatistics(stats);
		stats.IncludeHdrFtr = false;
		return stats;
	}

	function createTextBox(text, isInline, docContent)
	{
		let drawing    = new AscWord.ParaDrawing(50, 20, null, logicDocument.GetDrawingDocument(), docContent || logicDocument, null);
		let shapeTrack = new AscFormat.NewShapeTrack("rect", 0, 0, logicDocument.theme, null, null, null, 0);
		shapeTrack.track({}, 50, 20);
		let shape = shapeTrack.getShape(true, logicDocument.GetDrawingDocument(), null);
		shape.setBDeleted(false);
		shape.setParent(drawing);
		drawing.Set_GraphicObject(shape);
		if (!isInline)
		{
			drawing.Set_DrawingType(drawing_Anchor);
			drawing.Set_WrappingType(WRAPPING_TYPE_NONE);
			drawing.Set_Distance(0, 0, 0, 0);
		}
		shape.createTextBoxContent();
		shape.getDocContent().GetElement(0).AddToContentToEnd(createRun(text));
		return drawing;
	}

	QUnit.test("A deleted or hidden part inside an inserted word does not split it", function (assert)
	{
		AscTest.ClearDocument();
		let hidden = createRun("y", reviewtype_Add);
		hidden.Pr.Vanish = true;
		hidden.Recalc_CompiledPr(true);
		createParagraph(createRun("neu", reviewtype_Add), createRun("x", reviewtype_Remove), createRun("e", reviewtype_Add),
			createRun(" alt"), createRun("er", reviewtype_Add), hidden, createRun("e", reviewtype_Add));

		let stats = count();
		assert.strictEqual(stats.Words, 2, "neue, altere");
		assert.strictEqual(stats.AddedWords, 2, "neue and the inserted ending of altere, each once");
		assert.strictEqual(stats.DeletedWords, 1, "x");
	});

	QUnit.test("Deleted hidden text is not counted as deleted", function (assert)
	{
		AscTest.ClearDocument();
		let hidden = createRun("verborgen gelöscht ", reviewtype_Remove);
		hidden.Pr.Vanish = true;
		hidden.Recalc_CompiledPr(true);
		createParagraph(createRun("eins "), hidden, createRun("zwei ", reviewtype_Remove));

		let stats = count();
		assert.strictEqual(stats.Words, 1, "eins");
		assert.strictEqual(stats.DeletedWords, 1, "only the visible deleted word");
	});

	QUnit.test("A text box in a footnote is counted once", function (assert)
	{
		[false, true].forEach(function(isInline)
		{
			AscTest.ClearDocument();
			let footnote = logicDocument.Footnotes.CreateFootnote();
			footnote.LogicDocument = logicDocument;
			let notePara = footnote.GetElement(0);
			notePara.LogicDocument = logicDocument;
			let noteRun = createRun("foot note text");
			noteRun.AddToContent(noteRun.GetElementsCount(), createTextBox("alpha beta gamma delta", isInline));
			notePara.AddToContentToEnd(noteRun);

			let run = createRun("one two three");
			run.AddToContent(run.GetElementsCount(), new AscWord.CRunFootnoteReference(footnote));
			createParagraph(run);

			let stats = countWithLayout(false);
			let kind  = isInline ? "inline" : "floating";
			assert.strictEqual(stats.Words, 10, kind + ": 3 words of the text, 3 of the note and 4 of its text box");
			assert.strictEqual(stats.Paragraphs, 3, kind + ": text, note and text box paragraph");
		});
	});

	QUnit.test("Headers and footers count only with the option, once each", function (assert)
	{
		AscTest.ClearDocument();
		createParagraph(createRun("body text"));
		let header = AscTest.CreateDefaultHeader(logicDocument.GetFinalSectPr());
		header.GetElement(0).LogicDocument = logicDocument;
		let headerRun = createRun("Kopfzeile mit Text");
		headerRun.AddToContent(headerRun.GetElementsCount(), createTextBox("im Kasten", true, header));
		header.GetElement(0).AddToContentToEnd(headerRun);

		assert.strictEqual(countWithLayout(false).Words, 2, "without the option only the body counts");
		assert.strictEqual(countWithLayout(true).Words, 7, "with the option the header text and its text box count too");
		AscTest.GetFinalSection().Set_Header_Default(null);
	});
});
