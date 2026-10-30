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
});
