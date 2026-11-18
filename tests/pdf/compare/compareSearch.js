/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	// Reading the text of a PDF needs the native PDF engine, which does not run headless.
	// These tests give the text lines directly and check the sentence split, the diff and the search.
	const api = window["Asc"]["PDFEditorApi"].prototype;

	function Lines(aTexts, nPage)
	{
		return aTexts.map(function(sText)
		{
			return {page : nPage || 0, text : sText};
		});
	}

	function Sentences(aTexts)
	{
		return api.getPdfSentences(Lines(aTexts));
	}

	function Texts(aSentences)
	{
		return aSentences.map(function(oSentence)
		{
			return oSentence.text;
		});
	}

	// An api object whose other file has the given lines (null = cannot be read) and whose open file has aOpenLines
	function CreateApi(aOpenLines, oOtherLines)
	{
		let oApi = Object.create(api);
		oApi.getPdfLines = function()
		{
			return aOpenLines;
		};
		oApi.getPdfLinesFromBytes = function(aBytes)
		{
			return oOtherLines[aBytes] !== undefined ? oOtherLines[aBytes] : null;
		};
		return oApi;
	}

	QUnit.module("PDF compare and search");

	QUnit.test("Sentences", function (assert)
	{
		assert.deepEqual(Texts(Sentences(["One. Two! Three? Four"])), ["One.", "Two!", "Three?", "Four"], "Split at . ! ? followed by a space");
		assert.deepEqual(Texts(Sentences(["This sentence goes", "over two lines. Next one."])),
			["This sentence goes over two lines.", "Next one."], "A sentence over a line break is one sentence");
		assert.deepEqual(Texts(Sentences(["Version 1.5 is out"])), ["Version 1.5 is out"], "No split inside a number");
		assert.deepEqual(Texts(Sentences(["Note: read this; then go"])), ["Note:", "read this;", "then go"], "Split at : and ;");
		assert.deepEqual(Texts(Sentences(["", "   "])), [], "Empty lines give no sentences");

		let aSentences = api.getPdfSentences([{page : 0, text : "Start of a"}, {page : 1, text : "sentence. Second."}]);
		assert.strictEqual(aSentences[0].page, 0, "A sentence keeps the page where it starts");
		assert.strictEqual(aSentences[1].page, 1, "The next sentence has its own page");

		assert.strictEqual(Sentences(["Sp ac es  inside"])[0].key, Sentences(["Spaces inside"])[0].key, "Spaces are ignored in the key");
		assert.strictEqual(Sentences(["a   b"])[0].text, "a b", "Spaces are collapsed in the shown text");
	});

	QUnit.test("Diff of sentences", function (assert)
	{
		let aA = Sentences(["One. Two. Three. Four."]);

		let oResult = api.comparePdfSentences(aA, Sentences(["One. Two. Three. Four."]));
		assert.deepEqual([oResult.removed, oResult.added, oResult.tooMany], [[], [], false], "Same text: no differences");

		oResult = api.comparePdfSentences(aA, Sentences(["One. Two changed. Three. Four. Five."]));
		assert.deepEqual(Texts(oResult.removed), ["Two."], "Changed sentence is removed");
		assert.deepEqual(Texts(oResult.added), ["Two changed.", "Five."], "Changed and new sentence are added");

		oResult = api.comparePdfSentences(aA, Sentences(["Two. Three."]));
		assert.deepEqual(Texts(oResult.removed), ["One.", "Four."], "Sentences at the start and end removed");
		assert.deepEqual(oResult.added, [], "Nothing added");

		oResult = api.comparePdfSentences(aA, Sentences(["One. Two.", "Three.", "Four."]));
		assert.deepEqual([oResult.removed, oResult.added], [[], []], "Other line wrapping is no difference");

		oResult = api.comparePdfSentences([], Sentences(["New. Text."]));
		assert.deepEqual([Texts(oResult.removed), Texts(oResult.added)], [[], ["New.", "Text."]], "Empty first file");
		oResult = api.comparePdfSentences(aA, []);
		assert.deepEqual([Texts(oResult.removed), Texts(oResult.added)], [["One.", "Two.", "Three.", "Four."], []], "Empty second file");
		oResult = api.comparePdfSentences([], []);
		assert.deepEqual([oResult.removed, oResult.added, oResult.tooMany], [[], [], false], "Both empty");

		oResult = api.comparePdfSentences(Sentences(["A. B. C. D."]), Sentences(["E. F. G. H."]), 7);
		assert.strictEqual(oResult.tooMany, true, "Stops after the limit of differences");
		oResult = api.comparePdfSentences(Sentences(["A. B. C. D."]), Sentences(["E. F. G. H."]), 8);
		assert.strictEqual(oResult.tooMany, false, "Within the limit");
		assert.strictEqual(oResult.removed.length + oResult.added.length, 8, "All sentences differ");

		oResult = api.comparePdfSentences(aA, Sentences(["One. Two. Three. Four."]));
		assert.deepEqual(Object.keys(oResult.removed.concat(oResult.added)[0] || {page : 0, text : ""}).sort(), ["page", "text"], "Results have page and text only");
	});

	QUnit.test("Diff is a shortest edit script", function (assert)
	{
		// The number of differences equals |A| + |B| - 2 * LCS, checked against a simple LCS on random texts
		function Lcs(aA, aB)
		{
			let aRow = new Array(aB.length + 1).fill(0);
			for (let i = 1; i <= aA.length; ++i)
			{
				let nPrev = 0;
				for (let j = 1; j <= aB.length; ++j)
				{
					let nTmp = aRow[j];
					aRow[j] = aA[i - 1].key === aB[j - 1].key ? nPrev + 1 : Math.max(aRow[j], aRow[j - 1]);
					nPrev = nTmp;
				}
			}
			return aRow[aB.length];
		}

		let nSeed = 7;
		function Random(nMax)
		{
			nSeed = (nSeed * 1103515245 + 12345) % 2147483648;
			return nSeed % nMax;
		}

		for (let nTest = 0; nTest < 30; ++nTest)
		{
			let aA = [], aB = [];
			for (let i = Random(20); i > 0; --i)
				aA.push("S" + Random(5) + ".");
			for (let i = Random(20); i > 0; --i)
				aB.push("S" + Random(5) + ".");
			let aSA = Sentences([aA.join(" ")]), aSB = Sentences([aB.join(" ")]);
			let oResult = api.comparePdfSentences(aSA, aSB);
			let nDiff = oResult.removed.length + oResult.added.length;
			assert.strictEqual(nDiff, aSA.length + aSB.length - 2 * Lcs(aSA, aSB), "Random case " + nTest + ": " + aA.join("") + " / " + aB.join(""));

			// the removed sentences are taken from A in order, the added ones from B in order
			let fIsSubsequence = function(aOut, aAll)
			{
				let nPos = 0;
				aAll.forEach(function(oSentence)
				{
					if (nPos < aOut.length && aOut[nPos].text === oSentence.text)
						++nPos;
				});
				return nPos === aOut.length;
			};
			assert.ok(fIsSubsequence(oResult.removed, aSA) && fIsSubsequence(oResult.added, aSB), "Random case " + nTest + ": results in document order");
		}
	});

	QUnit.test("Compare with another file", function (assert)
	{
		let aOpen = Lines(["First sentence. Second", "sentence."], 0).concat(Lines(["Third sentence."], 1));
		let oApi = CreateApi(aOpen, {"other" : Lines(["First sentence. Second sentence.", "Third sentence. Fourth."], 0), "bad" : null});

		let oResult = oApi.asc_ComparePdfText("other");
		assert.deepEqual(oResult, {removed : [], added : [{page : 0, text : "Fourth."}], tooMany : false}, "Only the new sentence");
		assert.strictEqual(oApi.asc_ComparePdfText("bad"), null, "Unreadable file gives null");
	});

	QUnit.test("Search in several files", function (assert)
	{
		let oApi = CreateApi([], {
			"a"   : Lines(["Hello World", "nothing here"], 0).concat(Lines(["hello again, w o r l d"], 2)),
			"b"   : Lines(["no hit"], 0),
			"bad" : null
		});
		let aFiles = [{name : "a.pdf", path : "C:/a.pdf", bytes : "a"}, {name : "b.pdf", path : "b.pdf", bytes : "b"}, {name : "c.pdf", path : "c.pdf", bytes : "bad"}];

		assert.deepEqual(oApi.asc_SearchInPdfFiles("world", aFiles), [
			{name : "a.pdf", path : "C:/a.pdf", page : 0, text : "Hello World"},
			{name : "a.pdf", path : "C:/a.pdf", page : 2, text : "hello again, w o r l d"},
			{name : "c.pdf", path : "c.pdf", page : -1, text : ""}
		], "Case and spaces are ignored; unreadable file has page -1");
		assert.deepEqual(oApi.asc_SearchInPdfFiles("HelloWorld", aFiles.slice(0, 2)).length, 1, "Spaces in the search text are ignored too");
		assert.deepEqual(oApi.asc_SearchInPdfFiles("   ", aFiles), [], "Empty search text finds nothing");
		assert.deepEqual(oApi.asc_SearchInPdfFiles("world", []), [], "No files");
	});
});
