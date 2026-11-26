/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	// Reading the text of a PDF needs the native PDF engine, which does not run here (no fonts, file:// page).
	// These tests give the text lines directly and check the sentence split, the diff, the search and how the
	// files are opened and closed in the engine (with a stand-in for it).
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
		oApi.getPdfLines = function(fCallback)
		{
			fCallback(aOpenLines);
		};
		oApi.getPdfLinesFromBytes = function(aBytes, fCallback)
		{
			fCallback(oOtherLines[aBytes] !== undefined ? oOtherLines[aBytes] : null);
		};
		return oApi;
	}

	// Calls a function of the api that takes a callback as its last parameter and gives back what the callback got
	function CallSync(oApi, sName)
	{
		let aArgs = Array.prototype.slice.call(arguments, 2);
		let oResult = "callback not called";
		aArgs.push(function(oValue)
		{
			oResult = oValue;
		});
		oApi[sName].apply(oApi, aArgs);
		return oResult;
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
		assert.deepEqual([oResult.removed, oResult.added, oResult.coarse], [[], [], false], "Same text: no differences");

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
		assert.deepEqual([oResult.removed, oResult.added, oResult.coarse], [[], [], false], "Both empty");

		oResult = api.comparePdfSentences(Sentences(["A. B. C. D."]), Sentences(["E. F. G. H."]), 7);
		assert.strictEqual(oResult.coarse, true, "Over the limit of differences the coarse comparison is used");
		assert.deepEqual([Texts(oResult.removed), Texts(oResult.added)], [["A.", "B.", "C.", "D."], ["E.", "F.", "G.", "H."]], "and still lists the differences");
		oResult = api.comparePdfSentences(Sentences(["A. B. C. D."]), Sentences(["E. F. G. H."]), 8);
		assert.strictEqual(oResult.coarse, false, "Within the limit");
		assert.strictEqual(oResult.removed.length + oResult.added.length, 8, "All sentences differ");

		oResult = api.comparePdfSentences(aA, Sentences(["One. Two. Three. Four."]));
		assert.deepEqual(Object.keys(oResult.removed.concat(oResult.added)[0] || {page : 0, text : ""}).sort(), ["page", "text"], "Results have page and text only");
	});

	QUnit.test("Sentences in Chinese and Japanese", function (assert)
	{
		// made-up text; full-width sentence ends have no space after them
		assert.deepEqual(Texts(Sentences(["今天天气很好。我们去公园散步吧！你想一起来吗？"])),
			["今天天气很好。", "我们去公园散步吧！", "你想一起来吗？"], "Split at 。！？");
		assert.deepEqual(Texts(Sentences(["第一个句子．第二个句子｡第三个"])), ["第一个句子．", "第二个句子｡", "第三个"], "Split at ．and ｡");
		assert.deepEqual(Texts(Sentences(["「こんにちは。」と先生は言いました。元気ですか？！"])),
			["「こんにちは。」", "と先生は言いました。", "元気ですか？！"], "A closing bracket stays with its sentence, ？！ is one end");
		assert.deepEqual(Texts(Sentences(["这是一个很长的句子，", "分成两行。下一句。"])),
			["这是一个很长的句子，分成两行。", "下一句。"], "A sentence over a line break is one sentence, without a space in CJK text");
		assert.deepEqual(Texts(Sentences(["English first. 中文第二。Third one"])), ["English first.", "中文第二。", "Third one"], "Mixed text");

		let aBase = [];
		for (let i = 0; i < 20; ++i)
			aBase.push("这是第" + i + "个测试句子。");
		let aChanged = aBase.slice();
		aChanged[7] = "这是被修改过的句子。";
		let oResult = api.comparePdfSentences(Sentences([aBase.join("")]), Sentences([aChanged.join("")]));
		assert.deepEqual([Texts(oResult.removed), Texts(oResult.added)], [["这是第7个测试句子。"], ["这是被修改过的句子。"]],
			"One changed sentence in a CJK text marks only that sentence");
	});

	QUnit.test("Very different files still give a result", function (assert)
	{
		let aA = [], aB = [];
		for (let i = 0; i < 2000; ++i)
		{
			aA.push("Sentence " + i + ".");
			aB.push(i < 1501 ? "Changed sentence " + i + "." : "Sentence " + i + ".");
		}
		// 1501 changed sentences are 3002 edits, more than the default limit of 3000
		let oResult = api.comparePdfSentences(Sentences([aA.join(" ")]), Sentences([aB.join(" ")]));
		assert.strictEqual(oResult.coarse, true, "Coarse comparison over the limit");
		assert.strictEqual(oResult.removed.length, 1501, "All changed sentences are removed");
		assert.strictEqual(oResult.added.length, 1501, "All changed sentences are added");
		assert.deepEqual([oResult.removed[0].text, oResult.added[1500].text], ["Sentence 0.", "Changed sentence 1500."], "In document order");

		oResult = api.comparePdfSentences(Sentences(["A. B. C."]), Sentences(["C. A. B."]), 0);
		assert.deepEqual([oResult.coarse, oResult.removed, oResult.added], [true, [], []], "Coarse: moved sentences are not listed");
		oResult = api.comparePdfSentences(Sentences(["X. X. Y."]), Sentences(["X. Y. Z."]), 0);
		assert.deepEqual([Texts(oResult.removed), Texts(oResult.added)], [["X."], ["Z."]], "Coarse: a sentence counts as often as it is there");
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
		let oApi = CreateApi(aOpen, {"other" : Lines(["First sentence. Second sentence.", "Third sentence. Fourth."], 0), "bad" : null, "scan" : Lines(["", " "], 0)});

		assert.deepEqual(CallSync(oApi, "asc_ComparePdfText", "other"),
			{removed : [], added : [{page : 0, text : "Fourth."}], coarse : false, openHasText : true, otherHasText : true}, "Only the new sentence");
		assert.strictEqual(CallSync(oApi, "asc_ComparePdfText", "bad"), null, "Unreadable file gives null");

		let oResult = CallSync(oApi, "asc_ComparePdfText", "scan");
		assert.deepEqual([oResult.openHasText, oResult.otherHasText], [true, false], "Other file without text (scanned)");
		oResult = CallSync(CreateApi([], {"other" : aOpen}), "asc_ComparePdfText", "other");
		assert.deepEqual([oResult.openHasText, oResult.otherHasText], [false, true], "Open file without text");
	});

	QUnit.test("Search in several files", function (assert)
	{
		let oApi = CreateApi([], {
			"a"    : Lines(["Hello World", "nothing here"], 0).concat(Lines(["hello again, w o r l d"], 2)),
			"b"    : Lines(["no hit"], 0),
			"bad"  : null,
			"scan" : []
		});
		let aFiles = [{name : "a.pdf", path : "C:/a.pdf", bytes : "a"}, {name : "b.pdf", path : "b.pdf", bytes : "b"}, {name : "c.pdf", path : "c.pdf", bytes : "bad"},
			{name : "d.pdf", path : "d.pdf", bytes : "scan"}];

		assert.deepEqual(CallSync(oApi, "asc_SearchInPdfFiles", "world", aFiles), [
			{name : "a.pdf", path : "C:/a.pdf", page : 0, text : "Hello World"},
			{name : "a.pdf", path : "C:/a.pdf", page : 2, text : "hello again, w o r l d"},
			{name : "c.pdf", path : "c.pdf", page : -1, text : ""},
			{name : "d.pdf", path : "d.pdf", page : -2, text : ""}
		], "Case and spaces are ignored; unreadable file has page -1, file without text page -2");
		assert.deepEqual(CallSync(oApi, "asc_SearchInPdfFiles", "HelloWorld", aFiles.slice(0, 2)).length, 1, "Spaces in the search text are ignored too");
		assert.deepEqual(CallSync(oApi, "asc_SearchInPdfFiles", "   ", aFiles), [], "Empty search text finds nothing");
		assert.deepEqual(CallSync(oApi, "asc_SearchInPdfFiles", "world", []), [], "No files");
	});

	QUnit.test("Only one reading at a time", function (assert)
	{
		let oApi = CreateApi([], {});
		let fPending = null;
		oApi.getPdfLinesFromBytes = function(aBytes, fCallback)
		{
			fPending = fCallback;
		};
		let aResults = [];
		assert.strictEqual(oApi.asc_ComparePdfText("x", function(oResult) { aResults.push(oResult); }), true, "First comparison starts");
		assert.strictEqual(oApi.asc_SearchInPdfFiles("a", [{name : "a", path : "a", bytes : "a"}], function() { aResults.push("search"); }), false, "A search is refused while the comparison runs");
		assert.strictEqual(oApi.asc_ComparePdfText("y", function() { aResults.push("second"); }), false, "A second comparison is refused");
		fPending(null);
		assert.deepEqual(aResults, [null], "Only the first callback is called");
		assert.strictEqual(oApi.asc_ComparePdfText("z", function() {}), true, "After the end a new comparison starts");
	});

	QUnit.test("Busy status for the UI", function (assert)
	{
		let oApi = CreateApi([], {});
		let fPending = null;
		oApi.getPdfLinesFromBytes = function(aBytes, fCallback)
		{
			fPending = fCallback;
		};
		assert.strictEqual(oApi.asc_IsPdfTextReading(), false, "Nothing runs at the start");
		oApi.asc_SearchInPdfFiles("a", [{name : "a", path : "a", bytes : "a"}], function() {});
		assert.strictEqual(oApi.asc_IsPdfTextReading(), true, "A search runs");
		let isCalled = false;
		assert.strictEqual(oApi.asc_SearchInPdfFiles("b", [{name : "b", path : "b", bytes : "b"}], function() { isCalled = true; }), false,
			"A second search returns false");
		fPending([]);
		assert.strictEqual(isCalled, false, "The callback of the refused search is not called");
		assert.strictEqual(oApi.asc_IsPdfTextReading(), false, "Nothing runs after the end");
	});

	QUnit.test("Search results are capped", function (assert)
	{
		let oApi = CreateApi([], {
			"a" : Lines(["hit 1", "hit 2", "other", "hit 3"], 0),
			"b" : Lines(["hit 4", "hit 5"], 1),
			"c" : Lines(["hit 6"], 0)
		});
		let aRead = [];
		let fOldRead = oApi.getPdfLinesFromBytes;
		oApi.getPdfLinesFromBytes = function(aBytes, fCallback)
		{
			aRead.push(aBytes);
			fOldRead.call(this, aBytes, fCallback);
		};
		let aFiles = [{name : "a", path : "a", bytes : "a"}, {name : "b", path : "b", bytes : "b"}, {name : "c", path : "c", bytes : "c"}];
		let aResult = null, oInfo = null;
		oApi.asc_SearchInPdfFiles("hit", aFiles, function(aFound, oFoundInfo) { aResult = aFound; oInfo = oFoundInfo; }, 4);
		assert.deepEqual(Texts(aResult), ["hit 1", "hit 2", "hit 3", "hit 4"], "Stops at the limit");
		assert.deepEqual(oInfo, {capped : true, maxResults : 4, searchedFiles : 2}, "Tells that the result is capped");
		assert.deepEqual(aRead, ["a", "b"], "The files after the limit are not read");

		oApi.asc_SearchInPdfFiles("hit", aFiles, function(aFound, oFoundInfo) { aResult = aFound; oInfo = oFoundInfo; });
		assert.strictEqual(aResult.length, 6, "All hits within the default limit");
		assert.deepEqual(oInfo, {capped : false, maxResults : 500, searchedFiles : 3}, "Not capped");

		let aMany = [];
		for (let i = 0; i < 600; ++i)
			aMany.push("hit " + i);
		oApi = CreateApi([], {"many" : Lines(aMany, 0)});
		oApi.asc_SearchInPdfFiles("hit", [{name : "m", path : "m", bytes : "many"}], function(aFound, oFoundInfo) { aResult = aFound; oInfo = oFoundInfo; });
		assert.deepEqual([aResult.length, oInfo.capped, oInfo.maxResults], [500, true, 500], "Default limit of 500 hits");
	});

	QUnit.test("Search reads one file at a time", function (assert)
	{
		let aLog = [];
		let oApi = CreateApi([], {"A" : Lines(["find me"], 0), "B" : Lines(["nothing"], 0), "C" : Lines(["find me too"], 3)});
		let fOldRead = oApi.getPdfLinesFromBytes;
		oApi.getPdfLinesFromBytes = function(aBytes, fCallback)
		{
			aLog.push("read " + aBytes);
			fOldRead.call(this, aBytes, fCallback);
			aLog.push("released " + aBytes);
		};
		function LazyFile(sName, sBytes)
		{
			return {name : sName, path : sName, getBytes : function(fDone)
			{
				aLog.push("load " + sName);
				fDone(sBytes);
			}};
		}
		let aResult = null;
		oApi.asc_SearchInPdfFiles("find me", [LazyFile("a", "A"), LazyFile("b", "B"), LazyFile("x", null), LazyFile("c", "C")], function(aFound) { aResult = aFound; });
		assert.deepEqual(aLog, ["load a", "read A", "released A", "load b", "read B", "released B", "load x", "load c", "read C", "released C"],
			"Each file is loaded only after the last one is read and its call has ended");
		assert.deepEqual(aResult, [
			{name : "a", path : "a", page : 0, text : "find me"},
			{name : "x", path : "x", page : -1, text : ""},
			{name : "c", path : "c", page : 3, text : "find me too"}
		], "Hits of the lazy files; a file that can't be loaded has page -1");
	});

	QUnit.test("Search reads one file at a time (loading in the background)", function (assert)
	{
		let done = assert.async();
		let aLog = [];
		let nLoading = 0, nMaxLoading = 0;
		let oApi = CreateApi([], {"A" : Lines(["word"], 0), "B" : Lines(["word"], 1)});
		let fOldRead = oApi.getPdfLinesFromBytes;
		oApi.getPdfLinesFromBytes = function(aBytes, fCallback)
		{
			aLog.push("read " + aBytes);
			setTimeout(function()
			{
				--nLoading;
				fOldRead.call(oApi, aBytes, fCallback);
			}, 5);
		};
		function LazyFile(sName, sBytes)
		{
			return {name : sName, path : sName, getBytes : function(fDone)
			{
				aLog.push("load " + sName);
				nMaxLoading = Math.max(nMaxLoading, ++nLoading);
				setTimeout(function() { fDone(sBytes); }, 5);
			}};
		}
		let isStarted = oApi.asc_SearchInPdfFiles("word", [LazyFile("a", "A"), LazyFile("b", "B")], function(aFound)
		{
			assert.deepEqual(aLog, ["load a", "read A", "load b", "read B"], "In order");
			assert.strictEqual(nMaxLoading, 1, "Never two files in memory at once");
			assert.strictEqual(aFound.length, 2, "Both hits");
			done();
		});
		assert.strictEqual(isStarted, true, "The search starts");
	});

	QUnit.test("Picking files without reading them yet (desktop app)", function (assert)
	{
		let oOldDesktop = window["AscDesktopEditor"];
		let aLoaded = [];
		window["AscDesktopEditor"] = {
			"OpenFilenameDialog" : function(sFilter, isMulti, fCallback) { fCallback(["C:/docs/one.pdf", "D:\\two.pdf"]); },
			"loadLocalFile"      : function(sPath, fCallback)
			{
				aLoaded.push(sPath);
				fCallback(sPath === "D:\\two.pdf" ? null : [37, 80]);
			}
		};
		try
		{
			let aFiles = null;
			api.asc_PickPdfFiles(true, function(aPicked) { aFiles = aPicked; }, true);
			assert.deepEqual(aFiles.map(function(oFile) { return [oFile.name, oFile.path, oFile.bytes]; }),
				[["one.pdf", "C:/docs/one.pdf", undefined], ["two.pdf", "D:\\two.pdf", undefined]], "Lazy: names and paths only");
			assert.deepEqual(aLoaded, [], "Lazy: nothing is read yet");
			let aBytes = "not called";
			aFiles[0].getBytes(function(aData) { aBytes = aData; });
			assert.deepEqual([aLoaded, Array.from(aBytes)], [["C:/docs/one.pdf"], [37, 80]], "getBytes reads the file");
			aFiles[1].getBytes(function(aData) { aBytes = aData; });
			assert.strictEqual(aBytes, null, "getBytes gives null for a file that can't be read");

			aLoaded = [];
			api.asc_PickPdfFiles(true, function(aPicked) { aFiles = aPicked; });
			assert.deepEqual(aLoaded, ["C:/docs/one.pdf", "D:\\two.pdf"], "Not lazy: all files are read");
			assert.deepEqual(aFiles.map(function(oFile) { return [oFile.name, Array.from(oFile.bytes)]; }), [["one.pdf", [37, 80]]],
				"Not lazy: bytes of the readable files");
		}
		finally
		{
			window["AscDesktopEditor"] = oOldDesktop;
		}
	});

	QUnit.test("Pages whose fonts are still loading are read again", function (assert)
	{
		let done = assert.async();
		let nCalls = 0;
		let oFile = {
			pages        : [{}, {}, {}],
			getPageLines : function(nPage)
			{
				++nCalls;
				// page 2 has no text in the first two rounds (its fonts are loading)
				if (1 === nPage && nCalls < 5)
					return null;
				return ["Page " + (nPage + 1)];
			}
		};
		let aEvents = [];
		api.readPdfLines(oFile, function() { aEvents.push("before"); }, function() { aEvents.push("after"); }, function(aLines)
		{
			assert.deepEqual(aLines, [{page : 0, text : "Page 1"}, {page : 1, text : "Page 2"}, {page : 2, text : "Page 3"}], "All pages once their fonts are there");
			assert.strictEqual(nCalls, 5, "Pages with text are not read again");
			assert.deepEqual(aEvents, ["before", "after", "before", "after", "before", "after"], "Before and after each round");

			let oNeverFile = {pages : [{}, {}], getPageLines : function(nPage) { return nPage ? null : ["Only"]; }};
			api.readPdfLines(oNeverFile, null, null, function(aLines)
			{
				assert.deepEqual(aLines, [{page : 0, text : "Only"}], "Stops after the time limit with the pages it has");

				let oBrokenFile = {pages : [{}], getPageLines : function() { throw new Error("engine"); }};
				let aEvents2 = [];
				api.readPdfLines(oBrokenFile, null, function() { aEvents2.push("after"); }, function(aLines)
				{
					assert.strictEqual(aLines, null, "An engine error gives null");
					assert.deepEqual(aEvents2, ["after"], "fAfter is called after an error too");
					done();
				});
			}, 120);
		});
	});

	QUnit.test("The other file is closed and the engine gets a new font storage", function (assert)
	{
		// The engine frees its shared font storage when any file is closed; opening a file makes a new one.
		// Without a new one the open document crashed on its next drawing ("remainder by zero").
		let done = assert.async();
		let oViewer = window["AscViewer"];
		let fOldCreate = oViewer.createFile, fOldDrawingFile = oViewer["CDrawingFile"], oOldGlobal = window["drawingFile"];
		let aLog = [];
		let nKeeper = 0;
		function CKeeper()
		{
			this.id = "empty" + (++nKeeper);
		}
		CKeeper.prototype["loadFromData"] = function(aBytes)
		{
			aLog.push("open " + this.id + (String.fromCharCode.apply(null, aBytes.slice(0, 5)) === "%PDF-" ? " (pdf)" : ""));
			window["drawingFile"] = this;
			return 0;
		};
		CKeeper.prototype["close"] = function()
		{
			aLog.push("close " + this.id);
			window["drawingFile"] = null;
		};
		let oOpenNative = {id : "open document"};
		let oOtherNative = {id : "other"};
		function Restore()
		{
			oViewer.createFile = fOldCreate;
			oViewer["CDrawingFile"] = fOldDrawingFile;
			window["drawingFile"] = oOldGlobal;
		}
		oViewer["CDrawingFile"] = CKeeper;
		oViewer.createFile = function(aBytes)
		{
			aLog.push("open other");
			window["drawingFile"] = oOtherNative;
			if ("bad" === aBytes)
			{
				// createFile closes a file it can't read
				aLog.push("close other");
				window["drawingFile"] = null;
				return null;
			}
			return {
				nativeFile   : oOtherNative,
				pages        : [{}],
				getPageLines : function()
				{
					aLog.push("read with " + (window["drawingFile"] && window["drawingFile"].id));
					return ["Text."];
				},
				close        : function()
				{
					aLog.push("close other");
					window["drawingFile"] = null;
				}
			};
		};
		window["drawingFile"] = oOpenNative;
		try
		{
			api.getPdfLinesFromBytes("good", function(aLines)
			{
				assert.deepEqual(aLines, [{page : 0, text : "Text."}], "Lines of the other file");
				assert.deepEqual(aLog, ["open other", "read with other", "close other", "open empty1 (pdf)"], "An empty PDF is opened after the other file is closed");
				assert.strictEqual(window["drawingFile"], oOpenNative, "The engine's current file is the open document again");
				aLog = [];
				api.getPdfLinesFromBytes("bad", function(aLines)
				{
					assert.strictEqual(aLines, null, "Unreadable file gives null");
					assert.deepEqual(aLog, ["open other", "close other", "close empty1", "open empty2 (pdf)"], "Also after a file that can't be read; only one empty PDF stays open");
					assert.strictEqual(window["drawingFile"], oOpenNative, "The engine's current file is the open document again");
					Restore();
					done();
				});
			});
		}
		catch (e)
		{
			Restore();
			assert.ok(false, "Exception: " + e);
			done();
		}
	});
});
