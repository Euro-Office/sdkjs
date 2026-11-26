/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	// The merge itself (pages, annotations and form fields of the other PDF) is done by
	// CPDFDoc.MergePagesBinary with the native PDF engine, which does not run headless.
	// These tests check the API wrapper around it: one history action, the insert position
	// and the cases where nothing is inserted.
	function CreateDoc(nPagesCount, isMergeOk, isMergeLocked)
	{
		let oDoc = {
			pagesCount   : nPagesCount,
			actions      : [],
			merges       : [],
			inAction     : false,
			mergedPagesData : [{mergeName : "Merged_0"}],
			Viewer       : {
				navigated  : [],
				thumbnails : {
					needResize : true,
					resized    : 0,
					isNeedResize : function()
					{
						return this.needResize;
					},
					resize     : function()
					{
						this.resized++;
						this.needResize = false;
					}
				},
				navigateToPage : function(nPage)
				{
					this.navigated.push(nPage);
				}
			},
			GetPagesCount: function()
			{
				return this.pagesCount;
			},
			IsMergeLock  : function()
			{
				return !!isMergeLocked;
			},
			DoAction     : function(fAction, nDescription, oThis)
			{
				this.actions.push(nDescription);
				this.inAction = true;
				let result = fAction.call(oThis);
				this.inAction = false;
				return result;
			},
			MergePagesBinary: function(nPos, aBytes)
			{
				this.merges.push({pos : nPos, bytes : aBytes, inAction : this.inAction});
				// like the real merge: the file is recorded in the history before the engine reads it
				let oHistory = AscCommon.History;
				oHistory.Points[oHistory.Index].Items.push("StartMerge", "PartMerge", "EndMerge");
				this.mergedPagesData.push({mergeName : "Merged_1"});
				if (isMergeOk)
					oHistory.Points[oHistory.Index].Items.push("AddPage");
				return isMergeOk;
			}
		};
		return oDoc;
	}

	function CreateApi(oDoc)
	{
		return {
			getPDFDoc: function()
			{
				return oDoc;
			},
			asc_InsertPagesFromBytes  : AscCommon.PDFEditorApi.prototype.asc_InsertPagesFromBytes,
			asc_GetInsertPagesMaxSize : AscCommon.PDFEditorApi.prototype.asc_GetInsertPagesMaxSize
		};
	}

	// A history with one open point that already has an item
	function WithHistory(fTest)
	{
		let oOldHistory = AscCommon.History;
		let oHistory = AscCommon.History = {Index : 0, Points : [{Items : ["Earlier"]}]};
		try
		{
			fTest(oHistory);
		}
		finally
		{
			AscCommon.History = oOldHistory;
		}
	}

	function InsertPages(oDoc, nPos, aBytes)
	{
		let result;
		WithHistory(function()
		{
			result = AscCommon.PDFEditorApi.prototype.asc_InsertPagesFromBytes.call(CreateApi(oDoc), nPos, aBytes);
		});
		return result;
	}

	function ToBytes(sText)
	{
		let aResult = new Uint8Array(sText.length);
		for (let i = 0; i < sText.length; i++)
			aResult[i] = sText.charCodeAt(i);
		return aResult;
	}

	let aBytes = ToBytes("%PDF-1.7\n1 0 obj\n<< /Type /Catalog >>\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
	let Result = Asc.c_oAscPdfInsertPagesResult;

	QUnit.module("PDF: insert the pages of another PDF");

	QUnit.test("All pages are inserted in one undo step", function (assert)
	{
		let oDoc = CreateDoc(8, true);
		assert.strictEqual(InsertPages(oDoc, 2, aBytes), Result.Ok, "Insert reports success");
		assert.strictEqual(oDoc.actions.length, 1, "One history action");
		assert.strictEqual(oDoc.actions[0], AscDFH.historydescription_Pdf_AddPage, "The action is described as adding pages");
		assert.strictEqual(oDoc.merges.length, 1, "The file is merged once");
		assert.strictEqual(oDoc.merges[0].inAction, true, "The merge runs inside the action");
		assert.strictEqual(oDoc.merges[0].pos, 2, "The pages go to the given position");
		assert.strictEqual(oDoc.merges[0].bytes, aBytes, "The bytes are passed on unchanged");
	});

	QUnit.test("Position", function (assert)
	{
		let oDoc = CreateDoc(8, true);
		InsertPages(oDoc, 0, aBytes);
		assert.strictEqual(oDoc.merges[0].pos, 0, "Before the first page");

		oDoc = CreateDoc(8, true);
		InsertPages(oDoc, 8, aBytes);
		assert.strictEqual(oDoc.merges[0].pos, 8, "After the last page");

		oDoc = CreateDoc(8, true);
		InsertPages(oDoc, undefined, aBytes);
		assert.strictEqual(oDoc.merges[0].pos, 8, "No position: at the end");

		oDoc = CreateDoc(8, true);
		InsertPages(oDoc, 20, aBytes);
		assert.strictEqual(oDoc.merges[0].pos, 8, "Position after the end: at the end");

		oDoc = CreateDoc(8, true);
		InsertPages(oDoc, -1, aBytes);
		assert.strictEqual(oDoc.merges[0].pos, 8, "Negative position: at the end");
	});

	QUnit.test("Nothing is inserted", function (assert)
	{
		let oDoc = CreateDoc(8, true);
		assert.strictEqual(InsertPages(oDoc, 2, new Uint8Array(0)), Result.NoFile, "Empty file");
		assert.strictEqual(InsertPages(oDoc, 2, null), Result.NoFile, "No file");
		assert.strictEqual(oDoc.actions.length, 0, "No history action without a file");

		oDoc = CreateDoc(8, true, true);
		assert.strictEqual(InsertPages(oDoc, 2, aBytes), Result.Locked, "Pages locked by another user");
		assert.strictEqual(oDoc.actions.length, 0, "No history action while locked");

		oDoc = CreateDoc(8, false);
		assert.strictEqual(InsertPages(oDoc, 2, aBytes), Result.Damaged, "The engine could not read the file");

		assert.strictEqual(AscCommon.PDFEditorApi.prototype.asc_InsertPagesFromBytes.call(CreateApi(null), 2, aBytes), Result.NoFile, "No document");
	});

	QUnit.test("A file the engine can't read leaves no history record", function (assert)
	{
		WithHistory(function(oHistory)
		{
			let oDoc = CreateDoc(8, false);
			let nResult = AscCommon.PDFEditorApi.prototype.asc_InsertPagesFromBytes.call(CreateApi(oDoc), 2, aBytes);
			assert.strictEqual(nResult, Result.Damaged, "Insert reports failure");
			assert.deepEqual(oHistory.Points[0].Items, ["Earlier"], "The merge record is dropped, earlier items stay");
			assert.deepEqual(oDoc.mergedPagesData, [{mergeName : "Merged_0"}], "The merge data is dropped, the earlier merge stays");
			assert.deepEqual(oDoc.Viewer.navigated, [], "No navigation");
			assert.strictEqual(oDoc.Viewer.thumbnails.resized, 0, "Thumbnails not touched");
		});

		WithHistory(function(oHistory)
		{
			let oDoc = CreateDoc(8, true);
			AscCommon.PDFEditorApi.prototype.asc_InsertPagesFromBytes.call(CreateApi(oDoc), 2, aBytes);
			assert.deepEqual(oHistory.Points[0].Items, ["Earlier", "StartMerge", "PartMerge", "EndMerge", "AddPage"], "A successful merge keeps its records");
			assert.strictEqual(oDoc.mergedPagesData.length, 2, "A successful merge keeps its data");
		});
	});

	QUnit.test("The first new page is shown", function (assert)
	{
		let oDoc = CreateDoc(8, true);
		InsertPages(oDoc, 3, aBytes);
		assert.strictEqual(oDoc.Viewer.thumbnails.resized, 1, "The thumbnails lay out the new pages at once");
		assert.deepEqual(oDoc.Viewer.navigated, [3], "The page view goes to the first new page");

		oDoc = CreateDoc(8, true);
		oDoc.Viewer.thumbnails.needResize = false;
		InsertPages(oDoc, 3, aBytes);
		assert.strictEqual(oDoc.Viewer.thumbnails.resized, 0, "No extra layout when the thumbnails are up to date");

		oDoc = CreateDoc(8, true);
		oDoc.Viewer.thumbnails = null;
		InsertPages(oDoc, 3, aBytes);
		assert.deepEqual(oDoc.Viewer.navigated, [3], "Works without thumbnails");
	});

	QUnit.test("The file dialog tells whether pages were inserted", function (assert)
	{
		let oOldDesktop = window["AscDesktopEditor"];
		let sData = null;
		window["AscDesktopEditor"] = {
			"OpenFilenameDialog" : function(sFilter, isMultiple, fOnFile)
			{
				fOnFile(sData === undefined ? "" : "C:/other.pdf");
			},
			"loadLocalFile"      : function(sFile, fOnData)
			{
				fOnData(sData);
			}
		};
		function InsertFromFile(oDoc)
		{
			let aResults = [];
			WithHistory(function()
			{
				AscCommon.PDFEditorApi.prototype.asc_InsertPagesFromFile.call(CreateApi(oDoc), 2, function(isDone, nResult)
				{
					aResults.push(isDone, nResult);
				});
			});
			return aResults;
		}
		try
		{
			sData = aBytes.buffer;
			assert.deepEqual(InsertFromFile(CreateDoc(8, true)), [true, Result.Ok], "Inserted");
			assert.deepEqual(InsertFromFile(CreateDoc(8, false)), [false, Result.Damaged], "Damaged file");
			assert.deepEqual(InsertFromFile(CreateDoc(8, true, true)), [false, Result.Locked], "Locked by another user");
			sData = null;
			assert.deepEqual(InsertFromFile(CreateDoc(8, true)), [false, Result.NoFile], "The file could not be read");
			sData = undefined;
			assert.deepEqual(InsertFromFile(CreateDoc(8, true)), [], "No file chosen: no callback");
		}
		finally
		{
			window["AscDesktopEditor"] = oOldDesktop;
		}
	});

	QUnit.test("The reason why nothing was inserted", function (assert)
	{
		assert.strictEqual(typeof AscCommon.PDFEditorApi.prototype["asc_GetInsertPagesMaxSize"], "function", "asc_GetInsertPagesMaxSize is exported");
		assert.strictEqual(Asc["c_oAscPdfInsertPagesResult"]["Locked"], Result.Locked, "Asc.c_oAscPdfInsertPagesResult is exported");

		let oDoc = CreateDoc(8, true, true);
		assert.strictEqual(InsertPages(oDoc, 2, aBytes), Result.Locked, "A merge lock is not reported as a damaged file");
		assert.strictEqual(oDoc.merges.length, 0, "The engine is not called while locked");

		let aEncrypted = ToBytes("%PDF-1.7\n1 0 obj\n<< /Type /Catalog >>\nendobj\ntrailer\n<< /Root 1 0 R /Encrypt 5 0 R >>\n%%EOF\n");
		assert.strictEqual(InsertPages(CreateDoc(8, false), 2, aEncrypted), Result.Password, "Encrypted file the engine can't read");
		assert.strictEqual(InsertPages(CreateDoc(8, true), 2, aEncrypted), Result.Ok, "An encrypted file the engine can read is inserted");

		assert.strictEqual(InsertPages(CreateDoc(8, false), 2, aBytes), Result.Damaged, "Not encrypted, not readable: damaged");
		assert.strictEqual(InsertPages(CreateDoc(8, false), 2, ToBytes("PK\u0003\u0004 /Encrypt")), Result.Damaged, "Not a PDF at all");

		// a linearized file has its first trailer at the start; padding puts the end out of the search range
		let sPadding = new Array(1100000).join(" ");
		let aLinearized = ToBytes("%PDF-1.5\ntrailer\n<< /Encrypt 9 0 R >>\n" + sPadding + "%%EOF\n");
		assert.strictEqual(InsertPages(CreateDoc(8, false), 2, aLinearized), Result.Password, "/Encrypt at the start");
		let aAtEnd = ToBytes("%PDF-1.5\n" + sPadding + "trailer\n<< /Encrypt 9 0 R >>\n%%EOF\n");
		assert.strictEqual(InsertPages(CreateDoc(8, false), 2, aAtEnd), Result.Password, "/Encrypt at the end");
	});

	QUnit.test("Files larger than the limit are refused", function (assert)
	{
		let nMax = 100 * 1024 * 1024;
		let oApi = CreateApi(CreateDoc(8, true));
		assert.strictEqual(oApi.asc_GetInsertPagesMaxSize(), nMax, "100 MB by default");
		oApi.maxChangesSize = 10 * 1024 * 1024;
		assert.strictEqual(oApi.asc_GetInsertPagesMaxSize(), 10 * 1024 * 1024, "The server's smaller limit for changes");
		oApi.maxChangesSize = 500 * 1024 * 1024;
		assert.strictEqual(oApi.asc_GetInsertPagesMaxSize(), nMax, "A larger server limit does not raise it");

		// only the length is read before the engine gets the file
		let oDoc = CreateDoc(8, true);
		let oHuge = {length : nMax + 1};
		assert.strictEqual(InsertPages(oDoc, 2, oHuge), Result.TooLarge, "One byte over the limit");
		assert.strictEqual(oDoc.merges.length, 0, "The engine is not called");
		assert.strictEqual(oDoc.actions.length, 0, "No history action");

		oDoc = CreateDoc(8, true);
		assert.strictEqual(InsertPages(oDoc, 2, {length : nMax}), Result.Ok, "Exactly the limit is fine");
	});

	QUnit.test("The browser file dialog does not read a file over the limit", function (assert)
	{
		let fOldCreate = document.createElement;
		let oOldDesktop = window["AscDesktopEditor"];
		let fOldReader = window.FileReader;
		let nReads = 0;
		let oInput = null;
		window["AscDesktopEditor"] = undefined;
		window.FileReader = function() { nReads++; this.readAsArrayBuffer = function() {}; };
		document.createElement = function(sTag)
		{
			let oElement = fOldCreate.call(document, sTag);
			if (sTag === "input")
			{
				oInput = oElement;
				oElement.click = function() {};
			}
			return oElement;
		};
		try
		{
			let aResults = [];
			AscCommon.PDFEditorApi.prototype.asc_InsertPagesFromFile.call(CreateApi(CreateDoc(8, true)), 2, function(isDone, nResult)
			{
				aResults.push(isDone, nResult);
			});
			Object.defineProperty(oInput, "files", {value : [{size : 100 * 1024 * 1024 + 1}]});
			oInput.onchange();
			assert.deepEqual(aResults, [false, Result.TooLarge], "Too large");
			assert.strictEqual(nReads, 0, "The file is not read");
		}
		finally
		{
			document.createElement = fOldCreate;
			window["AscDesktopEditor"] = oOldDesktop;
			window.FileReader = fOldReader;
		}
	});
});
