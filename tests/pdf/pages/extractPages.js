/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	// The new file is written by CPDFDoc.GetPagesBinary with the native PDF engine (the same call as "copy page"),
	// which does not run headless. These tests check the API around it: which pages can be written, the indexes
	// given to the engine, and that the opened document is not changed.
	function CreateAnnot(oProps)
	{
		oProps = oProps || {};
		return {
			IsRedact     : function() { return !!oProps.redact; },
			GetRedactId  : function() { return oProps.redactId; },
			IsShapeBased : function() { return !!oProps.shape; },
			IsChanged    : function() { return !!oProps.changed; }
		};
	}

	function CreateDoc(nPagesCount)
	{
		let oDoc = {
			binaryCalls        : [],
			actions            : 0,
			appliedRedactsData : [],
			pages              : [],
			Viewer             : {file : {type : 0, pages : []}},
			GetPagesCount      : function() { return this.pages.length; },
			GetPageInfo        : function(nPage) { return this.pages[nPage]; },
			DoAction           : function() { this.actions++; },
			RemovePages        : function() { throw new Error("must not change the document"); },
			GetPagesBinary     : function(aIndexes, bytesToBase64)
			{
				this.binaryCalls.push({indexes : aIndexes.slice(), base64 : bytesToBase64});
				return new Uint8Array([0x25, 0x50, 0x44, 0x46, aIndexes.length]);
			}
		};
		for (let i = 0; i < nPagesCount; i++)
		{
			oDoc.Viewer.file.pages.push({originIndex : i});
			oDoc.pages.push({
				id       : "page" + i,
				drawings : [],
				annots   : [],
				GetId    : function() { return this.id; }
			});
		}
		return oDoc;
	}

	function CreateApi(oDoc)
	{
		let oProto = AscCommon.PDFEditorApi.prototype;
		return {
			getPDFDoc            : function() { return oDoc; },
			asc_CanGetPagesAsPdf : oProto.asc_CanGetPagesAsPdf,
			asc_GetPagesAsPdf    : oProto.asc_GetPagesAsPdf,
			asc_SavePagesAsPdf   : oProto.asc_SavePagesAsPdf
		};
	}

	QUnit.module("PDF: extract the selected pages into a new PDF");

	QUnit.test("The API is exported", function (assert)
	{
		let oProto = AscCommon.PDFEditorApi.prototype;
		assert.strictEqual(typeof oProto["asc_CanGetPagesAsPdf"], "function", "asc_CanGetPagesAsPdf");
		assert.strictEqual(typeof oProto["asc_GetPagesAsPdf"], "function", "asc_GetPagesAsPdf");
		assert.strictEqual(typeof oProto["asc_SavePagesAsPdf"], "function", "asc_SavePagesAsPdf");
	});

	QUnit.test("The selected pages are written by the engine, the document stays unchanged", function (assert)
	{
		let oDoc = CreateDoc(5);
		let oApi = CreateApi(oDoc);
		let aBytes = oApi.asc_GetPagesAsPdf([3, 1, 3]);
		assert.ok(aBytes instanceof Uint8Array, "The bytes of a file are returned");
		assert.strictEqual(oDoc.binaryCalls.length, 1, "The engine is called once");
		assert.deepEqual(oDoc.binaryCalls[0].indexes, [1, 3], "Sorted page indexes without duplicates");
		assert.strictEqual(oDoc.binaryCalls[0].base64, false, "Bytes, not base64");
		assert.strictEqual(oDoc.GetPagesCount(), 5, "No page was removed");
		assert.strictEqual(oDoc.actions, 0, "No history action");
	});

	QUnit.test("Invalid page lists give nothing", function (assert)
	{
		let oDoc = CreateDoc(3);
		let oApi = CreateApi(oDoc);
		assert.strictEqual(oApi.asc_GetPagesAsPdf([]), null, "No pages");
		assert.strictEqual(oApi.asc_GetPagesAsPdf(undefined), null, "No list");
		assert.strictEqual(oApi.asc_GetPagesAsPdf([3]), null, "Page after the last one");
		assert.strictEqual(oApi.asc_GetPagesAsPdf([-1]), null, "Negative index");
		assert.strictEqual(oApi.asc_GetPagesAsPdf([0.5]), null, "Not an integer");
		assert.strictEqual(oApi.asc_GetPagesAsPdf(["1"]), null, "Not a number");
		assert.strictEqual(oDoc.binaryCalls.length, 0, "The engine is not called");
	});

	QUnit.test("Only PDF files", function (assert)
	{
		let oDoc = CreateDoc(3);
		oDoc.Viewer.file.type = 1; // XPS
		assert.strictEqual(CreateApi(oDoc).asc_CanGetPagesAsPdf([0]), false, "An XPS file can't be split");
		assert.strictEqual(CreateApi(oDoc).asc_GetPagesAsPdf([0]), null, "Nothing is returned");
		assert.strictEqual(oDoc.binaryCalls.length, 0, "The engine is not called");
	});

	QUnit.test("Pages with content the engine can't write are refused", function (assert)
	{
		function Check(sCase, fChange, isOtherPage)
		{
			let oDoc = CreateDoc(4);
			fChange(oDoc, isOtherPage ? 0 : 2);
			let oApi = CreateApi(oDoc);
			let sName = sCase + (isOtherPage ? " (on a page that is not extracted)" : "");
			assert.strictEqual(oApi.asc_CanGetPagesAsPdf([1, 2]), !!isOtherPage, sName);
			assert.strictEqual(oApi.asc_GetPagesAsPdf([1, 2]) !== null, !!isOtherPage, sName + ": bytes only when allowed");
			assert.strictEqual(oDoc.binaryCalls.length, isOtherPage ? 1 : 0, sName + ": engine called only when allowed");
		}

		let aCases = [
			["Drawing added in the editor", function(oDoc, nPage) { oDoc.pages[nPage].drawings.push({}); }],
			["Edited page text", function(oDoc, nPage) { oDoc.Viewer.file.pages[nPage].isRecognized = true; }],
			["New blank page", function(oDoc, nPage) { oDoc.Viewer.file.pages[nPage].originIndex = undefined; }],
			["Changed drawn annotation", function(oDoc, nPage) { oDoc.pages[nPage].annots.push(CreateAnnot({shape : true, changed : true})); }],
			["Applied redaction annotation", function(oDoc, nPage) { oDoc.pages[nPage].annots.push(CreateAnnot({redact : true, redactId : "r1"})); }],
			["Applied redaction data", function(oDoc, nPage) { oDoc.appliedRedactsData.push({pageId : "page" + nPage}); }]
		];
		aCases.forEach(function(aCase)
		{
			Check(aCase[0], aCase[1], false);
			Check(aCase[0], aCase[1], true);
		});
	});

	QUnit.test("Unchanged annotations and comments are fine", function (assert)
	{
		let oDoc = CreateDoc(2);
		oDoc.pages[0].annots.push(CreateAnnot({shape : true}));
		oDoc.pages[0].annots.push(CreateAnnot({changed : true}));
		oDoc.pages[0].annots.push(CreateAnnot({redact : true}));
		assert.strictEqual(CreateApi(oDoc).asc_CanGetPagesAsPdf([0]), true, "A drawn annotation as opened, a changed comment and a redaction not applied yet");
	});

	QUnit.test("A changed page order is refused", function (assert)
	{
		let oDoc = CreateDoc(4);
		oDoc.Viewer.file.pages[0].originIndex = 2;
		oDoc.Viewer.file.pages[2].originIndex = 0;
		let oApi = CreateApi(oDoc);
		assert.strictEqual(oApi.asc_CanGetPagesAsPdf([0, 2]), false, "Pages moved against each other");
		assert.strictEqual(oApi.asc_CanGetPagesAsPdf([1, 3]), true, "Pages still in the order of the file");
		assert.strictEqual(oApi.asc_CanGetPagesAsPdf([0]), true, "One moved page");
	});

	QUnit.test("An engine error gives nothing", function (assert)
	{
		let oDoc = CreateDoc(2);
		oDoc.GetPagesBinary = function() { throw new Error("engine"); };
		assert.strictEqual(CreateApi(oDoc).asc_GetPagesAsPdf([0]), null, "Exception");
		oDoc.GetPagesBinary = function() { return new Uint8Array(0); };
		assert.strictEqual(CreateApi(oDoc).asc_GetPagesAsPdf([0]), null, "Empty result");
	});

	QUnit.test("Saving downloads the file", function (assert)
	{
		let oOldDownload = AscCommon.DownloadFileFromBytes;
		let aDownloads = [];
		AscCommon.DownloadFileFromBytes = function(aData, sName, sMime)
		{
			aDownloads.push({data : aData, name : sName, mime : sMime});
		};
		try
		{
			let oDoc = CreateDoc(3);
			let oApi = CreateApi(oDoc);
			assert.strictEqual(oApi.asc_SavePagesAsPdf([0, 2], "Report (pages 1, 3).pdf"), true, "Saved");
			assert.strictEqual(aDownloads.length, 1, "One download");
			assert.strictEqual(aDownloads[0].name, "Report (pages 1, 3).pdf", "File name");
			assert.strictEqual(aDownloads[0].mime, "application/pdf", "PDF type");
			assert.ok(aDownloads[0].data instanceof Uint8Array, "The bytes from the engine");

			oDoc.pages[0].drawings.push({});
			assert.strictEqual(oApi.asc_SavePagesAsPdf([0], "x.pdf"), false, "Refused pages are not saved");
			assert.strictEqual(aDownloads.length, 1, "No second download");
		}
		finally
		{
			AscCommon.DownloadFileFromBytes = oOldDownload;
		}
	});
});
