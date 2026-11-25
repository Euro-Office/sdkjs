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
			permissions        : {},
			checkedPerms       : [],
			Viewer             : {file : {type : 0, pages : []}, savedPassword : ""},
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
		// CheckPerm of the engine: true unless the bit is switched off in oDoc.permissions
		oDoc.Viewer.file.nativeFile = {
			"CheckPerm" : function(nPerm)
			{
				oDoc.checkedPerms.push(nPerm);
				return oDoc.permissions[nPerm] !== false;
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
			currentPassword        : "",
			getPDFDoc              : function() { return oDoc; },
			asc_GetPagesAsPdfBlock : oProto.asc_GetPagesAsPdfBlock,
			asc_CanGetPagesAsPdf   : oProto.asc_CanGetPagesAsPdf,
			asc_GetPagesAsPdf      : oProto.asc_GetPagesAsPdf,
			asc_SavePagesAsPdf     : oProto.asc_SavePagesAsPdf
		};
	}

	QUnit.module("PDF: extract the selected pages into a new PDF");

	QUnit.test("The API is exported", function (assert)
	{
		let oProto = AscCommon.PDFEditorApi.prototype;
		assert.strictEqual(typeof oProto["asc_CanGetPagesAsPdf"], "function", "asc_CanGetPagesAsPdf");
		assert.strictEqual(typeof oProto["asc_GetPagesAsPdf"], "function", "asc_GetPagesAsPdf");
		assert.strictEqual(typeof oProto["asc_SavePagesAsPdf"], "function", "asc_SavePagesAsPdf");
		assert.strictEqual(typeof oProto["asc_GetPagesAsPdfBlock"], "function", "asc_GetPagesAsPdfBlock");
		assert.strictEqual(Asc["c_oAscPdfPagesAsPdfBlock"]["Permissions"], 3, "Asc.c_oAscPdfPagesAsPdfBlock");
	});

	QUnit.test("GetPagesBinary gives the bytes without making base64", function (assert)
	{
		let aEngineBytes = new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2D]);
		let aSplitCalls  = [];
		let oFakeDoc     = {
			Viewer : {
				SaveForSplit : function() { return null; },
				file         : {
					originalPagesCount : 3,
					pages              : [{originIndex : 0}, {originIndex : 1}, {originIndex : 2}],
					nativeFile         : {
						"SplitPages" : function(aIndexes) { aSplitCalls.push(aIndexes.slice()); return aEngineBytes; }
					}
				}
			}
		};

		let fOldEncode = AscCommon.Base64.encode;
		let nEncodeCalls = 0;
		AscCommon.Base64.encode = function()
		{
			nEncodeCalls++;
			return fOldEncode.apply(this, arguments);
		};
		try
		{
			let aBytes = AscPDF.CPDFDoc.prototype.GetPagesBinary.call(oFakeDoc, [2, 0], false);
			assert.strictEqual(aBytes, aEngineBytes, "The bytes of the engine are returned");
			assert.deepEqual(aSplitCalls[0], [0, 2], "The engine gets the sorted origin indexes");
			assert.strictEqual(nEncodeCalls, 0, "No base64 string is made for bytes");

			let sBase64 = AscPDF.CPDFDoc.prototype.GetPagesBinary.call(oFakeDoc, [1]);
			assert.strictEqual(typeof sBase64, "string", "base64 is still the default");
			assert.strictEqual(nEncodeCalls, 1, "base64 is made when it is asked for");
		}
		finally
		{
			AscCommon.Base64.encode = fOldEncode;
		}
	});

	QUnit.test("The permissions of a protected file are respected", function (assert)
	{
		let oBlock = Asc.c_oAscPdfPagesAsPdfBlock;

		let oDoc = CreateDoc(3);
		let oApi = CreateApi(oDoc);
		assert.strictEqual(oApi.asc_GetPagesAsPdfBlock([0]), oBlock.None, "No restriction");
		assert.deepEqual(oDoc.checkedPerms, [5], "The copy/extract permission (bit 5) is checked");

		oDoc.permissions[5] = false;
		assert.strictEqual(oApi.asc_GetPagesAsPdfBlock([0]), oBlock.Permissions, "Copying and extracting not permitted");
		assert.strictEqual(oApi.asc_CanGetPagesAsPdf([0]), false, "Can't extract");
		assert.strictEqual(oApi.asc_GetPagesAsPdf([0]), null, "No bytes");
		assert.strictEqual(oDoc.binaryCalls.length, 0, "The engine is not called");

		oApi.isPdfOwnerPasswordChecked = true;
		assert.strictEqual(oApi.asc_GetPagesAsPdfBlock([0]), oBlock.None, "Allowed after the owner password was entered");
		assert.ok(oApi.asc_GetPagesAsPdf([0]) instanceof Uint8Array, "Bytes for the owner");

		let oDoc2 = CreateDoc(3);
		oDoc2.Viewer.file.nativeFile = undefined;
		assert.strictEqual(CreateApi(oDoc2).asc_GetPagesAsPdfBlock([0]), oBlock.Permissions, "Refused when the permissions can't be checked");
	});

	QUnit.test("A file with an open password is refused", function (assert)
	{
		let oBlock = Asc.c_oAscPdfPagesAsPdfBlock;

		let oDoc = CreateDoc(2);
		let oApi = CreateApi(oDoc);
		oApi.currentPassword = "secret";
		assert.strictEqual(oApi.asc_GetPagesAsPdfBlock([0]), oBlock.Password, "Opened with a password (online)");
		assert.strictEqual(oApi.asc_GetPagesAsPdf([0]), null, "No bytes");

		oApi.currentPassword = "";
		oDoc.Viewer.savedPassword = "secret";
		assert.strictEqual(oApi.asc_GetPagesAsPdfBlock([0]), oBlock.Password, "Opened with a password (desktop)");

		oApi.isPdfOwnerPasswordChecked = true;
		assert.strictEqual(oApi.asc_GetPagesAsPdfBlock([0]), oBlock.None, "Allowed after the owner password was entered");
		assert.strictEqual(oDoc.binaryCalls.length, 0, "The engine was not called before");
	});

	QUnit.test("The block reasons", function (assert)
	{
		let oBlock = Asc.c_oAscPdfPagesAsPdfBlock;
		let oDoc = CreateDoc(2);
		let oApi = CreateApi(oDoc);
		assert.strictEqual(oApi.asc_GetPagesAsPdfBlock([5]), oBlock.Invalid, "Not a page");
		oDoc.pages[0].drawings.push({});
		assert.strictEqual(oApi.asc_GetPagesAsPdfBlock([0]), oBlock.Changes, "Changes the engine can't write");
		oDoc.permissions[5] = false;
		assert.strictEqual(oApi.asc_GetPagesAsPdfBlock([0]), oBlock.Permissions, "Permissions are told before changes");
	});

	QUnit.test("Entering the owner password is remembered", function (assert)
	{
		let oProto = AscCommon.PDFEditorApi.prototype;
		let sOwner = "owner";
		let oDoc = {
			Viewer : {
				file : {
					nativeFile : {
						"CheckPerm"          : function() { return false; },
						"CheckOwnerPassword" : function(sPassword) { return sPassword === sOwner; }
					}
				}
			}
		};
		let oApi = {getPDFDoc : function() { return oDoc; }};
		assert.strictEqual(oProto.asc_CheckEditPassword.call(oApi, null), false, "Only the permission is checked");
		assert.notOk(oApi.isPdfOwnerPasswordChecked, "Not without a password");
		assert.strictEqual(oProto.asc_CheckEditPassword.call(oApi, "wrong"), false, "Wrong password");
		assert.notOk(oApi.isPdfOwnerPasswordChecked, "Not with a wrong password");
		assert.strictEqual(oProto.asc_CheckPrintPassword.call(oApi, sOwner), true, "Owner password");
		assert.strictEqual(oApi.isPdfOwnerPasswordChecked, true, "Remembered");
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
