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
			}
		};
	}

	function InsertPages(oDoc, nPos, aBytes)
	{
		return AscCommon.PDFEditorApi.prototype.asc_InsertPagesFromBytes.call(CreateApi(oDoc), nPos, aBytes);
	}

	let aBytes = new Uint8Array([0x25, 0x50, 0x44, 0x46]);

	QUnit.module("PDF: insert the pages of another PDF");

	QUnit.test("All pages are inserted in one undo step", function (assert)
	{
		let oDoc = CreateDoc(8, true);
		assert.strictEqual(InsertPages(oDoc, 2, aBytes), true, "Insert reports success");
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
		assert.strictEqual(InsertPages(oDoc, 2, new Uint8Array(0)), false, "Empty file");
		assert.strictEqual(InsertPages(oDoc, 2, null), false, "No file");
		assert.strictEqual(oDoc.actions.length, 0, "No history action without a file");

		oDoc = CreateDoc(8, true, true);
		assert.strictEqual(InsertPages(oDoc, 2, aBytes), false, "Pages locked by another user");
		assert.strictEqual(oDoc.actions.length, 0, "No history action while locked");

		oDoc = CreateDoc(8, false);
		assert.strictEqual(InsertPages(oDoc, 2, aBytes), false, "The engine could not read the file");

		assert.strictEqual(AscCommon.PDFEditorApi.prototype.asc_InsertPagesFromBytes.call(CreateApi(null), 2, aBytes), false, "No document");
	});
});
