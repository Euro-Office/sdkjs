/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

// A document saved with the markup of its tracked insertions and deletions hidden
// (w:revisionView w:insDel="0", ODF ShowRedlineChanges false): the setting survives the binary
// format, and converting the document (PDF, image, HTML) renders it as Final.

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();
	logicDocument.RemoveFromContent(0, logicDocument.GetElementsCount(), false);

	QUnit.module("Test the saved revision view (w:revisionView w:insDel)");

	// "Hello world" with "Hello " replaced by "Hi " while tracking: [Remove "Hello "][Add "Hi "]["world"]
	function fillDocumentWithChanges()
	{
		AscTest.SetTrackRevisions(false);
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.PushToContent(p);
		AscTest.MoveCursorToParagraph(p);
		AscTest.EnterText("Hello world");
		AscTest.SetTrackRevisions(true);
		AscTest.SelectParagraphRange(p, 0, 6);
		AscTest.EnterText("Hi ");
		AscTest.SetTrackRevisions(false);
		logicDocument.RemoveSelection();
		return p;
	}

	function writeAndRead(revisionViewInsDel)
	{
		// The writer reads the macros of the editor; the test editor has none.
		let editorApi = logicDocument.DrawingDocument.m_oWordControl.m_oApi;
		if (!editorApi.macros)
			editorApi.macros = new AscCommon.CDocumentMacros();

		try
		{
			logicDocument.Settings.RevisionViewInsDel = revisionViewInsDel;
			let data = new AscCommonWord.BinaryFileWriter(logicDocument).Write();
			logicDocument.Settings.RevisionViewInsDel = "not read";
			new AscCommonWord.BinaryFileReader(logicDocument, {}).Read(data);
			return logicDocument.Settings.RevisionViewInsDel;
		}
		finally
		{
			logicDocument.Settings.RevisionViewInsDel = undefined;
		}
	}

	// The conversion path (x2t -> doctrenderer) calls asc_nativeCalculateFile before it draws the pages.
	// The test editor replaces window["asc_docs_api"]; the word editor class is still Asc.asc_docs_api.
	function calculateForConversion()
	{
		let api = {WordControl : {m_oLogicDocument : logicDocument}};
		Asc.asc_docs_api.prototype["asc_nativeCalculateFile"].call(api);
	}

	const changesKept = [
		[reviewtype_Remove, "Hello "],
		[reviewtype_Add, "Hi "],
		[reviewtype_Common, "world"]
	];

	QUnit.test("RevisionViewInsDel false survives writing and reading the binary format", function (assert)
	{
		AscTest.ClearDocument();
		logicDocument.PushToContent(AscTest.CreateParagraph());
		assert.strictEqual(writeAndRead(false), false, "insDel false is written and read back");
	});

	QUnit.test("No RevisionViewInsDel record when the setting is not false", function (assert)
	{
		AscTest.ClearDocument();
		logicDocument.PushToContent(AscTest.CreateParagraph());
		assert.strictEqual(writeAndRead(undefined), "not read", "nothing is written for an unset value");
		assert.strictEqual(writeAndRead(true), "not read", "nothing is written for true (the default view)");
	});

	QUnit.test("Converting a document saved without change markup accepts its changes", function (assert)
	{
		let p = fillDocumentWithChanges();
		assert.deepEqual(AscTest.GetParagraphReviewText(p), changesKept, "tracked changes before the conversion");

		logicDocument.Settings.RevisionViewInsDel = false;
		try
		{
			calculateForConversion();
		}
		finally
		{
			logicDocument.Settings.RevisionViewInsDel = undefined;
		}

		assert.deepEqual(AscTest.GetParagraphReviewText(p), [[reviewtype_Common, "Hi world"]], "rendered as Final");
	});

	QUnit.test("Converting keeps the changes when the markup was not hidden", function (assert)
	{
		let p = fillDocumentWithChanges();
		calculateForConversion();
		assert.deepEqual(AscTest.GetParagraphReviewText(p), changesKept, "unset: changes kept");

		p = fillDocumentWithChanges();
		logicDocument.Settings.RevisionViewInsDel = true;
		try
		{
			calculateForConversion();
		}
		finally
		{
			logicDocument.Settings.RevisionViewInsDel = undefined;
		}
		assert.deepEqual(AscTest.GetParagraphReviewText(p), changesKept, "true: changes kept");
	});

	QUnit.test("The mobile editor keeps the changes", function (assert)
	{
		let p = fillDocumentWithChanges();
		logicDocument.Settings.RevisionViewInsDel = false;
		window["IS_NATIVE_EDITOR"] = true;
		try
		{
			calculateForConversion();
		}
		finally
		{
			delete window["IS_NATIVE_EDITOR"];
			logicDocument.Settings.RevisionViewInsDel = undefined;
		}
		assert.deepEqual(AscTest.GetParagraphReviewText(p), changesKept, "IS_NATIVE_EDITOR: changes kept");
	});

	QUnit.test("Docbuilder keeps the changes", function (assert)
	{
		// docbuilder saves a PDF from the document it keeps open; a later save must still have the changes.
		let p = fillDocumentWithChanges();
		logicDocument.Settings.RevisionViewInsDel = false;
		window["builderJS"] = {};
		try
		{
			calculateForConversion();
		}
		finally
		{
			delete window["builderJS"];
			logicDocument.Settings.RevisionViewInsDel = undefined;
		}
		assert.deepEqual(AscTest.GetParagraphReviewText(p), changesKept, "builderJS: changes kept");
	});
});
