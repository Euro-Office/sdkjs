/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	const logicDocument = AscTest.CreateLogicDocument();
	
	function countCalls(object, method, f)
	{
		let original = object[method];
		let count    = 0;
		object[method] = function()
		{
			++count;
			return original.apply(this, arguments);
		};
		try
		{
			f();
		}
		finally
		{
			object[method] = original;
		}
		return count;
	}
	
	QUnit.module("Section updates on content changes");
	
	QUnit.test("Adding to or removing from table cell content doesn't scan the table rows", function (assert)
	{
		AscTest.ClearDocument();
		let table = AscTest.CreateTable(50, 2);
		logicDocument.PushToContent(table);
		
		let cellContent = table.GetRow(49).GetCell(1).GetContent();
		let scans = countCalls(AscWord.CTable.prototype, "IsUseInDocument", function()
		{
			cellContent.AddToContent(1, AscTest.CreateParagraph());
			cellContent.RemoveFromContent(1, 1);
		});
		
		assert.strictEqual(cellContent.GetElementsCount(), 1, "Paragraph added to and removed from the cell");
		assert.strictEqual(scans, 0, "CTable.IsUseInDocument is not called for cell content");
	});
	
	QUnit.test("Adding to or removing from the document body still updates sections", function (assert)
	{
		AscTest.ClearDocument();
		let sections = logicDocument.GetSections();
		let updates  = countCalls(sections, "UpdateOnAdd", function()
		{
			logicDocument.AddToContent(0, AscTest.CreateParagraph());
		});
		updates += countCalls(sections, "UpdateOnRemove", function()
		{
			logicDocument.RemoveFromContent(0, 1);
		});
		
		assert.strictEqual(updates, 2, "Sections updated once on add and once on remove");
	});
	
	QUnit.test("Content parented directly to the document (not in its body) can be filled", function (assert)
	{
		AscTest.ClearDocument();
		let content = new AscWord.CDocumentContent(logicDocument, logicDocument.GetDrawingDocument(), 0, 0, 0, 0, true, false, false);
		content.AddToContent(0, AscTest.CreateParagraph());
		content.RemoveFromContent(0, 1);
		assert.strictEqual(content.GetElementsCount(), 1, "Paragraph added to and removed from detached content");
	});

	QUnit.test("Content of a block-level content control without a parent can be filled", function (assert)
	{
		// A content control created from a remote change has no parent until it is added to the document
		AscTest.ClearDocument();
		let sdt = new AscWord.CBlockLevelSdt();
		sdt.Content.LogicDocument = logicDocument;
		sdt.Content.AddToContent(0, AscTest.CreateParagraph());
		sdt.Content.RemoveFromContent(0, 1);
		assert.strictEqual(sdt.Content.GetElementsCount(), 1, "Paragraph added to and removed from the content control");
	});
});
