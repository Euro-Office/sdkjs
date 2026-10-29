/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

$(function () {

	let logicDocument = AscTest.CreateLogicDocument();
	let styleManager  = logicDocument.GetStyles();

	function InitDocument()
	{
		AscTest.ClearDocument();
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		AscTest.AddTextToParagraph(p, "ABC");
	}

	function AddTable(content, pos, styleName)
	{
		let table = AscTest.CreateTable(2, 2);
		content.AddToContent(pos, table);
		table.Set_TableStyle(styleManager.GetStyleIdByName(styleName));
		return table;
	}

	// Returns the names of the table styles that the style table writer puts into the file
	function GetWrittenTableStyles(copyParams)
	{
		let writer = new BinaryStyleTableWriter(new AscCommon.CMemory(), logicDocument, {}, copyParams ? copyParams : null, new AscCommonWord.DocSaveParams());
		let names  = [];

		let writeStyle = writer.WriteStyle;
		writer.WriteStyle = function(id, style, isDefault)
		{
			if (style.IsTableStyle())
				names.push(style.GetName());

			return writeStyle.call(this, id, style, isDefault);
		};
		writer.Write();
		return names.sort();
	}

	function AddTextBoxWithTable(styleName)
	{
		let shapeTrack = new AscFormat.NewShapeTrack("rect", 0, 0, logicDocument.theme, null, null, null, 0);
		shapeTrack.track({}, 50, 30);
		let shape = shapeTrack.getShape(true, logicDocument.GetDrawingDocument(), null);
		shape.setBDeleted(false);
		shape.createTextBoxContent();
		AddTable(shape.getDocContent(), 0, styleName);
		return shape;
	}

	function CreateGroup(drawings)
	{
		let group = new AscFormat.CGroupShape();
		group.setSpPr(new AscFormat.CSpPr());
		group.spPr.setParent(group);
		group.spPr.setXfrm(new AscFormat.CXfrm());
		group.spPr.xfrm.setParent(group.spPr);
		for (let i = 0; i < drawings.length; ++i)
		{
			drawings[i].setGroup(group);
			group.addToSpTree(group.spTree.length, drawings[i]);
		}
		group.setBDeleted(false);
		return group;
	}

	QUnit.module("Table styles written on save");

	QUnit.test("Document without tables", function(assert)
	{
		InitDocument();
		assert.deepEqual(GetWrittenTableStyles(), ["Normal Table", "Table Grid"], "Only the default table styles are written");
	});

	QUnit.test("Document with a styled table", function(assert)
	{
		InitDocument();
		AddTable(logicDocument, 1, "Grid Table 4 - Accent 1");
		assert.deepEqual(GetWrittenTableStyles(), ["Grid Table 4 - Accent 1", "Normal Table", "Table Grid"], "The used style and its BasedOn chain are written");

		InitDocument();
		let headerContent = AscTest.CreateDefaultHeader(AscTest.GetFinalSection());
		AddTable(headerContent, 0, "Plain Table 2");
		assert.deepEqual(GetWrittenTableStyles(), ["Normal Table", "Plain Table 2", "Table Grid"], "A style used in a header is written");

		InitDocument();
		let outerTable = AddTable(logicDocument, 1, "Grid Table 1 Light");
		AddTable(outerTable.GetRow(0).GetCell(0).GetContent(), 0, "List Table 3");
		assert.deepEqual(GetWrittenTableStyles(), ["Grid Table 1 Light", "List Table 3", "Normal Table", "Table Grid"], "A style used in a nested table is written");
	});

	QUnit.test("Modified built-in table style", function(assert)
	{
		InitDocument();
		let style     = styleManager.Get(styleManager.GetStyleIdByName("Plain Table 1"));
		let oldTextPr = style.TextPr.Copy();
		let newTextPr = style.TextPr.Copy();
		newTextPr.Bold = true;
		style.Set_TextPr(newTextPr);

		assert.deepEqual(GetWrittenTableStyles(), ["Normal Table", "Plain Table 1", "Table Grid"], "A modified built-in style is written even if it is not used");

		style.Set_TextPr(oldTextPr);
		assert.deepEqual(GetWrittenTableStyles(), ["Normal Table", "Table Grid"], "After restoring it, the style is skipped again");
	});

	QUnit.test("Table style gallery after save", function(assert)
	{
		InitDocument();
		let count = logicDocument.GetAllTableStyles().length;
		GetWrittenTableStyles();
		assert.ok(count > 100, "The document has all the built-in table styles");
		assert.strictEqual(logicDocument.GetAllTableStyles().length, count, "Saving does not remove table styles from the document");
	});

	QUnit.test("Custom style based on a built-in table style", function(assert)
	{
		InitDocument();
		let style = AscTest.CreateStyle(styletype_Table, "Custom Table");
		style.SetBasedOn(styleManager.GetStyleIdByName("Plain Table 3"));

		assert.deepEqual(GetWrittenTableStyles(), ["Custom Table", "Normal Table", "Plain Table 3", "Table Grid"], "The BasedOn chain of a written style is written");

		styleManager.Remove(style.GetId());
	});

	QUnit.test("Copy/paste writes the used styles", function(assert)
	{
		InitDocument();
		let usedStyleMap = {};
		usedStyleMap[styleManager.GetStyleIdByName("Plain Table 2")] = 1;
		usedStyleMap[styleManager.GetStyleIdByName("Grid Table 4 - Accent 1")] = 1;
		assert.deepEqual(GetWrittenTableStyles({oUsedStyleMap : usedStyleMap}), ["Grid Table 4 - Accent 1", "Plain Table 2"], "Copy/paste is not filtered");
	});

	QUnit.test("Table in a grouped shape", function(assert)
	{
		InitDocument();
		let innerGroup = CreateGroup([AddTextBoxWithTable("Grid Table 5 Dark - Accent 3")]);
		let outerGroup = CreateGroup([AddTextBoxWithTable("Grid Table 4 - Accent 2"), innerGroup]);

		let drawing = new ParaDrawing(100, 100, null, logicDocument.GetDrawingDocument(), logicDocument, null);
		outerGroup.setParent(drawing);
		drawing.Set_GraphicObject(outerGroup);
		let run = new AscWord.CRun();
		logicDocument.GetElement(0).AddToContent(0, run);
		run.AddToContent(0, drawing);

		assert.deepEqual(GetWrittenTableStyles(), ["Grid Table 4 - Accent 2", "Grid Table 5 Dark - Accent 3", "Normal Table", "Table Grid"], "Styles used in grouped and nested grouped shapes are written");
	});
});
