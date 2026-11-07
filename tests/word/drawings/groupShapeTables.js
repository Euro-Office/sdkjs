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

	function CreateTextBox()
	{
		let shapeTrack = new AscFormat.NewShapeTrack("rect", 0, 0, logicDocument.theme, null, null, null, 0);
		shapeTrack.track({}, 50, 30);
		let shape = shapeTrack.getShape(true, logicDocument.GetDrawingDocument(), null);
		shape.setBDeleted(false);
		shape.createTextBoxContent();
		return shape;
	}

	function AddTableToTextBox(shape)
	{
		let table = AscTest.CreateTable(2, 2);
		shape.getDocContent().AddToContent(0, table);
		return table;
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

	function AddDrawingToDocument(graphicObject)
	{
		let drawing = new ParaDrawing(100, 100, null, logicDocument.GetDrawingDocument(), logicDocument, null);
		graphicObject.setParent(drawing);
		drawing.Set_GraphicObject(graphicObject);

		let p   = AscTest.CreateParagraph();
		let run = new AscWord.CRun();
		p.AddToContent(0, run);
		run.AddToContent(0, drawing);
		logicDocument.AddToContent(logicDocument.GetElementsCount(), p);
		return drawing;
	}

	QUnit.module("Tables in grouped shapes");

	QUnit.test("CDocument.GetAllTables collects tables from grouped shapes", function(assert)
	{
		AscTest.ClearDocument();

		let bodyTable = AscTest.CreateTable(2, 2);
		logicDocument.AddToContent(0, bodyTable);

		// A group with two text boxes and a nested group, which also has two text boxes
		let outerShape = CreateTextBox();
		let innerShape = CreateTextBox();
		let outerTable = AddTableToTextBox(outerShape);
		let innerTable = AddTableToTextBox(innerShape);

		let innerGroup = CreateGroup([innerShape, CreateTextBox()]);
		let outerGroup = CreateGroup([outerShape, CreateTextBox(), innerGroup]);
		AddDrawingToDocument(outerGroup);

		let tables = logicDocument.GetAllTables();
		assert.strictEqual(tables.length, 3, "The document returns the body table and both tables of the group");
		assert.ok(-1 !== tables.indexOf(bodyTable), "The body table is returned");
		assert.ok(-1 !== tables.indexOf(outerTable), "The table of a grouped text box is returned");
		assert.ok(-1 !== tables.indexOf(innerTable), "The table of a text box in a nested group is returned");
		assert.strictEqual(new Set(tables).size, tables.length, "No table is returned twice");

		let groupTables = outerGroup.GetAllTables();
		assert.strictEqual(groupTables.length, 2, "The group returns the tables of its shapes and of its nested group");
		assert.ok(outerTable === groupTables[0] && innerTable === groupTables[1], "The group keeps the order of its shapes");
	});
});
