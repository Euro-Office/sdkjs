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
		return p;
	}

	function GetFirstParagraph(shape)
	{
		return shape.getDocContent().GetElement(0);
	}

	QUnit.module("Paragraphs in grouped shapes");

	QUnit.test("CDocument.GetAllParagraphs collects paragraphs from grouped shapes", function(assert)
	{
		AscTest.ClearDocument();

		// A group with two text boxes and a nested group, which also has two text boxes
		let outerShape  = CreateTextBox();
		let outerShape2 = CreateTextBox();
		let innerShape  = CreateTextBox();
		let innerShape2 = CreateTextBox();

		let innerGroup = CreateGroup([innerShape, innerShape2]);
		let outerGroup = CreateGroup([outerShape, outerShape2, innerGroup]);
		let drawingParagraph = AddDrawingToDocument(outerGroup);

		let shapeParagraphs = [
			GetFirstParagraph(outerShape),
			GetFirstParagraph(outerShape2),
			GetFirstParagraph(innerShape),
			GetFirstParagraph(innerShape2)
		];

		let paragraphs = logicDocument.GetAllParagraphs({All : true});
		assert.ok(-1 !== paragraphs.indexOf(drawingParagraph), "The paragraph holding the group is returned");
		assert.ok(-1 !== paragraphs.indexOf(shapeParagraphs[0]), "The paragraph of a grouped text box is returned");
		assert.ok(-1 !== paragraphs.indexOf(shapeParagraphs[2]), "The paragraph of a text box in a nested group is returned");
		assert.strictEqual(new Set(paragraphs).size, paragraphs.length, "No paragraph is returned twice");

		let groupParagraphs = outerGroup.GetAllParagraphs({All : true});
		assert.deepEqual(groupParagraphs, shapeParagraphs, "The group returns the paragraphs of its shapes and of its nested group, in order");

		let passedArray = [];
		assert.strictEqual(outerGroup.GetAllParagraphs({All : true}, passedArray), passedArray, "The group fills and returns the array it is given");
		assert.strictEqual(passedArray.length, 4, "The given array gets every paragraph of the group");
	});
});
