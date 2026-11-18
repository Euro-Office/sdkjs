/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function () {

	let logicDocument = AscTest.CreateLogicDocument();

	function CreateTextBox(name)
	{
		let shapeTrack = new AscFormat.NewShapeTrack("rect", 0, 0, logicDocument.theme, null, null, null, 0);
		shapeTrack.track({}, 50, 30);
		let shape = shapeTrack.getShape(true, logicDocument.GetDrawingDocument(), null);
		shape.setBDeleted(false);
		shape.createTextBoxContent();
		// like the test header, the test text box does not find the logic document by itself
		shape.textBoxContent.LogicDocument  = logicDocument;
		shape.textBoxContent.DrawingObjects = logicDocument.DrawingObjects;
		shape.textBoxContent.GetElement(0).LogicDocument = logicDocument;

		let drawing = new ParaDrawing(50, 30, null, logicDocument.GetDrawingDocument(), logicDocument, null);
		shape.setParent(drawing);
		drawing.Set_GraphicObject(shape);
		drawing.docPr.setName(name);
		return drawing;
	}

	function CreateImage(name)
	{
		let drawing = AscTest.CreateImage(50, 50);
		drawing.docPr.setName(name);
		return drawing;
	}

	function MakeFloating(drawing, behind, relativeHeight)
	{
		drawing.Set_DrawingType(drawing_Anchor);
		drawing.Set_WrappingType(WRAPPING_TYPE_NONE);
		drawing.Set_BehindDoc(behind);
		drawing.Set_PositionH(Asc.c_oAscRelativeFromH.Page, false, 20);
		drawing.Set_PositionV(Asc.c_oAscRelativeFromV.Page, false, 20);
		drawing.Set_RelativeHeight(relativeHeight);
	}

	function AddToParagraph(paragraph, drawing)
	{
		let run = new AscWord.CRun();
		paragraph.AddToContentToEnd(run);
		run.AddToContent(0, drawing);
	}

	// the api methods only need the logic document and the current page
	let api = {
		private_GetLogicDocument : function() { return logicDocument; },
		WordControl              : {m_oDrawingDocument : {m_lCurrentPage : 0}}
	};

	function GetPageDrawingObjects(pageIndex)
	{
		return Asc.asc_docs_api.prototype.asc_GetPageDrawingObjects.call(api, pageIndex);
	}

	function SelectDrawingObjectById(id, pageIndex)
	{
		return Asc.asc_docs_api.prototype.asc_SelectDrawingObjectById.call(api, id, pageIndex);
	}

	let scrolledTo = null;
	logicDocument.GetDrawingDocument().m_oWordControl.ScrollToPosition = function(x, y, pageIndex) { scrolledTo = pageIndex; };

	QUnit.module("Selection pane");

	QUnit.test("Objects in the header are listed after the objects of the page", function(assert)
	{
		AscTest.ClearDocument();

		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		AscTest.AddTextToParagraph(paragraph, "Some text");
		let pagePicture = CreateImage("Page picture");
		MakeFloating(pagePicture, true, 10);
		AddToParagraph(paragraph, pagePicture);

		let header = AscTest.CreateDefaultHeader(logicDocument.GetFinalSectPr());
		// the test header does not find the logic document by itself
		header.LogicDocument  = logicDocument;
		header.DrawingObjects = logicDocument.DrawingObjects;
		let logo = CreateImage("Logo");
		MakeFloating(logo, false, 50);
		AddToParagraph(header.GetElement(0), logo);
		AscTest.Recalculate();

		let items = GetPageDrawingObjects(0);
		assert.deepEqual(items.map(function(item) { return item["Name"]; }), ["Page picture", "Logo"],
			"The header object comes after the page objects, as the header is drawn first");
		assert.deepEqual(items.map(function(item) { return item["HdrFtr"]; }), [false, true], "The header object is marked");

		assert.ok(SelectDrawingObjectById(logo.GetId(), 0), "The header object is selected");
		assert.ok(logicDocument.DrawingObjects.getSelectedArray().map(function(o) { return o.parent; }).indexOf(logo) !== -1, "The logo is the selected object");

		logicDocument.GetFinalSectPr().Set_Header_Default(null);
	});
	QUnit.test("A header object is selected on the page asked for", function(assert)
	{
		AscTest.ClearDocument();

		let firstParagraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(firstParagraph);
		AscTest.AddTextToParagraph(firstParagraph, "First page");
		let secondParagraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(secondParagraph);
		AscTest.AddTextToParagraph(secondParagraph, "Second page");
		secondParagraph.Set_PageBreakBefore(true);

		let header = AscTest.CreateDefaultHeader(logicDocument.GetFinalSectPr());
		header.LogicDocument  = logicDocument;
		header.DrawingObjects = logicDocument.DrawingObjects;
		let logo = CreateImage("Logo");
		MakeFloating(logo, false, 50);
		AddToParagraph(header.GetElement(0), logo);
		AscTest.Recalculate();
		assert.strictEqual(logicDocument.GetPagesCount(), 2, "The document has two pages");

		AscTest.MoveCursorToParagraph(firstParagraph, true);
		assert.strictEqual(logicDocument.CurPage, 0, "The cursor is on the first page");
		assert.deepEqual(GetPageDrawingObjects(1).map(function(item) { return item["Name"]; }), ["Logo"], "The logo is listed on the second page");

		scrolledTo = null;
		assert.ok(SelectDrawingObjectById(logo.GetId(), 1), "The header object is selected");
		assert.strictEqual(scrolledTo, 1, "The view scrolls to the second page");
		assert.ok(logicDocument.DrawingObjects.getSelectedArray().map(function(o) { return o.parent; }).indexOf(logo) !== -1, "The logo is the selected object");
		assert.strictEqual(logo.GraphicObj.selectStartPage, 1, "The logo is selected on the second page");

		logicDocument.GetFinalSectPr().Set_Header_Default(null);
	});
	QUnit.test("A picture inside a text box is not listed on its own", function(assert)
	{
		AscTest.ClearDocument();

		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		AscTest.AddTextToParagraph(paragraph, "Some text");

		let textBox = CreateTextBox("Text box");
		MakeFloating(textBox, false, 20);
		let innerPicture = CreateImage("Picture in the text box");
		AddToParagraph(textBox.GraphicObj.textBoxContent.GetElement(0), innerPicture);
		AddToParagraph(paragraph, textBox);
		AscTest.Recalculate();

		assert.ok(innerPicture.isShapeChild(), "The picture belongs to the text box");
		assert.deepEqual(GetPageDrawingObjects(0).map(function(item) { return item["Name"]; }), ["Text box"],
			"Only the text box is listed");
	});
	QUnit.test("The objects of a page are listed from the front to the back", function(assert)
	{
		AscTest.ClearDocument();

		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		AscTest.AddTextToParagraph(paragraph, "Some text");

		let inline = CreateImage("Inline picture");
		let behind = CreateImage("Picture behind text");
		let front  = CreateTextBox("Text box in front");
		let middle = CreateImage("Picture in the middle");

		MakeFloating(behind, true, 10);
		MakeFloating(front, false, 30);
		MakeFloating(middle, false, 20);

		AddToParagraph(paragraph, inline);
		AddToParagraph(paragraph, behind);
		AddToParagraph(paragraph, front);
		AddToParagraph(paragraph, middle);
		AscTest.Recalculate();

		let items = GetPageDrawingObjects(0);
		assert.deepEqual(items.map(function(item) { return item["Name"]; }),
			["Text box in front", "Picture in the middle", "Inline picture", "Picture behind text"],
			"Front objects come first, sorted by their order, then inline, then behind the text");
		assert.deepEqual(items.map(function(item) { return item["Layer"]; }),
			["front", "front", "inline", "behind"], "Each object reports its layer");
		assert.deepEqual(items.map(function(item) { return item["Type"]; }),
			["textbox", "image", "image", "image"], "Each object reports its type");
		assert.strictEqual(items[0]["Id"], front.GetId(), "The id is the id of the drawing");
		assert.ok(items.every(function(item) { return !item["HdrFtr"] && !item["Selected"]; }), "Nothing is in a header or selected");

		assert.deepEqual(GetPageDrawingObjects(1), [], "A page that does not exist has no objects");
	});

	QUnit.test("An object behind the text can be selected from the list", function(assert)
	{
		AscTest.ClearDocument();

		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		AscTest.AddTextToParagraph(paragraph, "Some text");

		let behind = CreateImage("Picture behind text");
		let front  = CreateImage("Picture in front");
		MakeFloating(behind, true, 10);
		MakeFloating(front, false, 20);
		AddToParagraph(paragraph, behind);
		AddToParagraph(paragraph, front);
		AscTest.Recalculate();

		scrolledTo = null;
		assert.ok(SelectDrawingObjectById(behind.GetId(), 0), "The object is selected");
		assert.strictEqual(scrolledTo, 0, "The view scrolls to the page of the object");
		let selected = logicDocument.GetSelectedDrawingObjects ? logicDocument.GetSelectedDrawingObjects() : logicDocument.DrawingObjects.getSelectedArray().map(function(o) { return o.parent; });
		assert.deepEqual(selected, [behind], "Only the object behind the text is selected");

		let items = GetPageDrawingObjects(0);
		assert.deepEqual(items.map(function(item) { return item["Selected"]; }), [false, true], "The list marks the selected object");

		assert.notOk(SelectDrawingObjectById("no-such-id", 0), "An unknown id selects nothing");
	});

});
