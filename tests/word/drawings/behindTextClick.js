/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function () {

	let logicDocument = AscTest.CreateLogicDocument();
	// the test drawing document has no mobile check, which the table border hit test asks for
	let drawingDocument = logicDocument.GetDrawingDocument();
	if (!drawingDocument.IsMobileVersion)
		drawingDocument.IsMobileVersion = function() { return false; };
	// the border drag also locks the mouse cursor, which the test drawing document does not have
	["LockCursorTypeCur", "UpdateTableRuler", "LockCursorType", "UnlockCursorType", "SetCursorType"].forEach(function(name)
	{
		if (!drawingDocument[name])
			drawingDocument[name] = function() {};
	});

	function CreateTextBox(w, h)
	{
		let shapeTrack = new AscFormat.NewShapeTrack("rect", 0, 0, logicDocument.theme, null, null, null, 0);
		shapeTrack.track({}, w, h);
		let shape = shapeTrack.getShape(true, logicDocument.GetDrawingDocument(), null);
		shape.setBDeleted(false);
		shape.createTextBoxContent();
		shape.textBoxContent.LogicDocument  = logicDocument;
		shape.textBoxContent.DrawingObjects = logicDocument.DrawingObjects;
		shape.textBoxContent.GetElement(0).LogicDocument = logicDocument;

		let drawing = new ParaDrawing(w, h, null, logicDocument.GetDrawingDocument(), logicDocument, null);
		shape.setParent(drawing);
		drawing.Set_GraphicObject(shape);
		drawing.Set_DrawingType(drawing_Anchor);
		drawing.Set_WrappingType(WRAPPING_TYPE_NONE);
		drawing.Set_BehindDoc(true);
		drawing.Set_PositionH(Asc.c_oAscRelativeFromH.Page, false, 0);
		drawing.Set_PositionV(Asc.c_oAscRelativeFromV.Page, false, 0);
		drawing.Set_RelativeHeight(10);
		return drawing;
	}

	function AddToParagraph(paragraph, drawing)
	{
		let run = new AscWord.CRun();
		paragraph.AddToContentToEnd(run);
		run.AddToContent(0, drawing);
	}

	// a page with text and a table, and a text box behind the text that covers the whole page
	function setup(withText)
	{
		AscTest.ClearDocument();
		let paragraph = AscTest.CreateParagraph();
		logicDocument.PushToContent(paragraph);
		AscTest.AddTextToParagraph(paragraph, "Text in front of the text box");

		let table = AscTest.CreateTable(2, 2, [60, 60]);
		logicDocument.PushToContent(table);
		AscTest.AddTextToParagraph(table.GetRow(0).GetCell(0).GetContent().GetElement(0), "Cell");
		logicDocument.PushToContent(AscTest.CreateParagraph());

		let textBox = CreateTextBox(210, 297);
		if (withText)
			AscTest.AddTextToParagraph(textBox.GraphicObj.textBoxContent.GetElement(0), "Text in the text box behind");
		AddToParagraph(paragraph, textBox);
		AscTest.Recalculate();
		logicDocument.RemoveSelection();
		logicDocument.SetDocPosType(docpostype_Content);
		return {paragraph : paragraph, table : table, textBox : textBox};
	}

	function click(x, y)
	{
		let e = {ClickCount : 1, Type : AscCommon.g_mouse_event_type_down, CtrlKey : false, ShiftKey : false, Button : 0};
		logicDocument.CurPage = 0;
		logicDocument.Selection_SetStart(x, y, e);
		e.Type = AscCommon.g_mouse_event_type_up;
		logicDocument.Selection_SetEnd(x, y, e);
		return logicDocument.GetDocPosType();
	}

	function textPoint(paragraph)
	{
		let line = paragraph.Lines[0];
		return {x : paragraph.Pages[0].X + 10, y : paragraph.Pages[0].Y + line.Y - 1};
	}

	function borderPoint(table)
	{
		let x = table.Pages[0].X + table.GetRow(0).GetCellInfo(1).X_grid_start;
		let y = table.RowsInfo[0].Y[0] + table.RowsInfo[0].H[0] / 2;
		return {x : x, y : y};
	}

	QUnit.module("Click over an object behind the text");

	[false, true].forEach(function(withText)
	{
		let name = withText ? " (text box with text)" : " (empty text box)";

		QUnit.test("A click on text goes to the text" + name, function (assert)
		{
			let s = setup(withText);
			let p = textPoint(s.paragraph);
			assert.strictEqual(logicDocument.DrawingObjects.IsInDrawingObject(p.x, p.y, 0, logicDocument), DRAWING_ARRAY_TYPE_BEHIND, "the point is over the text box behind the text");
			assert.strictEqual(click(p.x, p.y), docpostype_Content, "the click goes to the text");
		});

		QUnit.test("A click on a table border can move the border" + name, function (assert)
		{
			let s = setup(withText);
			let p = borderPoint(s.table);
			assert.notStrictEqual(logicDocument.IsTableBorder(p.x, p.y, 0), null, "the point is on the table border");
			assert.strictEqual(click(p.x, p.y), docpostype_Content, "the click goes to the table, not the text box");
			assert.strictEqual(logicDocument.DrawingObjects.getSelectedObjects().length, 0, "the text box is not selected");
		});

		QUnit.test("A click in a table cell goes to the cell" + name, function (assert)
		{
			let s = setup(withText);
			let p = borderPoint(s.table);
			assert.strictEqual(click(p.x + 20, p.y), docpostype_Content, "the click goes to the cell");
			assert.strictEqual(logicDocument.DrawingObjects.getSelectedObjects().length, 0, "the text box is not selected");
		});
	});

	QUnit.test("A click where there is no text or table still selects the text box behind", function (assert)
	{
		let s = setup(false);
		let page = logicDocument.Pages[0];
		assert.strictEqual(click(page.Width - 15, page.Height - 40), docpostype_DrawingObjects, "the text box is selected");
	});
});
