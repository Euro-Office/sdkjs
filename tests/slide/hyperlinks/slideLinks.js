/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

// A link to a slide is stored as the target's position ("ppaction://hlinksldjumpslide<N>")
// and saved to PPTX as "slide<N+1>.xml". When slides are inserted, moved or deleted, the
// stored position must be rewritten so the link still opens the slide it was made for.

$(function ()
{
	const logicDocument = AscTest.CreateLogicDocument();
	const SLIDE_JUMP = "ppaction://hlinksldjumpslide";

	// The slide api aliases getLogicDocument to private_GetLogicDocument; the test editor
	// keeps the base stub, which returns null.
	editor.WordControl.m_oLogicDocument = logicDocument;
	editor.getLogicDocument = editor.private_GetLogicDocument;

	// Builds a presentation with the given slides and returns them in order.
	function MakeSlides(count)
	{
		while (logicDocument.Slides.length)
			logicDocument.removeSlide(0);

		for (let i = 0; i < count; ++i)
		{
			editor.WordControl.GoToPage(logicDocument.Slides.length - 1);
			logicDocument.addNextSlide(0);
		}
		editor.WordControl.Thumbnails.CalculatePlaces();
		AscCommon.History.Clear();
		return logicDocument.Slides.slice();
	}

	// "New slide" on the slide at the given position: the new slide goes right after it.
	function NewSlideAfter(position)
	{
		editor.WordControl.GoToPage(position);
		logicDocument.addNextSlide(0);
	}

	function AddRectangle(slide)
	{
		const track = new AscFormat.NewShapeTrack('rect', 0, 0, AscFormat.GetDefaultTheme(), null, null, null, 0);
		track.track({}, 40, 40);
		const shape = track.getShape(false, AscTest.DrawingDocument, null);
		shape.setBDeleted(false);
		shape.setParent(slide);
		slide.addToSpTreeToPos(slide.cSld.spTree.length, shape);
		return shape;
	}

	function LinkShape(slide, position)
	{
		const shape = AddRectangle(slide);
		const link = new AscFormat.CT_Hyperlink();
		link.id = SLIDE_JUMP + position;
		shape.getCNvProps().setHlinkClick(link);
		return shape;
	}

	function AddLinkToParagraph(paragraph, position)
	{
		const hyperlink = new AscCommonWord.ParaHyperlink();
		const run = new AscCommonWord.ParaRun(paragraph, false);
		run.AddText("link");
		hyperlink.AddToContent(0, run);
		paragraph.AddToContent(0, hyperlink);
		hyperlink.SetValue(SLIDE_JUMP + position);
		return hyperlink;
	}

	function LinkText(slide, position)
	{
		const shape = AddRectangle(slide);
		if (!shape.getDocContent())
			shape.createTextBody();
		return AddLinkToParagraph(shape.getDocContent().Content[0], position);
	}

	function FindTextLinks(slide)
	{
		const links = [];
		slide.cSld.spTree.forEach(function (drawing)
		{
			const content = drawing.getDocContent && drawing.getDocContent();
			if (!content)
				return;
			content.GetAllParagraphs(undefined, []).forEach(function (paragraph)
			{
				paragraph.Content.forEach(function (item)
				{
					if (item instanceof AscCommonWord.ParaHyperlink)
						links.push(item);
				});
			});
		});
		return links;
	}

	// What the PPTX writer stores for a shape link.
	function SavedTarget(shape)
	{
		const writer = new AscCommon.CBinaryFileWriter();
		let target = null;
		writer._WriteString2 = function (type, value)
		{
			if (type === 0)
				target = value;
		};
		writer.Write_Hyperlink2(shape.getCNvProps().hlinkClick);
		return target;
	}

	QUnit.module("Links to slides follow their target slide", {
		beforeEach : function ()
		{
			AscCommon.History.Clear();
		}
	});

	QUnit.test("Shape link after a slide is inserted before the target", function (assert)
	{
		const slides = MakeSlides(3);
		const shape = LinkShape(slides[0], 2);

		NewSlideAfter(1);

		assert.strictEqual(logicDocument.Slides.indexOf(slides[2]), 3, "The target is now at position 3");
		assert.strictEqual(shape.getCNvProps().hlinkClick.id, SLIDE_JUMP + "3", "The link points to position 3");
		assert.strictEqual(SavedTarget(shape), "slide4.xml", "The saved file links to slide4.xml");
	});

	QUnit.test("Text link after a slide is inserted before the target", function (assert)
	{
		const slides = MakeSlides(3);
		const hyperlink = LinkText(slides[0], 2);

		NewSlideAfter(1);

		assert.strictEqual(hyperlink.GetValue(), SLIDE_JUMP + "3", "The link points to position 3");

		const props = [];
		const syncProps = editor.sync_HyperlinkPropCallback;
		editor.sync_HyperlinkPropCallback = function (hyperProps) {props.push(hyperProps);};
		try
		{
			hyperlink.Document_UpdateInterfaceState();
		}
		finally
		{
			editor.sync_HyperlinkPropCallback = syncProps;
		}
		assert.strictEqual(props[0] && props[0].get_Value(), SLIDE_JUMP + "3", "The link dialog shows position 3");
		assert.strictEqual(hyperlink.Copy(false, {}).GetValue(), SLIDE_JUMP + "3", "A copy points to position 3");
	});

	QUnit.test("Links before the target and to other slides", function (assert)
	{
		const slides = MakeSlides(4);
		const toFirst = LinkShape(slides[3], 0);
		const toLast = LinkShape(slides[0], 3);
		const toSecond = LinkShape(slides[2], 1);

		NewSlideAfter(1);

		assert.strictEqual(toFirst.getCNvProps().hlinkClick.id, SLIDE_JUMP + "0", "A link to a slide before the insertion is unchanged");
		assert.strictEqual(toSecond.getCNvProps().hlinkClick.id, SLIDE_JUMP + "1", "A link to the slide the new one follows is unchanged");
		assert.strictEqual(toLast.getCNvProps().hlinkClick.id, SLIDE_JUMP + "4", "A link to a slide after the insertion moves on by one");
	});

	QUnit.test("Moving the target slide", function (assert)
	{
		const slides = MakeSlides(4);
		const shape = LinkShape(slides[0], 3);

		logicDocument.shiftSlides(1, [3]);

		assert.strictEqual(logicDocument.Slides.indexOf(slides[3]), 1, "The target is now at position 1");
		assert.strictEqual(shape.getCNvProps().hlinkClick.id, SLIDE_JUMP + "1", "The link points to position 1");
	});

	QUnit.test("Deleting a slide before the target", function (assert)
	{
		const slides = MakeSlides(4);
		const shape = LinkShape(slides[0], 3);

		logicDocument.deleteSlides([slides[1]]);

		assert.strictEqual(shape.getCNvProps().hlinkClick.id, SLIDE_JUMP + "2", "The link points to position 2");
	});

	QUnit.test("A link to a deleted slide keeps its number", function (assert)
	{
		const slides = MakeSlides(3);
		const shape = LinkShape(slides[0], 2);

		logicDocument.deleteSlides([slides[2]]);

		assert.strictEqual(shape.getCNvProps().hlinkClick.id, SLIDE_JUMP + "2", "The value is unchanged");
	});

	QUnit.test("Undo restores the slide and the link", function (assert)
	{
		const slides = MakeSlides(3);
		const shape = LinkShape(slides[0], 2);
		AscCommon.History.Clear();

		NewSlideAfter(1);
		assert.strictEqual(shape.getCNvProps().hlinkClick.id, SLIDE_JUMP + "3", "After the insertion: position 3");

		logicDocument.Document_Undo();
		assert.strictEqual(logicDocument.Slides.length, 3, "The new slide is gone");
		assert.strictEqual(shape.getCNvProps().hlinkClick.id, SLIDE_JUMP + "2", "After undo: position 2");

		logicDocument.Document_Redo();
		assert.strictEqual(logicDocument.Slides.length, 4, "Redo brings the slide back");
		assert.strictEqual(shape.getCNvProps().hlinkClick.id, SLIDE_JUMP + "3", "After redo: position 3");
	});

	QUnit.test("Duplicating a slide that holds a link", function (assert)
	{
		const slides = MakeSlides(3);
		LinkText(slides[0], 2);

		logicDocument.shiftSlides(1, [0], true);

		const copy = logicDocument.Slides[1];
		assert.notStrictEqual(copy, slides[0], "A copy was inserted at position 1");
		assert.deepEqual(FindTextLinks(slides[0]).map(function (h) {return h.GetValue();}), [SLIDE_JUMP + "3"], "The original link points to position 3");
		assert.deepEqual(FindTextLinks(copy).map(function (h) {return h.GetValue();}), [SLIDE_JUMP + "3"], "The copied link points to position 3");
	});

	QUnit.test("Links in every cell of a table", function (assert)
	{
		const slides = MakeSlides(3);
		editor.WordControl.GoToPage(0);
		logicDocument.Add_FlowTable(2, 2);
		const frame = slides[0].cSld.spTree.find(function (drawing)
		{
			return drawing.getObjectType() === AscDFH.historyitem_type_GraphicFrame;
		});
		const table = frame.graphicObject;
		const inFirstCell = AddLinkToParagraph(table.Content[0].Get_Cell(0).Content.Content[0], 2);
		const inLastCell = AddLinkToParagraph(table.Content[1].Get_Cell(1).Content.Content[0], 2);
		table.CurCell = table.Content[0].Get_Cell(0);
		table.Selection.Use = false;

		NewSlideAfter(1);

		assert.strictEqual(inFirstCell.GetValue(), SLIDE_JUMP + "3", "Link in the cell with the cursor");
		assert.strictEqual(inLastCell.GetValue(), SLIDE_JUMP + "3", "Link in another cell");
	});

	QUnit.test("Reading the position from a link value", function (assert)
	{
		const read = AscCommonSlide.GetSlideJumpPosition;
		assert.strictEqual(read(SLIDE_JUMP + "0"), 0, "Position 0");
		assert.strictEqual(read(SLIDE_JUMP + "12"), 12, "Position 12");
		assert.strictEqual(read(SLIDE_JUMP), null, "No number");
		assert.strictEqual(read(SLIDE_JUMP + "2x"), null, "Not a number");
		assert.strictEqual(read("ppaction://hlinkshowjump?jump=nextslide"), null, "Another action");
		assert.strictEqual(read("https://example.com"), null, "An external link");
		assert.strictEqual(read(null), null, "No value");
	});
});
