/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

// An internal link to a slide is stored as the target's position
// ("ppaction://hlinksldjumpslide<N>"), and saved to PPTX as "slide<N+1>.xml".
// These tests check that the link follows the slide it was made for when
// slides are inserted, moved, deleted or the change is undone (bug 73918).

$(function ()
{
	const logicDocument = AscTest.CreateLogicDocument();
	const SLIDE_JUMP = "ppaction://hlinksldjumpslide";

	// The real slide api aliases getLogicDocument to private_GetLogicDocument (slide/api.js);
	// the test editor keeps the base stub, which returns null.
	editor.WordControl.m_oLogicDocument = logicDocument;
	editor.getLogicDocument = editor.private_GetLogicDocument;

	function ResetPresentation(slideCount)
	{
		while (logicDocument.Slides.length)
			logicDocument.removeSlide(0);

		for (let i = 0; i < slideCount; ++i)
		{
			editor.WordControl.GoToPage(logicDocument.Slides.length - 1);
			logicDocument.addNextSlide(0);
		}
		editor.WordControl.Thumbnails.CalculatePlaces();
		AscCommon.History.Clear();
		return logicDocument.Slides.slice();
	}

	// The "New slide" command: adds a slide right after the given position.
	function AddSlideAfter(pos)
	{
		editor.WordControl.GoToPage(pos);
		logicDocument.addNextSlide(0);
	}

	function AddShape(slide)
	{
		const shapeTrack = new AscFormat.NewShapeTrack('rect', 0, 0, AscFormat.GetDefaultTheme(), null, null, null, 0);
		shapeTrack.track({}, 50, 50);
		const shape = shapeTrack.getShape(false, AscTest.DrawingDocument, null);
		shape.setBDeleted(false);
		shape.setParent(slide);
		slide.addToSpTreeToPos(slide.cSld.spTree.length, shape);
		return shape;
	}

	// Link a whole shape to a slide, the way Insert > Hyperlink does for a selected shape.
	function AddShapeLink(slide, targetIndex)
	{
		const shape = AddShape(slide);
		const hyper = new AscFormat.CT_Hyperlink();
		hyper.id = SLIDE_JUMP + targetIndex;
		shape.getCNvProps().setHlinkClick(hyper);
		return shape.getCNvProps().hlinkClick;
	}

	// Link a piece of text inside a shape to a slide.
	function AddTextLink(slide, targetIndex)
	{
		const shape = AddShape(slide);
		if (!shape.getDocContent())
			shape.createTextBody();

		const paragraph = shape.getDocContent().Content[0];
		const hyperlink = new AscCommonWord.ParaHyperlink();
		const run = new AscCommonWord.ParaRun(paragraph, false);
		run.AddText("go");
		hyperlink.AddToContent(0, run);
		paragraph.AddToContent(0, hyperlink);
		hyperlink.SetValue(SLIDE_JUMP + targetIndex);
		return hyperlink;
	}

	// Returns what the PPTX writer stores for a shape link: {id, action}.
	function WriteShapeLink(hlink)
	{
		const writer = new AscCommon.CBinaryFileWriter();
		const written = {};
		writer._WriteString2 = function (type, value)
		{
			if (type === 0)
				written.id = value;
			else if (type === 2)
				written.action = value;
		};
		writer.Write_Hyperlink2(hlink);
		return written;
	}

	// Returns what the PPTX writer stores for a text link: {id, action}.
	function WriteTextLink(hyperlink)
	{
		const writer = new AscCommon.CBinaryFileWriter();
		const written = {};
		writer._WriteString1 = function (type, value)
		{
			if (type === 0)
				written.id = value;
		};
		writer._WriteString2 = function (type, value)
		{
			if (type === 2)
				written.action = value;
		};
		writer.WriteHyperlink({Value : hyperlink.GetValue(), tooltip : hyperlink.ToolTip, slideJumpTargetId : hyperlink.slideJumpTargetId});
		return written;
	}

	QUnit.module("Internal links to slides", {
		beforeEach : function ()
		{
			AscCommon.History.Clear();
		}
	});

	QUnit.test("Shape link keeps its target when a slide is inserted before it", function (assert)
	{
		const slides = ResetPresentation(3);
		const hlink = AddShapeLink(slides[0], 2);

		AddSlideAfter(0);

		assert.strictEqual(logicDocument.Slides.indexOf(slides[2]), 3, "The target slide moved to position 3");
		assert.strictEqual(WriteShapeLink(hlink).id, "slide4.xml", "The saved file links to slide4.xml");
		assert.strictEqual(logicDocument.getSlideJumpTargetIndex(hlink), 3, "The link points to position 3");
		assert.strictEqual(logicDocument.resolveSlideJumpValue(hlink), SLIDE_JUMP + "3", "Click and tooltip use position 3");
	});

	QUnit.test("Text link keeps its target when a slide is inserted before it", function (assert)
	{
		const slides = ResetPresentation(3);
		const hyperlink = AddTextLink(slides[0], 2);

		AddSlideAfter(0);

		assert.strictEqual(WriteTextLink(hyperlink).id, "slide4.xml", "The saved file links to slide4.xml");
		assert.strictEqual(logicDocument.getSlideJumpTargetIndex(hyperlink), 3, "The link points to position 3");
	});

	QUnit.test("Link read from a file follows its target when slides are moved", function (assert)
	{
		const slides = ResetPresentation(4);
		// As read from PPTX: the value is set directly, without history, then bound on open.
		const shape = AddShape(slides[0]);
		const hyper = new AscFormat.CT_Hyperlink();
		hyper.id = SLIDE_JUMP + "3";
		shape.getCNvProps().hlinkClick = hyper;
		if (logicDocument.bindSlideJumpLinks)
			logicDocument.bindSlideJumpLinks();

		logicDocument.moveSlides([3], 1);

		assert.strictEqual(logicDocument.Slides.indexOf(slides[3]), 1, "The target slide moved to position 1");
		assert.deepEqual(WriteShapeLink(hyper), {id : "slide2.xml", action : "ppaction://hlinksldjump"}, "The saved file links to slide2.xml");
		assert.strictEqual(logicDocument.getSlideJumpTargetIndex(hyper), 1, "The link points to position 1");
	});

	QUnit.test("Link to a deleted slide is saved as no action", function (assert)
	{
		const slides = ResetPresentation(3);
		const hlink = AddShapeLink(slides[0], 2);

		logicDocument.deleteSlides([slides[2]]);

		assert.strictEqual(WriteShapeLink(hlink).action, "ppaction://noaction", "The saved file has no jump");
		assert.strictEqual(logicDocument.getSlideJumpTargetIndex(hlink), -1, "The link has no target");
		assert.strictEqual(logicDocument.resolveSlideJumpValue(hlink), "ppaction://noaction", "Click does nothing");
	});

	QUnit.test("Undoing the slide insertion restores the link position", function (assert)
	{
		const slides = ResetPresentation(3);
		const hlink = AddShapeLink(slides[0], 2);

		AddSlideAfter(0);
		assert.strictEqual(logicDocument.getSlideJumpTargetIndex(hlink), 3, "After insert: position 3");

		logicDocument.Document_Undo();
		assert.strictEqual(logicDocument.Slides.length, 3, "The inserted slide is gone");
		assert.strictEqual(logicDocument.getSlideJumpTargetIndex(hlink), 2, "After undo: position 2");
	});

	QUnit.test("Copied shape link points to the same slide", function (assert)
	{
		const slides = ResetPresentation(3);
		const hlink = AddShapeLink(slides[0], 2);
		const copy = hlink.createDuplicate();

		AddSlideAfter(0);

		assert.strictEqual(logicDocument.getSlideJumpTargetIndex(copy), 3, "The copy follows the same target slide");
	});
});
