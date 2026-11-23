/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	const A4 = {W : 210, H : 297};

	// Three sections with two paragraphs each (paragraphs 0-1, 2-3, 4-5) and margins 10, 11 and 12
	function setup(landscapeSecond)
	{
		AscTest.ClearDocument();
		let emptyCount = logicDocument.GetElementsCount();
		let sectPrs = [];
		for (let section = 0; section < 3; ++section)
		{
			let isLast = 2 === section;
			let sectPr = isLast ? logicDocument.GetFinalSectPr() : new AscWord.SectPr(logicDocument);
			for (let i = 0; i < 2; ++i)
			{
				let p = AscTest.CreateParagraph();
				AscTest.AddTextToParagraph(p, "Section " + (section + 1) + " paragraph " + (i + 1));
				logicDocument.PushToContent(p);
				if (!isLast && 1 === i)
					p.SetSectionPr(sectPr);
			}
			if (1 === section && landscapeSecond)
				sectPr.SetPageSize(A4.H, A4.W);
			else
				sectPr.SetPageSize(A4.W, A4.H);

			let m = 10 + section;
			sectPr.SetPageMargins(m, m, m, m);
			sectPrs.push(sectPr);
		}
		if (emptyCount > 0)
			logicDocument.RemoveFromContent(0, emptyCount);
		AscTest.Recalculate();
		logicDocument.RemoveSelection();
		AscCommon.History.Clear();
		return sectPrs;
	}

	function margins(sectPr)
	{
		return [sectPr.GetPageMarginLeft(), sectPr.GetPageMarginTop(), sectPr.GetPageMarginRight(), sectPr.GetPageMarginBottom()].map(function (v)
		{
			return Math.round(v * 100) / 100;
		});
	}

	function size(sectPr)
	{
		return [Math.round(sectPr.GetPageWidth()), Math.round(sectPr.GetPageHeight())];
	}

	// What the Margins menu of the Layout tab sends
	function setMargins(value)
	{
		let props = new Asc.CDocumentSectionProps();
		props.put_TopMargin(value);
		props.put_LeftMargin(value);
		props.put_BottomMargin(value);
		props.put_RightMargin(value);
		logicDocument.Set_SectionProps(props);
	}

	// What the "Custom margins" dialog sends: all values of the current section with the new margins
	function setMarginsByDialog(value, orientation)
	{
		let props = new Asc.CDocumentSectionProps(logicDocument.GetCurrentSectPr(), logicDocument);
		if (undefined !== orientation && orientation !== props.get_Orientation())
		{
			let h = props.get_H();
			props.put_H(props.get_W());
			props.put_W(h);
			props.put_Orientation(orientation);
		}
		props.put_TopMargin(value);
		props.put_LeftMargin(value);
		props.put_BottomMargin(value);
		props.put_RightMargin(value);
		logicDocument.Set_SectionProps(props);
	}

	// What the API does for the Orientation and Size menus (change_PageOrient, change_DocSize)
	function setOrientation(orientation)
	{
		logicDocument.StartAction(AscDFH.historydescription_Document_SetPageOrientation);
		logicDocument.Set_DocumentOrientation(orientation);
		logicDocument.FinalizeAction();
	}

	// change_DocSize gets the portrait size and turns it for a landscape current section
	function setPageSize(w, h)
	{
		let isPortrait = logicDocument.GetCurrentSectPr().GetOrientation() === Asc.c_oAscPageOrientation.PagePortrait;
		logicDocument.StartAction(AscDFH.historydescription_Document_SetPageSize);
		if (isPortrait)
			logicDocument.Set_DocumentPageSize(w, h);
		else
			logicDocument.Set_DocumentPageSize(h, w);
		logicDocument.FinalizeAction();
	}

	QUnit.module("Page setup of the selected sections");

	QUnit.test("Select all + margins: all three sections", function (assert)
	{
		let sectPrs = setup();
		logicDocument.SelectAll();
		setMargins(25);
		assert.deepEqual(sectPrs.map(margins), [[25, 25, 25, 25], [25, 25, 25, 25], [25, 25, 25, 25]], "all sections have the new margins");
	});

	QUnit.test("Selection over sections 2 and 3: only these two", function (assert)
	{
		let sectPrs = setup();
		AscTest.SelectDocumentRange(3, 4);
		setMargins(25);
		assert.deepEqual(sectPrs.map(margins), [[10, 10, 10, 10], [25, 25, 25, 25], [25, 25, 25, 25]], "section 1 unchanged");
	});

	QUnit.test("Cursor in section 2 without selection: only section 2", function (assert)
	{
		let sectPrs = setup();
		AscTest.MoveCursorToParagraph(logicDocument.GetElement(2));
		setMargins(25);
		assert.deepEqual(sectPrs.map(margins), [[10, 10, 10, 10], [25, 25, 25, 25], [12, 12, 12, 12]], "only the current section");
	});

	QUnit.test("Selection inside one section: only that section", function (assert)
	{
		let sectPrs = setup();
		AscTest.SelectDocumentRange(4, 5);
		setMargins(25);
		assert.deepEqual(sectPrs.map(margins), [[10, 10, 10, 10], [11, 11, 11, 11], [25, 25, 25, 25]], "only section 3");
	});

	QUnit.test("One undo step brings back all sections", function (assert)
	{
		let sectPrs = setup();
		logicDocument.SelectAll();
		setMargins(25);
		logicDocument.Document_Undo();
		assert.deepEqual(sectPrs.map(margins), [[10, 10, 10, 10], [11, 11, 11, 11], [12, 12, 12, 12]], "all margins are back after one undo");
		logicDocument.Document_Redo();
		assert.deepEqual(sectPrs.map(margins), [[25, 25, 25, 25], [25, 25, 25, 25], [25, 25, 25, 25]], "redo");
	});

	QUnit.test("Custom margins dialog with select all: a landscape section stays landscape", function (assert)
	{
		let sectPrs = setup(true);
		logicDocument.SelectAll();
		setMarginsByDialog(20);
		assert.deepEqual(sectPrs.map(margins), [[20, 20, 20, 20], [20, 20, 20, 20], [20, 20, 20, 20]], "all sections have the new margins");
		assert.deepEqual(sectPrs.map(size), [[210, 297], [297, 210], [210, 297]], "every section keeps its own size and orientation");
	});

	QUnit.test("Custom margins dialog with a new orientation: all selected sections turn", function (assert)
	{
		let sectPrs = setup();
		AscTest.SelectDocumentRange(0, 3);
		setMarginsByDialog(20, Asc.c_oAscPageOrientation.PageLandscape);
		assert.deepEqual(sectPrs.map(size), [[297, 210], [297, 210], [210, 297]], "sections 1 and 2 landscape, section 3 unchanged");
		assert.deepEqual(sectPrs.map(margins), [[20, 20, 20, 20], [20, 20, 20, 20], [12, 12, 12, 12]], "margins of sections 1 and 2");
	});

	QUnit.test("Orientation menu follows the selection", function (assert)
	{
		let sectPrs = setup();
		logicDocument.SelectAll();
		setOrientation(Asc.c_oAscPageOrientation.PageLandscape);
		assert.deepEqual(sectPrs.map(size), [[297, 210], [297, 210], [297, 210]], "select all: all landscape");
		logicDocument.Document_Undo();
		assert.deepEqual(sectPrs.map(size), [[210, 297], [210, 297], [210, 297]], "one undo: all portrait again");

		AscTest.SelectDocumentRange(2, 5);
		setOrientation(Asc.c_oAscPageOrientation.PageLandscape);
		assert.deepEqual(sectPrs.map(size), [[210, 297], [297, 210], [297, 210]], "sections 2 and 3");
	});

	QUnit.test("Orientation menu without selection: only the current section", function (assert)
	{
		let sectPrs = setup();
		AscTest.MoveCursorToParagraph(logicDocument.GetElement(3));
		setOrientation(Asc.c_oAscPageOrientation.PageLandscape);
		assert.deepEqual(sectPrs.map(size), [[210, 297], [297, 210], [210, 297]], "only section 2");
	});

	QUnit.test("Size menu follows the selection, every section keeps its orientation", function (assert)
	{
		let sectPrs = setup(true);
		logicDocument.SelectAll();
		setPageSize(148, 210); // A5
		assert.deepEqual(sectPrs.map(size), [[148, 210], [210, 148], [148, 210]], "all sections A5, section 2 stays landscape");
		logicDocument.Document_Undo();
		assert.deepEqual(sectPrs.map(size), [[210, 297], [297, 210], [210, 297]], "one undo: all back");

		AscTest.SelectDocumentRange(3, 5);
		setPageSize(148, 210);
		assert.deepEqual(sectPrs.map(size), [[210, 297], [210, 148], [148, 210]], "sections 2 and 3");
	});

	QUnit.test("Size menu without selection: only the current section", function (assert)
	{
		let sectPrs = setup();
		AscTest.MoveCursorToParagraph(logicDocument.GetElement(1));
		setPageSize(148, 210);
		assert.deepEqual(sectPrs.map(size), [[148, 210], [210, 297], [210, 297]], "only section 1");
	});
});
