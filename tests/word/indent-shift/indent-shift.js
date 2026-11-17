/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	function setup()
	{
		AscTest.ClearDocument();
		let paras = [];
		[[10, -5], [20, -5], [30, 0]].forEach(function (ind, i)
		{
			let p = AscTest.CreateParagraph();
			AscTest.AddTextToParagraph(p, "Ebene " + (i + 1));
			p.Set_Ind({Left : ind[0], FirstLine : ind[1], Right : 0}, false);
			logicDocument.PushToContent(p);
			paras.push(p);
		});
		AscTest.Recalculate();
		return paras;
	}

	function inds(paras)
	{
		return paras.map(function (p)
		{
			let ind = p.Get_CompiledPr2(false).ParaPr.Ind;
			return [Math.round(ind.Left * 100) / 100, Math.round(ind.FirstLine * 100) / 100, Math.round(ind.Right * 100) / 100];
		});
	}

	QUnit.module("Move indents together");

	QUnit.test("All selected paragraphs move by the same distance and keep their differences", function (assert)
	{
		let paras = setup();
		AscTest.SelectDocumentRange(0, 2);
		assert.true(logicDocument.ShiftParagraphIndents(5, 2, 5), "moved");
		assert.deepEqual(inds(paras), [[15, -5, 2], [25, -5, 2], [35, 0, 2]], "left +5, first line moves with it, right +2");
	});

	QUnit.test("Only the first line marker moved", function (assert)
	{
		let paras = setup();
		AscTest.SelectDocumentRange(0, 2);
		logicDocument.ShiftParagraphIndents(0, 0, 3);
		assert.deepEqual(inds(paras), [[10, -2, 0], [20, -2, 0], [30, 3, 0]], "only the first lines move by 3");
	});

	QUnit.test("The hanging marker: left moves, first line stays where it is", function (assert)
	{
		let paras = setup();
		AscTest.SelectDocumentRange(0, 2);
		logicDocument.ShiftParagraphIndents(4, 0, 0);
		assert.deepEqual(inds(paras), [[14, -9, 0], [24, -9, 0], [34, -4, 0]], "left +4, first line position unchanged");
	});

	QUnit.test("With one paragraph nothing special happens", function (assert)
	{
		let paras = setup();
		AscTest.MoveCursorToParagraph(paras[1]);
		assert.false(logicDocument.ShiftParagraphIndents(5, 0, 5), "not used for one paragraph");
		assert.deepEqual(inds(paras)[1], [20, -5, 0], "unchanged");
	});
});
