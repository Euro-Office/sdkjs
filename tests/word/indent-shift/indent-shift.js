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

	QUnit.test("Indents that come from the levels of a multilevel list", function (assert)
	{
		AscTest.ClearDocument();
		let numberingManager = logicDocument.GetNumberingManager();
		let num = numberingManager.CreateNum();
		AscWord.GetNumberingObjectByDeprecatedTypes(2, 7).FillNum(num);
		numberingManager.AddNum(num);

		let paras = [];
		for (let lvl = 0; lvl < 3; ++lvl)
		{
			let p = AscTest.CreateParagraph();
			AscTest.AddTextToParagraph(p, "Ebene " + (lvl + 1));
			p.SetNumPr(num.GetId(), lvl);
			logicDocument.PushToContent(p);
			paras.push(p);
		}
		AscTest.Recalculate();

		let before = inds(paras);
		assert.notEqual(before[0][0], before[1][0], "the levels have different indents");
		assert.equal(paras[0].Pr.Ind.Left, undefined, "no direct indent before");

		AscTest.SelectDocumentRange(0, 2);
		assert.true(logicDocument.ShiftParagraphIndents(5, 0, 5), "moved");
		assert.deepEqual(inds(paras), before.map(function (ind)
		{
			return [Math.round((ind[0] + 5) * 100) / 100, ind[1], ind[2]];
		}), "every level moves by 5, the hanging indents stay");
	});

	QUnit.test("One action: a single undo restores all paragraphs", function (assert)
	{
		let paras = setup();
		AscTest.SelectDocumentRange(0, 2);
		logicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphIndentFromRulers);
		logicDocument.ShiftParagraphIndents(5, 2, 5);
		logicDocument.FinalizeAction();
		assert.deepEqual(inds(paras), [[15, -5, 2], [25, -5, 2], [35, 0, 2]], "moved");

		logicDocument.Document_Undo();
		assert.deepEqual(inds(paras), [[10, -5, 0], [20, -5, 0], [30, 0, 0]], "all three back after one undo");
	});

	QUnit.test("Each paragraph stays inside the page and the text area", function (assert)
	{
		let paras = setup();
		let p = paras[0];
		let textWidth = p.XLimit - p.X;
		let startGap  = p.X;
		let endGap    = logicDocument.Get_PageLimits(0).XLimit - p.XLimit;
		assert.true(textWidth > 100, "the paragraphs are laid out (text width " + textWidth + " mm)");

		AscTest.SelectDocumentRange(0, 2);
		logicDocument.ShiftParagraphIndents(500, 0, 500);
		let r = inds(paras);
		assert.deepEqual(r.map(function (ind) {return ind[0];}), [textWidth - 20, textWidth - 20, textWidth - 20].map(function (v) {return Math.round(v * 100) / 100;}), "left stops 20 mm before the right indent");
		assert.deepEqual(r.map(function (ind) {return ind[1];}), [-5, -5, 0], "the hanging indents stay");

		paras = setup();
		AscTest.SelectDocumentRange(0, 2);
		logicDocument.ShiftParagraphIndents(-500, 0, -500);
		r = inds(paras);
		// the first line of the first two paragraphs is 5 mm left of their left indent; it stops at the page edge
		let round = function (v) {return Math.round(v * 100) / 100;};
		assert.deepEqual(r, [[round(5 - startGap), -5, 0], [round(5 - startGap), -5, 0], [round(-startGap), 0, 0]], "the first line stops at the left page edge");

		paras = setup();
		AscTest.SelectDocumentRange(0, 2);
		logicDocument.ShiftParagraphIndents(0, -500, 0);
		assert.deepEqual(inds(paras).map(function (ind) {return ind[2];}), [round(-endGap), round(-endGap), round(-endGap)], "the right indent stops at the right page edge");

		paras = setup();
		AscTest.SelectDocumentRange(0, 2);
		logicDocument.ShiftParagraphIndents(0, 0, 500);
		assert.deepEqual(inds(paras).map(function (ind) {return round(ind[0] + ind[1]);}), [round(textWidth - 20), round(textWidth - 20), round(textWidth - 20)], "the first line alone stops 20 mm before the right indent");
	});

	QUnit.test("Paragraphs with the other direction are left as they are", function (assert)
	{
		let paras = setup();
		paras[1].SetParagraphBidi(true);
		AscTest.SelectDocumentRange(0, 2);
		logicDocument.ShiftParagraphIndents(5, 0, 5, false);
		assert.deepEqual(inds(paras), [[15, -5, 0], [20, -5, 0], [35, 0, 0]], "the RTL paragraph does not move");
	});

	function createRuler()
	{
		let ruler = new CHorRuler();
		ruler.m_oWordControl = {m_oLogicDocument : logicDocument, m_oApi : AscTest.Editor};
		return ruler;
	}

	function dragRuler(ruler, start, left, right, first)
	{
		ruler.m_oIndentAtDragStart = {left : start[0], right : start[1], first : start[2]};
		ruler.m_dIndentLeft      = left;
		ruler.m_dIndentRight     = right;
		ruler.m_dIndentLeftFirst = first;
		ruler.SetPrProperties();
	}

	QUnit.test("Ruler with several paragraphs: differences stay, one undo step", function (assert)
	{
		let paras = setup();
		AscTest.SelectDocumentRange(0, 2);
		// the ruler shows the first paragraph (left 10, first line at 5); the left marker is dragged by 4 mm
		dragRuler(createRuler(), [10, 0, 5], 14, 0, 5);
		assert.deepEqual(inds(paras), [[14, -9, 0], [24, -9, 0], [34, -4, 0]], "every left indent +4, first lines stay");

		logicDocument.Document_Undo();
		assert.deepEqual(inds(paras), [[10, -5, 0], [20, -5, 0], [30, 0, 0]], "one undo restores all");
	});

	QUnit.test("Ruler with one paragraph works as before", function (assert)
	{
		let paras = setup();
		AscTest.MoveCursorToParagraph(paras[1]);
		// left marker from 20 to 30, the first line stays at 15, right marker from 0 to 3
		dragRuler(createRuler(), [20, 0, 15], 30, 3, 15);
		assert.deepEqual(inds(paras), [[10, -5, 0], [30, -15, 3], [30, 0, 0]], "only this paragraph gets exactly the ruler values");

		// without a recorded drag start the ruler also sets the values directly
		paras = setup();
		AscTest.MoveCursorToParagraph(paras[1]);
		let ruler = createRuler();
		ruler.m_dIndentLeft      = 25;
		ruler.m_dIndentRight     = 0;
		ruler.m_dIndentLeftFirst = 25;
		ruler.SetPrProperties();
		assert.deepEqual(inds(paras)[1], [25, 0, 0], "ruler values set");
	});
});
