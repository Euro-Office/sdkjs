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

$(function ()
{
	const logicDocument = AscTest.CreateLogicDocument();
	const builderApi = AscBuilder.Slide.Api;
	if (AscBuilder.Slide.init)
		AscBuilder.Slide.init();

	const BY_PARAGRAPH = AscFormat.TEXT_SEQUENCE_BY_PARAGRAPH;
	const AS_ONE_OBJECT = AscFormat.TEXT_SEQUENCE_AS_ONE_OBJECT;
	const DEFAULT_TEXTS = ["Point A", "", "Point B", "Sub point B1", "Point C"];

	// Creates a new slide with a shape: paragraphs from aTexts, paragraph 3 on level 1 (if present)
	function CreateTextShape(aTexts)
	{
		const slideIndex = logicDocument.GetSlidesCount();
		logicDocument.addNextSlide(slideIndex);
		editor.WordControl.Thumbnails.CalculatePlaces();
		logicDocument.Set_CurPage(slideIndex);

		const apiSlide = builderApi.GetPresentation().GetSlideByIndex(slideIndex);
		const fill = builderApi.CreateSolidFill(builderApi.CreateRGBColor(200, 200, 200));
		const stroke = builderApi.CreateStroke(0, builderApi.CreateNoFill());
		const apiShape = builderApi.CreateShape("rect", 300 * 36000, 150 * 36000, fill, stroke);
		apiSlide.AddObject(apiShape);
		const apiContent = apiShape.GetDocContent();
		apiContent.RemoveAllElements();
		for (let i = 0; i < aTexts.length; ++i)
		{
			const apiParagraph = builderApi.CreateParagraph();
			if (aTexts[i])
				apiParagraph.AddText(aTexts[i]);
			apiContent.Push(apiParagraph);
		}
		apiContent.RemoveElement(0);

		const shape = apiShape.Shape;
		if (shape.getDocContent().Content.length > 3)
			shape.getDocContent().Content[3].Set_PresentationLevel(1);

		return {slide : logicDocument.Slides[slideIndex], shape : shape};
	}

	function SelectShape(slide, shape)
	{
		slide.graphicObjects.resetSelection();
		slide.graphicObjects.selectObject(shape, 0);
		if (slide.timing)
			slide.timing.onChangeDrawingsSelection();
	}

	function AddFade(slide, shape)
	{
		SelectShape(slide, shape);
		logicDocument.AddAnimation(AscFormat.PRESET_CLASS_ENTR, AscFormat.ENTRANCE_FADE, 0, null, false, false);
		return slide.timing;
	}

	function SetTextSequence(slide, nSequence)
	{
		const props = slide.timing.getAnimProperties();
		props.asc_putTextSequence(nSequence);
		logicDocument.SetAnimationProperties(props, false);
	}

	function GetRanges(slide)
	{
		return slide.timing.getAllAnimEffects().map(function (effect)
		{
			const range = effect.getParagraphRange();
			return range ? [range.st, range.end] : null;
		});
	}

	function GetParagraphBuilds(slide, shape)
	{
		const timing = slide.timing;
		if (!timing.bldLst)
			return null;

		return timing.bldLst.list.map(function (bld)
		{
			return {
				isBldP : bld.getObjectType() === AscDFH.historyitem_type_BldP,
				isShape: bld.spid === shape.GetId(),
				grpId  : bld.grpId,
				build  : bld.build
			};
		});
	}

	function CheckParagraphBuild(assert, slide, shape, nBuild, sMessage)
	{
		const grpId = slide.timing.getAllAnimEffects()[0].cTn.grpId;
		assert.deepEqual(GetParagraphBuilds(slide, shape), [{isBldP : true, isShape : true, grpId : grpId, build : nBuild}], sMessage);
	}

	function CreateParagraph(docContent, sText)
	{
		const paragraph = new AscWord.Paragraph(docContent, true);
		const run = new AscWord.CRun(paragraph, false);
		run.AddText(sText);
		paragraph.AddToContent(0, run);
		return paragraph;
	}

	function Undo()
	{
		logicDocument.Document_Undo();
	}

	QUnit.module("Animation: text by paragraph");

	QUnit.test("Paragraph ranges", function (assert)
	{
		const {shape} = CreateTextShape(DEFAULT_TEXTS);
		assert.deepEqual(AscFormat.getParagraphBuildRanges(shape, 1), [{st : 0, end : 0}, {st : 2, end : 3}, {st : 4, end : 4}], "Empty paragraphs are skipped, sub levels join their parent");
		assert.strictEqual(AscFormat.getParagraphBuildRanges(shape, 5).length, 4, "Every non-empty paragraph on its own up to level 5");
	});

	QUnit.test("Switch to by paragraph and back", function (assert)
	{
		const {slide, shape} = CreateTextShape(DEFAULT_TEXTS);
		const timing = AddFade(slide, shape);
		assert.strictEqual(timing.getAllAnimEffects().length, 1, "One effect after adding");
		assert.strictEqual(timing.getAnimProperties().asc_getTextSequence(), AS_ONE_OBJECT, "Reported as one object");
		assert.notOk(timing.bldLst, "No bldLst for a whole object effect");

		SelectShape(slide, shape);
		SetTextSequence(slide, BY_PARAGRAPH);
		const effects = timing.getAllAnimEffects();
		assert.deepEqual(GetRanges(slide), [[0, 0], [2, 3], [4, 4]], "One effect per paragraph range");
		assert.ok(effects.every(function (effect) {return effect.cTn.grpId === effects[0].cTn.grpId;}), "Same group id");
		assert.ok(effects.every(function (effect) {return effect.cTn.nodeType === AscFormat.NODE_TYPE_CLICKEFFECT;}), "Every paragraph on click");
		assert.deepEqual(effects.map(function (effect) {return effect.getIndexInSequence();}), [1, 2, 3], "Separate clicks");
		assert.ok(effects.every(function (effect) {return effect.isSelected();}), "All paragraph effects stay selected");
		effects.forEach(function (effect)
		{
			let all = 0, withRange = 0;
			effect.traverse(function (child)
			{
				if (child.getObjectType() === AscDFH.historyitem_type_SpTgt)
				{
					++all;
					if (child.txEl && child.txEl.pRg)
						++withRange;
				}
				return false;
			});
			assert.ok(all > 0 && all === withRange, "Every target of the effect has the paragraph range");
		});
		CheckParagraphBuild(assert, slide, shape, AscFormat.ParaBuildType_p, "bldP with build=p");
		SelectShape(slide, shape);
		assert.strictEqual(timing.getAnimProperties().asc_getTextSequence(), BY_PARAGRAPH, "Reported as by paragraph");
		assert.deepEqual(effects.map(function (effect) {return effect.getObjectText();}), ["Point A", "Point B Sub point B1", "Point C"], "Animation pane shows the paragraph text");

		timing.resetSelection();
		timing.getAllAnimEffects()[1].select();
		assert.strictEqual(timing.getAnimProperties().asc_getTextSequence(), BY_PARAGRAPH, "A single paragraph effect is by paragraph");
		SetTextSequence(slide, AS_ONE_OBJECT);
		assert.deepEqual(GetRanges(slide), [null], "The whole group collapses into one effect");
		assert.notOk(timing.bldLst, "bldLst removed");
	});

	QUnit.test("Replace effect and change subtype keep the paragraphs", function (assert)
	{
		const {slide, shape} = CreateTextShape(DEFAULT_TEXTS);
		const timing = AddFade(slide, shape);
		SelectShape(slide, shape);
		SetTextSequence(slide, BY_PARAGRAPH);

		SelectShape(slide, shape);
		logicDocument.AddAnimation(AscFormat.PRESET_CLASS_ENTR, AscFormat.ENTRANCE_FLY_IN_FROM, 4, null, true, false);
		assert.deepEqual(GetRanges(slide), [[0, 0], [2, 3], [4, 4]], "Ranges after replace");
		assert.ok(timing.getAllAnimEffects().every(function (effect) {return effect.cTn.presetID === AscFormat.ENTRANCE_FLY_IN_FROM;}), "Preset replaced");
		CheckParagraphBuild(assert, slide, shape, AscFormat.ParaBuildType_p, "bldP kept after replace");

		SelectShape(slide, shape);
		const props = timing.getAnimProperties();
		props.asc_putSubtype(8);
		logicDocument.SetAnimationProperties(props, false);
		assert.ok(timing.getAllAnimEffects().every(function (effect) {return effect.cTn.presetSubtype === 8;}), "Subtype changed");
		assert.deepEqual(GetRanges(slide), [[0, 0], [2, 3], [4, 4]], "Ranges after subtype change");
	});

	QUnit.test("Binary round trip", function (assert)
	{
		const {slide, shape} = CreateTextShape(DEFAULT_TEXTS);
		AddFade(slide, shape);
		SelectShape(slide, shape);
		SetTextSequence(slide, BY_PARAGRAPH);

		const writer = new AscCommon.CBinaryFileWriter();
		writer.WriteRecord1(0, slide.timing, function (timing) {timing.toPPTY(writer);});
		const reader = new AscCommon.BinaryPPTYLoader();
		reader.stream = new AscCommon.FileStream(writer.data, writer.pos);
		reader.stream.cur = 1;
		const copy = AscFormat.ExecuteNoHistory(function ()
		{
			const timing = new AscFormat.CTiming();
			timing.fromPPTY(reader);
			return timing;
		}, this, []);

		const ranges = copy.getAllAnimEffects().map(function (effect)
		{
			let result = null;
			effect.traverse(function (child)
			{
				if (child.getObjectType() === AscDFH.historyitem_type_SpTgt && child.txEl && child.txEl.pRg)
				{
					result = [child.txEl.pRg.st, child.txEl.pRg.end];
					return true;
				}
				return false;
			});
			return result;
		});
		assert.deepEqual(ranges, [[0, 0], [2, 3], [4, 4]], "Paragraph ranges survive");
		const bldP = copy.bldLst && copy.bldLst.list[0];
		assert.ok(bldP && bldP.build === AscFormat.ParaBuildType_p && bldP.grpId === slide.timing.getAllAnimEffects()[0].cTn.grpId, "bldP survives");
	});

	QUnit.test("Imported builds are kept", function (assert)
	{
		const {slide, shape} = CreateTextShape(DEFAULT_TEXTS);
		const timing = AddFade(slide, shape);
		SelectShape(slide, shape);
		SetTextSequence(slide, BY_PARAGRAPH);
		const grpId = timing.getAllAnimEffects()[0].cTn.grpId;

		AscCommon.History.Create_NewPoint(0);
		timing.bldLst.list[0].setBuild(AscFormat.ParaBuildType_allAtOnce);
		const staleBldP = new AscFormat.CBldP();
		staleBldP.setSpid(shape.GetId());
		staleBldP.setGrpId(grpId + 10);
		timing.bldLst.push(staleBldP);
		timing.buildTree(timing.getEffectsSequences());
		assert.deepEqual(GetParagraphBuilds(slide, shape), [{isBldP : true, isShape : true, grpId : grpId, build : AscFormat.ParaBuildType_allAtOnce}], "allAtOnce is kept, the stale entry is removed");

		// character range build (e.g. from PowerPoint) without paragraph ranges
		SelectShape(slide, shape);
		SetTextSequence(slide, AS_ONE_OBJECT);
		const effect = timing.getAllAnimEffects()[0];
		AscCommon.History.Create_NewPoint(0);
		effect.traverse(function (child)
		{
			if (child.getObjectType() === AscDFH.historyitem_type_SpTgt)
			{
				const charRg = new AscFormat.CIndexRg();
				charRg.setSt(0);
				charRg.setEnd(3);
				const txEl = new AscFormat.CTxEl();
				txEl.setCharRg(charRg);
				child.setTxEl(txEl);
			}
			return false;
		});
		const custBldP = new AscFormat.CBldP();
		custBldP.setSpid(shape.GetId());
		custBldP.setGrpId(effect.cTn.grpId);
		custBldP.setBuild(AscFormat.ParaBuildType_cust);
		timing.setBldLst(new AscFormat.CBldLst());
		timing.bldLst.push(custBldP);
		timing.buildTree(timing.getEffectsSequences());
		assert.strictEqual(timing.bldLst && timing.bldLst.list.length, 1, "Character range build is kept");
	});

	QUnit.test("Undo and redo", function (assert)
	{
		const {slide, shape} = CreateTextShape(DEFAULT_TEXTS);
		AddFade(slide, shape);
		SelectShape(slide, shape);
		SetTextSequence(slide, BY_PARAGRAPH);

		Undo();
		assert.deepEqual(GetRanges(slide), [null], "Undo restores the whole object effect");
		assert.notOk(slide.timing.bldLst, "Undo removes the bldLst");
		logicDocument.Document_Redo();
		assert.deepEqual(GetRanges(slide), [[0, 0], [2, 3], [4, 4]], "Redo restores the paragraph effects");
		CheckParagraphBuild(assert, slide, shape, AscFormat.ParaBuildType_p, "Redo restores the bldP");
	});

	QUnit.test("Text changes keep the paragraph effects in sync", function (assert)
	{
		const {slide, shape} = CreateTextShape(DEFAULT_TEXTS);
		const timing = AddFade(slide, shape);
		SelectShape(slide, shape);
		SetTextSequence(slide, BY_PARAGRAPH);
		const docContent = shape.getDocContent();

		logicDocument.StartAction(0);
		docContent.Internal_Content_Add(docContent.Content.length, CreateParagraph(docContent, "Point D"));
		logicDocument.FinalizeAction();
		assert.deepEqual(GetRanges(slide), [[0, 0], [2, 3], [4, 4], [5, 5]], "New paragraph gets an effect");
		CheckParagraphBuild(assert, slide, shape, AscFormat.ParaBuildType_p, "Build stays p");
		Undo();
		assert.deepEqual([docContent.Content.length, GetRanges(slide)], [5, [[0, 0], [2, 3], [4, 4]]], "One undo restores text and effects");

		logicDocument.StartAction(0);
		const run = new AscWord.CRun(docContent.Content[1], false);
		run.AddText("x");
		docContent.Content[1].AddToContent(0, run);
		logicDocument.FinalizeAction();
		assert.deepEqual(GetRanges(slide), [[0, 0], [1, 1], [2, 3], [4, 4]], "Typing into the empty paragraph adds an effect");
		Undo();

		logicDocument.StartAction(0);
		docContent.Internal_Content_Remove(0, 1);
		logicDocument.FinalizeAction();
		assert.deepEqual(GetRanges(slide), [[1, 2], [3, 3]], "Removed paragraph drops its effect");
		Undo();
		assert.deepEqual(GetRanges(slide), [[0, 0], [2, 3], [4, 4]], "Undo removal");

		timing.resetSelection();
		timing.getAllAnimEffects()[1].select();
		AscCommon.History.Create_NewPoint(0);
		timing.removeSelectedEffects();
		assert.deepEqual(GetRanges(slide), [[0, 0], [4, 4]], "Single paragraph effect removed");
		CheckParagraphBuild(assert, slide, shape, AscFormat.ParaBuildType_cust, "Build becomes cust");
		logicDocument.StartAction(0);
		docContent.Internal_Content_Add(docContent.Content.length, CreateParagraph(docContent, "Point D"));
		logicDocument.FinalizeAction();
		assert.deepEqual(GetRanges(slide), [[0, 0], [4, 4]], "Custom builds are not synchronized");
	});

	QUnit.test("Text changes need the timing lock", function (assert)
	{
		const {slide, shape} = CreateTextShape(DEFAULT_TEXTS);
		AddFade(slide, shape);
		SelectShape(slide, shape);
		SetTextSequence(slide, BY_PARAGRAPH);
		const docContent = shape.getDocContent();
		const lock = slide.timingLock.Lock;

		lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeOther, false);
		logicDocument.StartAction(0);
		docContent.Internal_Content_Add(docContent.Content.length, CreateParagraph(docContent, "Point D"));
		logicDocument.FinalizeAction();
		assert.deepEqual([docContent.Content.length, GetRanges(slide)], [6, [[0, 0], [2, 3], [4, 4]]], "Animations locked by someone else: the text changes, the effects stay");

		lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeNone, false);
		logicDocument.StartAction(0);
		docContent.Internal_Content_Add(docContent.Content.length, CreateParagraph(docContent, "Point E"));
		logicDocument.FinalizeAction();
		assert.deepEqual(GetRanges(slide), [[0, 0], [2, 3], [4, 4], [5, 5], [6, 6]], "Lock free again: the next edit catches up");
		assert.strictEqual(lock.Get_Type(), AscCommon.c_oAscLockTypes.kLockTypeMine, "The timing lock is taken for the sync");
		CheckParagraphBuild(assert, slide, shape, AscFormat.ParaBuildType_p, "Build stays p");
	});

	QUnit.test("Removing text needs the timing lock", function (assert)
	{
		const {slide, shape} = CreateTextShape(DEFAULT_TEXTS);
		AddFade(slide, shape);
		SelectShape(slide, shape);
		SetTextSequence(slide, BY_PARAGRAPH);
		const docContent = shape.getDocContent();
		const lock = slide.timingLock.Lock;

		lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeOther, false);
		SelectShape(slide, shape);
		docContent.SetThisElementCurrent();
		logicDocument.MoveCursorToEndPos();
		const historyIndex = AscCommon.History.Index;
		for (let i = 0; i < 8; ++i)
			logicDocument.Remove(-1, true, false, false, false);
		let timingChanges = 0;
		for (let point = historyIndex + 1; point <= AscCommon.History.Index; ++point)
		{
			AscCommon.History.Points[point].Items.forEach(function (item)
			{
				if (item.Class instanceof AscFormat.CIndexRg || item.Class instanceof AscFormat.CTxEl)
					++timingChanges;
			});
		}
		assert.deepEqual([docContent.Content.length, GetRanges(slide), timingChanges], [4, [[0, 0], [2, 3], [4, 4]], 0], "Animations locked by someone else: the text targets are not clamped");

		lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeNone, false);
		logicDocument.EnterText("x".codePointsArray());
		assert.deepEqual(GetRanges(slide), [[0, 0], [2, 3]], "Lock free again: the next edit catches up");
		CheckParagraphBuild(assert, slide, shape, AscFormat.ParaBuildType_p, "Build stays p");
	});

	QUnit.test("By 2nd level paragraphs", function (assert)
	{
		const {slide, shape} = CreateTextShape(DEFAULT_TEXTS);
		const timing = AddFade(slide, shape);
		SelectShape(slide, shape);
		SetTextSequence(slide, BY_PARAGRAPH);
		const docContent = shape.getDocContent();
		assert.deepEqual(AscFormat.getParagraphBuildRanges(shape, 2), [{st : 0, end : 0}, {st : 2, end : 2}, {st : 3, end : 3}, {st : 4, end : 4}], "Level 2: the sub point is on its own");

		// build as imported from PowerPoint: build="p" bldLvl="2"
		AscCommon.History.Create_NewPoint(0);
		timing.bldLst.list[0].setBldLvl(2);
		const sequences = timing.getEffectsSequences();
		const effect = timing.getAllAnimEffects()[1];
		const copy = effect.createDuplicate();
		copy.setParagraphRange(3, 3);
		effect.setParagraphRange(2, 2);
		sequences[0].splice(sequences[0].indexOf(effect) + 1, 0, copy);
		timing.buildTree(sequences);
		assert.deepEqual(GetRanges(slide), [[0, 0], [2, 2], [3, 3], [4, 4]], "One effect per 2nd level paragraph");
		CheckParagraphBuild(assert, slide, shape, AscFormat.ParaBuildType_p, "Build stays p");

		logicDocument.StartAction(0);
		const subPoint = CreateParagraph(docContent, "Sub point C1");
		docContent.Internal_Content_Add(docContent.Content.length, subPoint);
		subPoint.Set_PresentationLevel(1);
		const subSubPoint = CreateParagraph(docContent, "Sub sub point C1a");
		docContent.Internal_Content_Add(docContent.Content.length, subSubPoint);
		subSubPoint.Set_PresentationLevel(2);
		logicDocument.FinalizeAction();
		assert.deepEqual(GetRanges(slide), [[0, 0], [2, 2], [3, 3], [4, 4], [5, 6]], "Text edit keeps level 2: 3rd level joins its parent");
		CheckParagraphBuild(assert, slide, shape, AscFormat.ParaBuildType_p, "Build stays p after the text edit");
		assert.strictEqual(timing.bldLst.list[0].bldLvl, 2, "bldLvl kept");
		Undo();
		assert.deepEqual(GetRanges(slide), [[0, 0], [2, 2], [3, 3], [4, 4]], "One undo restores text and effects");
	});

	QUnit.test("Editing text in the shape", function (assert)
	{
		const {slide, shape} = CreateTextShape(DEFAULT_TEXTS);
		AddFade(slide, shape);
		SelectShape(slide, shape);
		SetTextSequence(slide, BY_PARAGRAPH);
		const docContent = shape.getDocContent();

		SelectShape(slide, shape);
		docContent.SetThisElementCurrent();
		logicDocument.MoveCursorToEndPos();
		logicDocument.AddNewParagraph();
		assert.deepEqual([docContent.Content.length, GetRanges(slide)], [6, [[0, 0], [2, 3], [4, 4]]], "Enter: an empty paragraph has no effect");
		logicDocument.EnterText("Point D".codePointsArray());
		assert.deepEqual(GetRanges(slide), [[0, 0], [2, 3], [4, 4], [5, 5]], "Typing adds an effect");
		for (let i = 0; i < 7; ++i)
			logicDocument.Remove(-1, true, false, false, false);
		assert.deepEqual(GetRanges(slide), [[0, 0], [2, 3], [4, 4]], "Backspace to an empty paragraph removes the effect");
		logicDocument.Remove(-1, true, false, false, false);
		assert.deepEqual([docContent.Content.length, GetRanges(slide)], [5, [[0, 0], [2, 3], [4, 4]]], "Backspace removes the paragraph");
	});

	QUnit.test("Shapes without text and mixed selection", function (assert)
	{
		const {slide, shape} = CreateTextShape(DEFAULT_TEXTS);
		const timing = AddFade(slide, shape);
		SelectShape(slide, shape);
		SetTextSequence(slide, BY_PARAGRAPH);

		const apiEllipse = builderApi.CreateShape("ellipse", 100 * 36000, 100 * 36000, builderApi.CreateSolidFill(builderApi.CreateRGBColor(10, 10, 10)), builderApi.CreateStroke(0, builderApi.CreateNoFill()));
		builderApi.GetPresentation().GetSlideByIndex(logicDocument.GetSlidesCount() - 1).AddObject(apiEllipse);
		const ellipse = apiEllipse.Shape;
		AddFade(slide, ellipse);
		assert.strictEqual(timing.getAnimProperties().asc_getTextSequence(), null, "No text: not applicable");

		slide.graphicObjects.selectObject(shape, 0);
		timing.onChangeDrawingsSelection();
		assert.strictEqual(timing.getAnimProperties().asc_getTextSequence(), undefined, "Mixed selection");
		SetTextSequence(slide, AS_ONE_OBJECT);
		assert.deepEqual(GetRanges(slide), [null, null], "Only the text shape collapses");
		assert.notOk(timing.bldLst, "bldLst removed");

		SelectShape(slide, shape);
		SetTextSequence(slide, BY_PARAGRAPH);
		AscCommon.History.Create_NewPoint(0);
		slide.removeFromSpTreeById(shape.GetId());
		assert.ok(timing.getAllAnimEffects().every(function (effect) {return effect.getObjectId() !== shape.GetId();}), "Effects of the removed shape are gone");
		assert.notOk(timing.bldLst, "bldP of the removed shape is gone");
	});

	QUnit.test("Playback shows one paragraph per click", function (assert)
	{
		const done = assert.async();
		const {slide, shape} = CreateTextShape(DEFAULT_TEXTS);
		AddFade(slide, shape);
		SelectShape(slide, shape);
		SetTextSequence(slide, BY_PARAGRAPH);
		logicDocument.Recalculate();

		const player = new AscFormat.CAnimationPlayer(slide);
		const drawer = player.animationDrawer;

		function GetVisibleParagraphs()
		{
			const docStructure = drawer.getDrawingTextCache(shape.GetId());
			const result = [];
			for (let i = 0; i < docStructure.m_aContent.length; ++i)
			{
				const ids = [];
				(docStructure.m_aContent[i].m_aWrapperElementsCache || []).forEach(function (element)
				{
					element.forEachAnimationDrawing(function (drawing) {ids.push(drawing.GetId());});
				});
				if (shape.getDocContent().Content[i].IsEmpty() || !ids.length)
					continue;
				if (ids.every(function (id) {return drawer.isDrawingVisible(id);}))
					result.push(i);
			}
			return result;
		}

		function Wait(ms)
		{
			return new Promise(function (resolve) {setTimeout(resolve, ms);});
		}

		(async function ()
		{
			player.start();
			await Wait(300);
			assert.deepEqual(GetVisibleParagraphs(), [], "Start: no paragraph visible");
			assert.ok(drawer.isDrawingVisible(shape.GetId()), "Start: the shape itself is visible");
			player.onClick();
			await Wait(1200);
			assert.deepEqual(GetVisibleParagraphs(), [0], "Click 1");
			player.onClick();
			await Wait(1200);
			assert.deepEqual(GetVisibleParagraphs(), [0, 2, 3], "Click 2: paragraph with its sub point");
			player.onClick();
			await Wait(1200);
			assert.deepEqual(GetVisibleParagraphs(), [0, 2, 3, 4], "Click 3");
			player.stop();
			done();
		})();
	});
});
