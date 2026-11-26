/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	QUnit.module("Thumbnails of slides with remote changes");

	// A thumbnails manager with eight slides in view, each with a thumbnail of the right size. It counts the
	// thumbnails that it draws (DrawPage) and the paints of the panel (OnPaint). Its clock is set by the test.
	function createThumbnails()
	{
		let thumbnails = new CThumbnailsManager();
		thumbnails.isThumbnailsShown = function () { return true; };
		thumbnails.DigitWidths = [1];
		thumbnails.drawn = [];
		thumbnails.paints = 0;
		thumbnails.OnPaint = function () { this.paints++; };
		thumbnails.now = 10000;
		thumbnails.GetTime = function () { return this.now; };
		thumbnails.m_oCacheManager = {
			Lock   : function (w, h) { return {image : {width : w, height : h, ctx : {}}}; },
			UnLock : function () {}
		};
		thumbnails.slideIds = ["s0", "s1", "s2", "s3", "s4", "s5", "s6", "s7"];
		thumbnails.m_oWordControl = {
			m_oApi             : {isSaveFonts_Images : false, ShowParaMarks : false},
			m_oDrawingDocument : {IsEmptyPresentation : false},
			m_oLogicDocument   : {
				IsVisioEditor : function () { return false; },
				DrawPage      : function (index) { thumbnails.drawn.push(index); },
				GetSlide      : function (index)
				{
					return {Get_Id : function () { return thumbnails.slideIds[index]; }};
				}
			}
		};
		thumbnails.m_arrPages = [];
		for (let i = 0; i < 8; ++i)
		{
			let page = new CThPage();
			page.left = 0;
			page.top = 0;
			page.right = 100;
			page.bottom = 60;
			page.cachedImage = thumbnails.m_oCacheManager.Lock(100, 60);
			thumbnails.m_arrPages.push(page);
		}
		thumbnails.m_lDrawingFirst = 0;
		thumbnails.m_lDrawingEnd = 7;
		return thumbnails;
	}

	// The global objects of the editor that the code reads, for the time of fn.
	function withEditor(fn)
	{
		let api = Asc.editor;
		let globalEditor = window.editor;
		let collaborativeEditing = AscCommon.CollaborativeEditing;
		let graphics = AscCommon.CGraphics;
		Asc.editor = {
			isSlideShow    : function () { return false; },
			sync_EndAction : function () {}
		};
		window.editor = {sendEvent : function () {}};
		AscCommon.CollaborativeEditing = {Set_GlobalLock : function () {}, Set_GlobalLockSelection : function () {}};
		AscCommon.CGraphics = function () {};
		AscCommon.CGraphics.prototype.init = function () {};
		AscCommon.CGraphics.prototype.transform = function () {};
		try { fn(); } finally {
			Asc.editor = api;
			window.editor = globalEditor;
			AscCommon.CollaborativeEditing = collaborativeEditing;
			AscCommon.CGraphics = graphics;
		}
	}

	// The drawing document of the thumbnails, made while withEditor runs.
	function createDrawingDocument(thumbnails)
	{
		let drawingDocument = new CDrawingDocument();
		drawingDocument.m_oWordControl = {Thumbnails : thumbnails};
		drawingDocument.SlideCurrent = -1;
		return drawingDocument;
	}

	// Remote changes are recalculated through the code of the co-editing (CCollaborativeEditing.OnEnd_Load_Objects),
	// and the recalculation marks each slide (OnRecalculateSlide). Returns the presentation.
	function recalculateRemote(drawingDocument, indexes)
	{
		let presentation = {
			DrawingDocument               : drawingDocument,
			recalculated                  : false,
			History                       : {Get_RecalcData : function () { return {}; }},
			Recalculate                   : function ()
			{
				this.recalculated = true;
				indexes.forEach(function (index) { drawingDocument.OnRecalculateSlide(index); });
			},
			Document_UpdateSelectionState : function () {},
			Document_UpdateInterfaceState : function () {}
		};
		let coEditing = Object.create(AscCommon.SlideCollaborativeEditing.prototype);
		coEditing.GetPresentation = function () { return presentation; };
		coEditing.CoHistory = {GetAllChanges : function () { return []; }};
		coEditing.m_nRecalcIndexStart = 0;
		coEditing.m_nRecalcIndexEnd = 0;
		coEditing.OnEnd_Load_Objects();
		return presentation;
	}

	function changeByOthers(thumbnails, indexes)
	{
		recalculateRemote(createDrawingDocument(thumbnails), indexes);
	}

	// An own change of the slide.
	function changeOwn(thumbnails, index)
	{
		createDrawingDocument(thumbnails).OnRecalculateSlide(index);
	}

	function checkTimes(thumbnails, count)
	{
		for (let i = 0; i < count; ++i)
			thumbnails.onCheckUpdate();
	}

	// Others changed slide 0, and its thumbnail was drawn. Then they change it again: now it waits.
	function waitingSlide()
	{
		let thumbnails = createThumbnails();
		changeByOthers(thumbnails, [0]);
		thumbnails.onCheckUpdate();
		changeByOthers(thumbnails, [0]);
		return thumbnails;
	}

	QUnit.test("The own changes are drawn at once", function (assert)
	{
		withEditor(function ()
		{
			let thumbnails = createThumbnails();
			for (let i = 0; i < 8; ++i)
				changeOwn(thumbnails, i);

			checkTimes(thumbnails, 8);

			assert.deepEqual(thumbnails.drawn, [0, 1, 2, 3, 4, 5, 6, 7], "one thumbnail on each check, all eight");
			assert.strictEqual(thumbnails.paints, 8, "the panel is painted on each check");
		});
	});

	QUnit.test("Slides that others changed do not wait for each other", function (assert)
	{
		withEditor(function ()
		{
			let thumbnails = createThumbnails();
			changeByOthers(thumbnails, [0, 1, 2, 3, 4, 5, 6, 7]);

			checkTimes(thumbnails, 8);

			assert.deepEqual(thumbnails.drawn, [0, 1, 2, 3, 4, 5, 6, 7], "a change of others on eight slides: all are drawn");
		});
	});

	QUnit.test("A slide that others change again is drawn at most once in the interval", function (assert)
	{
		withEditor(function ()
		{
			let thumbnails = waitingSlide();
			let paints = thumbnails.paints;

			thumbnails.now += THUMBNAIL_REMOTE_REDRAW_INTERVAL - 1;
			checkTimes(thumbnails, 3);
			assert.deepEqual(thumbnails.drawn, [0], "the second change waits");
			assert.strictEqual(thumbnails.m_arrPages[0].IsRecalc, true, "the slide stays to be drawn");
			assert.strictEqual(thumbnails.paints, paints, "no paints while only the waiting slide is left");

			// Positive control: at the end of the interval the waiting slide is drawn.
			thumbnails.now += 1;
			thumbnails.onCheckUpdate();
			assert.deepEqual(thumbnails.drawn, [0, 0], "the slide is drawn after the interval");
		});
	});

	QUnit.test("The flag of the changes of others is off after their recalculation", function (assert)
	{
		withEditor(function ()
		{
			let thumbnails = createThumbnails();
			changeByOthers(thumbnails, [2]);
			assert.strictEqual(thumbnails.m_arrPages[2].IsRemoteRecalc, true, "the change of others is marked");

			changeOwn(thumbnails, 3);
			assert.strictEqual(thumbnails.m_arrPages[3].IsRemoteRecalc, false, "a later own change is not");
		});
	});

	QUnit.test("Remote changes are recalculated with a drawing document without SetRemoteRecalc", function (assert)
	{
		withEditor(function ()
		{
			// The drawing documents of other builds (mobile, native) can lack the method.
			let presentation = recalculateRemote({OnRecalculateSlide : function () {}}, [0]);
			assert.strictEqual(presentation.recalculated, true, "the remote changes are recalculated");
		});
	});

	QUnit.test("The panel is painted at once while a slide of others waits", function (assert)
	{
		withEditor(function ()
		{
			let thumbnails = waitingSlide();
			let paints = thumbnails.paints;

			// A scroll, a selection, a move or a delete asks for an update of the panel.
			thumbnails.m_bIsUpdate = true;
			thumbnails.onCheckUpdate();

			assert.strictEqual(thumbnails.paints, paints + 1, "the panel is painted");
			assert.deepEqual(thumbnails.drawn, [0], "the waiting slide is not drawn");
		});
	});

	QUnit.test("An own change is drawn at once while a slide of others waits", function (assert)
	{
		withEditor(function ()
		{
			let thumbnails = waitingSlide();

			changeOwn(thumbnails, 5);
			thumbnails.onCheckUpdate();

			assert.deepEqual(thumbnails.drawn, [0, 5], "the own change is drawn, the waiting slide is not");
		});
	});

	QUnit.test("An own change of the waiting slide is drawn at once", function (assert)
	{
		withEditor(function ()
		{
			let thumbnails = waitingSlide();

			changeOwn(thumbnails, 0);
			thumbnails.onCheckUpdate();

			assert.deepEqual(thumbnails.drawn, [0, 0], "the slide is drawn");
		});
	});

	QUnit.test("An own change that is not drawn yet stays own when others change the slide", function (assert)
	{
		withEditor(function ()
		{
			let thumbnails = createThumbnails();
			changeByOthers(thumbnails, [0]);
			thumbnails.onCheckUpdate();

			changeOwn(thumbnails, 0);
			changeByOthers(thumbnails, [0]);
			thumbnails.onCheckUpdate();

			assert.deepEqual(thumbnails.drawn, [0, 0], "the own change is drawn at once");
		});
	});

	QUnit.test("A thumbnail that now shows another slide does not wait", function (assert)
	{
		withEditor(function ()
		{
			let thumbnails = waitingSlide();

			// Others deleted the slide before it: the slide of index 0 is another one now.
			thumbnails.slideIds.shift();
			thumbnails.onCheckUpdate();

			assert.deepEqual(thumbnails.drawn, [0, 0], "the new slide of index 0 is drawn");
		});
	});

	QUnit.test("A thumbnail drawn for another slide in between does not wait", function (assert)
	{
		withEditor(function ()
		{
			let thumbnails = createThumbnails();
			changeByOthers(thumbnails, [0]);
			thumbnails.onCheckUpdate();

			// An own delete of a slide before it: slide s1 comes to index 0 and is drawn there.
			thumbnails.slideIds.shift();
			changeOwn(thumbnails, 0);
			thumbnails.onCheckUpdate();

			// Others insert a slide before it: slide s0 is back at index 0.
			thumbnails.slideIds.unshift("s0");
			changeByOthers(thumbnails, [0]);
			thumbnails.onCheckUpdate();

			assert.deepEqual(thumbnails.drawn, [0, 0, 0], "the thumbnail shows s1, so s0 is drawn at once");
		});
	});

	QUnit.test("A thumbnail of another size does not wait", function (assert)
	{
		withEditor(function ()
		{
			let thumbnails = waitingSlide();

			// The panel was resized.
			thumbnails.m_arrPages.forEach(function (page) { page.right = 120; });
			thumbnails.onCheckUpdate();

			assert.deepEqual(thumbnails.drawn, [0, 0], "the slide is drawn in the new size");
		});
	});

	QUnit.test("All the thumbnails are drawn at once after a full redraw", function (assert)
	{
		withEditor(function ()
		{
			let thumbnails = waitingSlide();

			// A theme change, a resize or an undo draws all the thumbnails again.
			thumbnails.ClearCacheAttack();
			checkTimes(thumbnails, 8);

			assert.deepEqual(thumbnails.drawn, [0, 0, 1, 2, 3, 4, 5, 6, 7], "the waiting slide does not wait");
		});
	});

	QUnit.test("A new thumbnail of a slide that others changed is drawn at once", function (assert)
	{
		withEditor(function ()
		{
			let thumbnails = waitingSlide();
			thumbnails.m_arrPages[0].cachedImage = null;

			thumbnails.onCheckUpdate();

			assert.deepEqual(thumbnails.drawn, [0, 0], "a slide without a thumbnail does not wait");
		});
	});
});
