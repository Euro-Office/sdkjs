/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	AscCommon.CViewer.prototype.createComponents = function(){};
	AscCommon.CViewer.prototype.resize = function(){};
	// no fonts are needed for scrolling (the test page has no font files)
	AscPDF.CPDFDoc.prototype.checkDefaultFonts = function(){};

	// three pages of 500 px between gaps of 20 px; the window is 600 px high (a page fits)
	function createViewer(pageHeight, windowHeight)
	{
		let viewer = new AscCommon.CViewer(null, AscTest.Editor);
		viewer.betweenPages = 20;
		viewer.height = windowHeight;
		viewer.scrollY = 0;
		viewer.drawingPages = [];
		let top = 20;
		for (let i = 0; i < 3; i++)
		{
			viewer.drawingPages.push({X : 0, Y : top, W : 400, H : pageHeight});
			top += pageHeight + 20;
		}
		let maxY = top - windowHeight;
		viewer.m_oScrollVerApi = {
			scrollByY : function (d) { viewer.scrollY = Math.max(0, Math.min(maxY, viewer.scrollY + d)); },
			scrollToY : function (y) { viewer.scrollY = Math.max(0, Math.min(maxY, y)); }
		};
		viewer.navigateToPage = function (n) { viewer.m_oScrollVerApi.scrollToY(viewer.drawingPages[n].Y - viewer.betweenPages); };
		viewer.timerSync = function () {};
		return viewer;
	}

	QUnit.module("PDF page by page");

	QUnit.test("Pages that fit into the window: one step is one page", function (assert)
	{
		let viewer = createViewer(500, 600);
		viewer.scrollPageByPage(100, true);
		assert.strictEqual(viewer.getPageByPageIndex(), 1, "a step down shows the second page");
		assert.strictEqual(viewer.scrollY, 520, "at its top");
		viewer.scrollPageByPage(100, true);
		assert.strictEqual(viewer.scrollY, 980, "the next step shows the last page (at the end of the document)");
		viewer.scrollPageByPage(100, true);
		assert.strictEqual(viewer.scrollY, 980, "there is no page after the last one");
		viewer.scrollPageByPage(-100, true);
		assert.strictEqual(viewer.scrollY, 520, "a step up shows the second page again");
		viewer.scrollPageByPage(-100, true);
		assert.strictEqual(viewer.scrollY, 0, "and the next one the first page");
	});

	QUnit.test("Pages higher than the window (fit to width): scrolled through first, then the next page", function (assert)
	{
		let viewer = createViewer(1000, 600);
		viewer.scrollPageByPage(300, true);
		assert.strictEqual(viewer.scrollY, 300, "inside the first page it scrolls normally");
		viewer.scrollPageByPage(300, true);
		assert.strictEqual(viewer.scrollY, 440, "but not over its end (page end + gap = window bottom)");
		viewer.scrollPageByPage(300, true);
		assert.strictEqual(viewer.scrollY, 1020, "then it jumps to the top of the second page");
		viewer.scrollPageByPage(-300, true);
		assert.strictEqual(viewer.scrollY, 440, "going up shows the end of the first page");
		viewer.scrollPageByPage(-300, true);
		assert.strictEqual(viewer.scrollY, 140, "and then scrolls up in it");
	});

	QUnit.test("A wheel turn of many small events moves only one page", function (assert)
	{
		let viewer = createViewer(500, 600);
		for (let i = 0; i < 10; i++)
			viewer.scrollPageByPage(30, false);
		assert.strictEqual(viewer.getPageByPageIndex(), 1, "ten quick events from a touchpad give one page");
		viewer.pageByPageTime = 0;
		viewer.scrollPageByPage(30, false);
		assert.strictEqual(viewer.scrollY, 980, "after a short pause the next turn goes on to the last page");
	});

	QUnit.test("A long touchpad burst (inertia) still moves only one page", function (assert)
	{
		let viewer = createViewer(500, 600);
		let realNow = Date.now;
		let time = 100000;
		Date.now = function () { return time; };
		try
		{
			// events every 60 ms for 720 ms: longer than the pause, but there is never a pause between them
			for (let i = 0; i <= 12; i++, time += 60)
				viewer.scrollPageByPage(30, false);
			assert.strictEqual(viewer.scrollY, 520, "the whole burst gives one page: the second page, at its top");
			time += 300;
			viewer.scrollPageByPage(30, false);
			assert.strictEqual(viewer.scrollY, 980, "after a pause the next turn goes on to the last page");
			time += 60;
			viewer.scrollPageByPage(-30, true);
			assert.strictEqual(viewer.scrollY, 520, "keys are never held back");
		}
		finally
		{
			Date.now = realNow;
		}
	});

	QUnit.test("A touchpad swipe on pages higher than the window: to the end of the page, then the next page", function (assert)
	{
		let viewer = createViewer(1000, 600);
		let realNow = Date.now;
		let time = 100000;
		Date.now = function () { return time; };
		function swipe(delta)
		{
			// 60 events, 16 ms apart, then a pause
			for (let i = 0; i < 60; i++, time += 16)
				viewer.scrollPageByPage(delta, false);
			time += 300;
		}
		try
		{
			swipe(30);
			assert.strictEqual(viewer.scrollY, 440, "the first swipe scrolls to the end of the first page and stops there");
			swipe(30);
			assert.strictEqual(viewer.scrollY, 1020, "the second swipe goes to the top of the second page");
			swipe(30);
			assert.strictEqual(viewer.scrollY, 1460, "the third swipe scrolls to the end of the second page");
			swipe(-30);
			assert.strictEqual(viewer.scrollY, 1020, "a swipe up scrolls to the top of the second page and stops there");
			swipe(-30);
			assert.strictEqual(viewer.scrollY, 440, "the next one shows the end of the first page");
		}
		finally
		{
			Date.now = realNow;
		}
	});

	QUnit.test("A page turned by 90 or 270 degrees is as high as it is wide (as in the page layout)", function (assert)
	{
		let viewer = createViewer(500, 600);
		let pages = [{Rotate : 0}, {Rotate : 90}, {Rotate : 180}];
		viewer.file = {isValid : function () { return true; }, pages : pages};
		assert.strictEqual(viewer.getPageByPageHeight(0), 500, "0 degrees: the height");
		assert.strictEqual(viewer.getPageByPageHeight(1), 400, "90 degrees: the width");
		assert.strictEqual(viewer.getPageByPageHeight(2), 500, "180 degrees: the height");
		pages[2].Rotate = 270;
		assert.strictEqual(viewer.getPageByPageHeight(2), 400, "270 degrees: the width");
	});

	QUnit.test("The page at the top of the window, in a document with 5000 pages", function (assert)
	{
		let viewer = createViewer(500, 600);
		viewer.drawingPages = [];
		let top = 20;
		for (let i = 0; i < 5000; i++)
		{
			// different heights, so that a wrong guess cannot hit the right page by chance
			let height = 300 + (i * 37) % 700;
			viewer.drawingPages.push({X : 0, Y : top, W : 400, H : height});
			top += height + 20;
		}

		// the way it was found before: look at every page from the first one
		function linearIndex()
		{
			for (let i = 0; i < viewer.drawingPages.length; i++)
			{
				let page = viewer.drawingPages[i];
				if (page.Y + page.H > viewer.scrollY)
					return i;
			}
			return viewer.drawingPages.length - 1;
		}

		let positions = [0, 1, 19, 20, 21, top + 1000];
		[1, 2, 777, 2500, 4998, 4999].forEach(function (n)
		{
			let page = viewer.drawingPages[n];
			// the page top, its last pixel, its bottom, the gap below it
			positions.push(page.Y, page.Y + page.H - 1, page.Y + page.H, page.Y + page.H + 10);
		});
		for (let y = 0; y < top; y += 9973)
			positions.push(y);

		let isSame = true;
		positions.forEach(function (y)
		{
			viewer.scrollY = y;
			let index = viewer.getPageByPageIndex();
			if (index !== linearIndex())
			{
				isSame = false;
				assert.strictEqual(index, linearIndex(), "scroll position " + y);
			}
		});
		assert.ok(isSame, "the same page as before at all " + positions.length + " scroll positions");

		viewer.scrollY = viewer.drawingPages[4999].Y - 25;
		assert.strictEqual(viewer.getPageByPageIndex(), 4998, "in the gap above the last page: the page before it");
		viewer.scrollY = top + 1000;
		assert.strictEqual(viewer.getPageByPageIndex(), 4999, "below all pages: the last page");
		viewer.drawingPages = [];
		assert.strictEqual(viewer.getPageByPageIndex(), -1, "no pages: -1");
	});

	QUnit.test("The mouse wheel: page by page only for a plain vertical wheel outside a form field", function (assert)
	{
		let viewer = createViewer(500, 600);
		viewer.file = {isValid : function () { return true; }, pages : [{}, {}, {}]};
		viewer.isVisibleHorScroll = true;
		viewer.isPageByPage = true;
		viewer.onMouseMove = function () {};

		let calls;
		viewer.scrollPageByPage = function (delta, isKey) { calls.push("page " + delta + " " + isKey); return true; };
		viewer.m_oScrollVerApi.scrollBy = function (x, y) { calls.push("vertical " + y); };
		viewer.m_oScrollHorApi = {scrollBy : function (x) { calls.push("horizontal " + x); }};
		let field = null;
		viewer.getPageFieldByMouse = function () { return field; };

		function wheel(e)
		{
			calls = [];
			e.preventDefault = function () {};
			e.stopPropagation = function () {};
			viewer.onMouseWhell(e);
			return calls.join(", ");
		}

		assert.strictEqual(wheel({wheelDelta : -120, wheelDeltaY : -120, wheelDeltaX : 0}), "page 45 false",
			"a wheel step down goes page by page");
		assert.strictEqual(wheel({wheelDelta : -120, wheelDeltaY : -120, wheelDeltaX : 0, ctrlKey : true}), "",
			"ctrl + wheel is left alone (it zooms), as before");
		assert.strictEqual(wheel({wheelDelta : -120, wheelDeltaY : 0, wheelDeltaX : -120}), "horizontal 45",
			"a horizontal wheel scrolls sideways, as before");

		let fieldScroll = [];
		field = {
			GetType : function () { return AscPDF.FIELD_TYPES.text; },
			GetPage : function () { return 0; },
			GetScrollInfo : function ()
			{
				return {
					docElem : {style : {display : "block"}},
					scroll : {
						settings : {vscrollStep : 10, hscrollStep : 10},
						scrollByY : function (d) { fieldScroll.push(d); },
						scrollByX : function () {}
					}
				};
			}
		};
		assert.strictEqual(wheel({wheelDelta : -120, wheelDeltaY : -120, wheelDeltaX : 0, deltaY : 100}), "",
			"over a text field with a scroll bar the page does not move");
		assert.deepEqual(fieldScroll, [10], "the field scrolls instead, as before");

		field = null;
		viewer.isPageByPage = false;
		assert.strictEqual(wheel({wheelDelta : -120, wheelDeltaY : -120, wheelDeltaX : 0}), "vertical 45",
			"switched off: the wheel scrolls as before");
	});

	QUnit.test("Switching it on lines up the current page", function (assert)
	{
		let viewer = createViewer(500, 600);
		viewer.file = {isValid : function () { return true; }, pages : [{}, {}, {}]};
		viewer.scrollY = 700;
		viewer.setPageByPage(true);
		assert.ok(viewer.isPageByPage, "it is on");
		assert.strictEqual(viewer.scrollY, 520, "the page at the top of the window is shown from its top");
	});

	QUnit.test("Page Down and Page Up on pages higher than the window: through the page first, then the next page", function (assert)
	{
		// 100 % zoom in a small window: an A4 page (1123 px) is higher than the window (650 px)
		let viewer = createViewer(1123, 650);
		viewer.isPageByPage = true;
		viewer.doc = {
			GetController : function () { return {}; },
			GetDrawingDocument : function () { return {}; },
			executeShortcut : function () { return false; },
			UpdateCopyCutState : function () {}
		};
		function key(code)
		{
			viewer.OnKeyDown({KeyCode : code, CtrlKey : false, ShiftKey : false, AltKey : false, MacCmdKey : false});
		}
		key(34);
		assert.strictEqual(viewer.scrollY, 513, "the first Page Down shows the lower part of the first page");
		key(34);
		assert.strictEqual(viewer.scrollY, 1143, "the next one goes to the top of the second page");
		key(33);
		assert.strictEqual(viewer.scrollY, 513, "Page Up shows the lower part of the first page again");
		key(33);
		assert.strictEqual(viewer.scrollY, 0, "and then its top");
	});
});
