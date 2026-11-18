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

	QUnit.test("Switching it on lines up the current page", function (assert)
	{
		let viewer = createViewer(500, 600);
		viewer.file = {isValid : function () { return true; }, pages : [{}, {}, {}]};
		viewer.scrollY = 700;
		viewer.setPageByPage(true);
		assert.ok(viewer.isPageByPage, "it is on");
		assert.strictEqual(viewer.scrollY, 520, "the page at the top of the window is shown from its top");
	});
});
