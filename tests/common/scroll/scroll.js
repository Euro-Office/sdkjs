/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

QUnit.module("ScrollObject");

(function ()
{
	let holderCount = 0;

	function createHolder(width, height)
	{
		let holder = document.createElement("div");
		holder.id = "test_scroll_" + (++holderCount);
		holder.style.position = "absolute";
		holder.style.width = width + "px";
		holder.style.height = height + "px";
		// the fixture is hidden off screen by qunit.css, element sizes still apply
		document.getElementById("qunit-fixture").appendChild(holder);
		return holder;
	}

	function createVertical(contentH, screenH)
	{
		let holder = createHolder(14, 200);
		let settings = new AscCommon.ScrollSettings();
		settings.contentH = contentH;
		settings.screenH = screenH;
		let scroll = new AscCommon.ScrollObject(holder.id, settings);
		let events = [];
		scroll.bind("scrollvertical", function (evt)
		{
			events.push({pos : evt.scrollD, max : evt.maxScrollY});
		});
		return {holder : holder, scroll : scroll, events : events};
	}

	function nextFrames(count)
	{
		return new Promise(function (resolve)
		{
			(function step()
			{
				if (0 === count--)
					resolve();
				else
					requestAnimationFrame(step);
			})();
		});
	}

	QUnit.test("native scroll range matches the scroll range", function (assert)
	{
		let t = createVertical(1000, 200);
		assert.strictEqual(t.scroll.getMaxScrolledY(), 800, "max scroll");
		assert.strictEqual(t.holder.scrollHeight - t.holder.clientHeight, 800, "native max scroll");

		t.scroll.Repos(Object.assign(new AscCommon.ScrollSettings(), {contentH : 500, screenH : 200}), undefined, true);
		assert.strictEqual(t.scroll.getMaxScrolledY(), 300, "max scroll after Repos");
		assert.strictEqual(t.holder.scrollHeight - t.holder.clientHeight, 300, "native max scroll after Repos");
	});

	QUnit.test("scrollToY moves the native scrollbar without echo", async function (assert)
	{
		let t = createVertical(1000, 200);
		t.scroll.scrollToY(300);
		assert.deepEqual(t.events, [{pos : 300, max : 800}], "one event from scrollToY");

		await nextFrames(3);
		assert.strictEqual(t.holder.scrollTop, 300, "native position follows");
		assert.strictEqual(t.events.length, 1, "the native scroll event does not fire scrollvertical again");

		t.scroll.scrollByY(1000);
		await nextFrames(3);
		assert.strictEqual(t.scroll.getCurScrolledY(), 800, "scrollByY stops at the end");
		assert.strictEqual(t.holder.scrollTop, 800, "native position at the end");
		assert.strictEqual(t.events.length, 2, "no echo after scrollByY");
	});

	QUnit.test("native scroll fires scrollvertical once per frame", async function (assert)
	{
		let t = createVertical(1000, 200);
		t.holder.scrollTop = 400;
		t.holder.scrollTop = 450;
		await nextFrames(3);
		assert.deepEqual(t.events, [{pos : 450, max : 800}], "one event with the last position");
		assert.strictEqual(t.scroll.getCurScrolledY(), 450, "position");

		t.holder.scrollTop = 800;
		await nextFrames(3);
		assert.strictEqual(t.scroll.getCurScrolledY(), 800, "native end is the scroll end");
		assert.strictEqual(t.events.length, 2, "no loop");
	});

	QUnit.test("scale above the native size cap round-trips", async function (assert)
	{
		// a sheet with 1,048,576 rows of 20 px, more than browsers allow for one element
		let t = createVertical(1048576 * 20, 0);
		let max = t.scroll.getMaxScrolledY();
		let nativeMax = t.holder.scrollHeight - t.holder.clientHeight;
		assert.strictEqual(max, 20971520, "max scroll");
		assert.strictEqual(nativeMax, 10000000, "native range is capped");

		t.scroll.scrollToY(max / 2);
		await nextFrames(3);
		assert.ok(Math.abs(t.holder.scrollTop - nativeMax / 2) <= 1, "native position is scaled: " + t.holder.scrollTop);
		assert.strictEqual(t.events.length, 1, "no echo");

		t.holder.scrollTop = nativeMax / 4;
		await nextFrames(3);
		let pos = t.scroll.getCurScrolledY();
		assert.ok(Math.abs(pos - max / 4) <= max / nativeMax, "position is scaled back: " + pos);
		assert.strictEqual(t.events.length, 2, "one event");

		t.holder.scrollTop = nativeMax;
		await nextFrames(3);
		assert.strictEqual(t.scroll.getCurScrolledY(), max, "native end is the scroll end");
	});

	QUnit.test("drag corrects the position and ends on mouseup", async function (assert)
	{
		let t = createVertical(1000, 200);
		let corrected = [];
		let mouseUps = [];
		t.scroll.bind("correctVerticalScroll", function (pos)
		{
			corrected.push(pos);
			return {isChange : true, Pos : pos + 1};
		});
		t.scroll.bind("mouseup", function ()
		{
			mouseUps.push(t.scroll.getCurScrolledY());
			// snap, as presentations do to the start of a slide
			t.scroll.scrollToY(250);
		});

		t.holder.scrollTop = 100;
		await nextFrames(3);
		assert.deepEqual(corrected, [], "no correction without a drag");

		t.holder.dispatchEvent(new MouseEvent("mousedown", {button : 0, bubbles : true}));
		assert.ok(t.scroll.getIsLockedMouse(), "mouse is down");
		t.holder.scrollTop = 200;
		await nextFrames(3);
		assert.deepEqual(corrected, [200], "drag position is corrected");
		assert.strictEqual(t.scroll.getCurScrolledY(), 201, "corrected position");
		assert.strictEqual(t.holder.scrollTop, 200, "native position stays under the mouse");

		// Chrome applies the last drag step after mouseup
		window.dispatchEvent(new MouseEvent("mouseup", {button : 0}));
		t.holder.scrollTop = 300;
		await nextFrames(4);
		assert.deepEqual(mouseUps, [301], "mouseup handlers see the last drag position");
		assert.notOk(t.scroll.getIsLockedMouse(), "mouse is up");
		assert.strictEqual(t.holder.scrollTop, 250, "native position follows the snap after the drag");
		assert.strictEqual(t.scroll.getCurScrolledY(), 250, "snapped position");
	});

	QUnit.test("native range is in pixels of the element whatever the caller's units", async function (assert)
	{
		// a form field: 100 units of content, 20 units visible in a 200 px element
		let t = createVertical(100, 20);
		assert.strictEqual(t.scroll.getMaxScrolledY(), 80, "max scroll in caller units");
		assert.strictEqual(t.holder.scrollHeight - t.holder.clientHeight, 800, "native range shows a fifth of the content");

		t.holder.scrollTop = 400;
		await nextFrames(3);
		assert.deepEqual(t.events, [{pos : 40, max : 80}], "native pixels map back to caller units");
	});

	QUnit.test("the last native pixel is not the end", async function (assert)
	{
		let t = createVertical(1000, 200);
		t.holder.scrollTop = 799;
		await nextFrames(3);
		assert.strictEqual(t.scroll.getCurScrolledY(), 799, "799 of 800");
	});

	QUnit.test("a lost mouseup does not keep the scrollbar locked", async function (assert)
	{
		let t = createVertical(1000, 200);
		t.holder.dispatchEvent(new MouseEvent("mousedown", {button : 0, bubbles : true}));
		window.dispatchEvent(new MouseEvent("mousemove", {buttons : 1}));
		await nextFrames(4);
		assert.ok(t.scroll.getIsLockedMouse(), "a move with the button held keeps the drag");

		// released outside the frame: the next move has no buttons
		window.dispatchEvent(new MouseEvent("mousemove", {buttons : 0}));
		await nextFrames(4);
		assert.notOk(t.scroll.getIsLockedMouse(), "a move without buttons ends the drag");

		t.holder.dispatchEvent(new MouseEvent("mousedown", {button : 0, bubbles : true}));
		window.dispatchEvent(new Event("blur"));
		await nextFrames(4);
		assert.notOk(t.scroll.getIsLockedMouse(), "blur ends the drag");

		t.scroll.scrollToY(300);
		await nextFrames(3);
		assert.strictEqual(t.holder.scrollTop, 300, "programmatic scrolls move the bar again");
	});

	QUnit.test("pointerup handlers run after a drag", async function (assert)
	{
		// the cell editor binds AscCommon.getPtrEvtName("up")
		let t = createVertical(1000, 200);
		let ups = [];
		t.scroll.bind("pointerup", function ()
		{
			ups.push(t.scroll.scrollerMouseDown);
		});
		t.holder.dispatchEvent(new MouseEvent("mousedown", {button : 0, bubbles : true}));
		window.dispatchEvent(new MouseEvent("mouseup", {button : 0}));
		await nextFrames(4);
		assert.deepEqual(ups, [true], "pointerup fires while the drag is still flagged");
	});

	QUnit.test("right-to-left horizontal scrollbar", async function (assert)
	{
		let holder = createHolder(300, 14);
		let settings = new AscCommon.ScrollSettings();
		settings.isVerticalScroll = false;
		settings.isHorizontalScroll = true;
		settings.contentW = 1300;
		settings.screenW = 300;
		let scroll = new AscCommon.ScrollObject(holder.id, settings);
		let events = [];
		scroll.bind("scrollhorizontal", function (evt)
		{
			events.push(evt.scrollD);
		});
		scroll.setRightToLeft(true);

		scroll.scrollToX(100);
		await nextFrames(3);
		assert.strictEqual(holder.scrollLeft, -100, "position 100 is 100 px from the right");

		holder.scrollLeft = -400;
		await nextFrames(3);
		assert.deepEqual(events, [100, 400], "native scroll to the left");
	});
})();
