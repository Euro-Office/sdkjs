/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	// The auto-scroll with the middle mouse button (pdf/src/viewer.js): the speed and when it stops.
	// The DOM part (listeners, marker, cursor) is tested with a stand-in for the viewer and simulated events.
	const CAutoScroll = AscPDF.CAutoScroll;

	QUnit.module("PDF auto-scroll");

	QUnit.test("Speed grows with the distance from the start point", function (assert)
	{
		assert.strictEqual(CAutoScroll.getSpeed(0), 0, "no scrolling at the start point");
		assert.strictEqual(CAutoScroll.getSpeed(CAutoScroll.DEAD_ZONE), 0, "no scrolling inside the dead zone");
		assert.strictEqual(CAutoScroll.getSpeed(-CAutoScroll.DEAD_ZONE), 0, "no scrolling inside the dead zone, upwards");

		let near = CAutoScroll.getSpeed(30);
		let far  = CAutoScroll.getSpeed(120);
		assert.ok(near > 0, "moving down scrolls down");
		assert.ok(far > near, "further away scrolls faster");
		assert.strictEqual(CAutoScroll.getSpeed(-30), -near, "moving up scrolls up at the same speed");
		assert.strictEqual(CAutoScroll.getSpeed(30), (30 - CAutoScroll.DEAD_ZONE) / CAutoScroll.SPEED_DIVISOR, "speed formula");
		assert.strictEqual(CAutoScroll.getSpeed(100000), CAutoScroll.MAX_SPEED, "speed is capped");
		assert.strictEqual(CAutoScroll.getSpeed(-100000), -CAutoScroll.MAX_SPEED, "speed is capped upwards");
	});

	QUnit.test("Step in both directions", function (assert)
	{
		let state = new CAutoScroll(100, 100);
		assert.deepEqual(state.getStep(), {x : 0, y : 0}, "no step before the mouse moves");

		state.onMove(105, 160);
		let step = state.getStep();
		assert.strictEqual(step.x, 0, "a small sideways move doesn't scroll sideways");
		assert.strictEqual(step.y, CAutoScroll.getSpeed(60), "vertical step from the distance");

		state.onMove(40, 100);
		step = state.getStep();
		assert.ok(step.x < 0, "left of the start point scrolls left");
		assert.strictEqual(step.y, 0, "back at the start height: no vertical step");
	});

	QUnit.test("Hold the button and move: the release stops", function (assert)
	{
		let state = new CAutoScroll(100, 100);
		assert.strictEqual(state.onButtonDown(), false, "a press while the button is held doesn't stop");
		state.onMove(100, 200);
		assert.strictEqual(state.moved, true, "the mouse has moved");
		assert.strictEqual(state.onButtonUp(), true, "the release stops");
	});

	QUnit.test("Click once: it goes on until the next click", function (assert)
	{
		let state = new CAutoScroll(100, 100);
		state.onMove(103, 104);
		assert.strictEqual(state.moved, false, "a few pixels are not a move");
		assert.strictEqual(state.onButtonUp(), false, "the release of a click doesn't stop");
		assert.strictEqual(state.held, false, "the button is no longer held");

		state.onMove(100, 300);
		assert.ok(state.getStep().y > 0, "it scrolls when the mouse moves after the click");
		assert.strictEqual(state.onButtonUp(), false, "a second release is ignored");
		assert.strictEqual(state.onButtonDown(), true, "the next click stops");
	});

	function createViewer()
	{
		return {autoScroll : null, m_oScrollVerApi : null, m_oScrollHorApi : null, isVisibleHorScroll : false};
	}
	function mouseEvent(type, props)
	{
		return new MouseEvent(type, Object.assign({bubbles : true, cancelable : true}, props));
	}

	QUnit.test("Start and stop: marker and scroll cursor", function (assert)
	{
		let viewer = createViewer();
		let styles = document.querySelectorAll("style").length;
		CAutoScroll.start(viewer, 100, 100);
		assert.ok(viewer.autoScroll, "auto-scroll is on");
		assert.ok(viewer.autoScroll.marker.parentNode, "the marker is shown");
		assert.strictEqual(document.querySelectorAll("style").length, styles + 1, "a scroll cursor rule is added");
		assert.ok(-1 !== viewer.autoScroll.cursorStyle.textContent.indexOf(CAutoScroll.CURSOR), "with the scroll cursor");

		let marker = viewer.autoScroll.marker;
		CAutoScroll.stop(viewer);
		assert.strictEqual(viewer.autoScroll, null, "auto-scroll is off");
		assert.notOk(marker.parentNode, "the marker is removed");
		assert.strictEqual(document.querySelectorAll("style").length, styles, "the normal cursor is back");
	});

	function pointerEvent(type, props)
	{
		return new PointerEvent(type, Object.assign({bubbles : true, cancelable : true, pointerType : "mouse"}, props));
	}
	// an element in place of the viewer canvas, which gets pointer and mouse events (viewer.js: onpointerup -> onMouseUp)
	function createTarget()
	{
		let target = document.createElement("div");
		document.body.appendChild(target);
		target.received = [];
		["pointerdown", "pointerup", "mouseup", "click", "auxclick", "contextmenu", "keydown"].forEach(function (type)
		{
			target.addEventListener(type, function (ev) {target.received.push(ev.type);});
		});
		return target;
	}
	// the events of one click, as the browser sends them; returns whether each one's default action was prevented
	function sendRelease(target, button)
	{
		let prevented = {};
		[
			pointerEvent("pointerup", {button : button}),
			mouseEvent("mouseup", {button : button}),
			1 === button ? new PointerEvent("auxclick", {bubbles : true, cancelable : true, button : 1}) : new PointerEvent("click", {bubbles : true, cancelable : true, button : 0})
		].forEach(function (ev)
		{
			target.dispatchEvent(ev);
			prevented[ev.type] = ev.defaultPrevented;
		});
		return prevented;
	}

	QUnit.test("The click that stops the auto-scroll does nothing else", function (assert)
	{
		let done = assert.async();
		let viewer = createViewer();
		let target = createTarget();
		CAutoScroll.start(viewer, 100, 100);
		// the release of the starting middle click, without a move: the auto-scroll goes on until the next click
		let prevented = sendRelease(target, 1);
		assert.ok(viewer.autoScroll, "a middle click starts it and it goes on");
		assert.deepEqual(target.received, [], "the release of the starting click doesn't reach the viewer");
		assert.ok(prevented["mouseup"] && prevented["auxclick"], "its default action (e.g. a paste on Linux) is prevented");

		setTimeout(function ()
		{
			let down = pointerEvent("pointerdown", {button : 0});
			target.dispatchEvent(down);
			assert.strictEqual(viewer.autoScroll, null, "a left click stops it");
			assert.ok(down.defaultPrevented, "the press is prevented");

			prevented = sendRelease(target, 0);
			assert.deepEqual(target.received, [], "neither the press nor pointerup, mouseup or click reach the viewer");
			assert.ok(prevented["pointerup"] && prevented["mouseup"] && prevented["click"], "and their default actions are prevented");

			setTimeout(function ()
			{
				target.dispatchEvent(pointerEvent("pointerdown", {button : 0}));
				sendRelease(target, 0);
				assert.deepEqual(target.received, ["pointerdown", "pointerup", "mouseup", "click"], "the next click works as usual");
				document.body.removeChild(target);
				done();
			}, 10);
		}, 10);
	});

	QUnit.test("The middle click that stops the auto-scroll does nothing else", function (assert)
	{
		let done = assert.async();
		let viewer = createViewer();
		let target = createTarget();
		CAutoScroll.start(viewer, 100, 100);
		sendRelease(target, 1);

		setTimeout(function ()
		{
			target.dispatchEvent(pointerEvent("pointerdown", {button : 1}));
			assert.strictEqual(viewer.autoScroll, null, "a second middle click stops it");
			let prevented = sendRelease(target, 1);
			assert.deepEqual(target.received, [], "its release doesn't reach the viewer");
			assert.ok(prevented["pointerup"] && prevented["mouseup"] && prevented["auxclick"], "no default action (no paste of the PRIMARY selection)");

			setTimeout(function ()
			{
				document.body.removeChild(target);
				done();
			}, 10);
		}, 10);
	});

	QUnit.test("Hold, move and release: the release does nothing else", function (assert)
	{
		let done = assert.async();
		let viewer = createViewer();
		let target = createTarget();
		CAutoScroll.start(viewer, 100, 100);
		document.dispatchEvent(mouseEvent("mousemove", {clientX : 100, clientY : 300}));
		let prevented = sendRelease(target, 1);
		assert.strictEqual(viewer.autoScroll, null, "the release stops it");
		assert.deepEqual(target.received, [], "pointerup, mouseup and auxclick don't reach the viewer");
		assert.ok(prevented["mouseup"] && prevented["auxclick"], "and have no default action");

		setTimeout(function ()
		{
			document.body.removeChild(target);
			done();
		}, 10);
	});

	QUnit.test("The Esc that stops the auto-scroll does nothing else", function (assert)
	{
		let viewer = createViewer();
		let target = createTarget();
		CAutoScroll.start(viewer, 100, 100);

		let other = new KeyboardEvent("keydown", {bubbles : true, cancelable : true, keyCode : 40});
		target.dispatchEvent(other);
		assert.ok(viewer.autoScroll, "another key doesn't stop it");
		assert.notOk(other.defaultPrevented, "and is not consumed");

		let esc = new KeyboardEvent("keydown", {bubbles : true, cancelable : true, keyCode : 27});
		target.dispatchEvent(esc);
		assert.strictEqual(viewer.autoScroll, null, "Esc stops it");
		assert.ok(esc.defaultPrevented, "the Esc is consumed");
		assert.deepEqual(target.received, ["keydown"], "only the other key reached the viewer");

		let next = new KeyboardEvent("keydown", {bubbles : true, cancelable : true, keyCode : 27});
		target.dispatchEvent(next);
		assert.deepEqual(target.received, ["keydown", "keydown"], "the next Esc works as usual");
		document.body.removeChild(target);
	});

	QUnit.test("Leaving the window stops", function (assert)
	{
		let viewer = createViewer();
		CAutoScroll.start(viewer, 100, 100);
		document.dispatchEvent(mouseEvent("mousemove", {clientX : 100, clientY : 400}));
		assert.ok(viewer.autoScroll.state.getStep().y > 0, "it scrolls down");

		document.body.dispatchEvent(mouseEvent("mouseout", {relatedTarget : document.documentElement}));
		assert.ok(viewer.autoScroll, "moving from one element to another doesn't stop");

		document.body.dispatchEvent(mouseEvent("mouseout", {relatedTarget : null}));
		assert.strictEqual(viewer.autoScroll, null, "leaving the window stops");

		CAutoScroll.stop(viewer);
	});

	QUnit.test("Hiding the page or losing the focus stops", function (assert)
	{
		let viewer = createViewer();
		CAutoScroll.start(viewer, 100, 100);
		document.dispatchEvent(new Event("visibilitychange"));
		assert.ok(viewer.autoScroll, "a visible page goes on");

		Object.defineProperty(document, "hidden", {configurable : true, get : function() {return true;}});
		try
		{
			document.dispatchEvent(new Event("visibilitychange"));
		}
		finally
		{
			delete document.hidden;
		}
		assert.strictEqual(viewer.autoScroll, null, "a hidden page stops");

		CAutoScroll.start(viewer, 100, 100);
		window.dispatchEvent(new Event("blur"));
		assert.strictEqual(viewer.autoScroll, null, "losing the focus stops");
		CAutoScroll.stop(viewer);
	});
});
