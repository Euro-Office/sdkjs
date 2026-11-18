/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	// The auto-scroll with the middle mouse button (pdf/src/viewer.js): the speed and when it stops.
	// The DOM part (listeners, marker, timer) is in the viewer and is not tested here.
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
});
