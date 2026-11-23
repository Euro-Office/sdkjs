/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

QUnit.module("Glyph bitmap");

QUnit.test("Color glyph keeps its colors and ignores the text color", function (assert)
{
	const bitmap = new AscFonts.CGlyphBitmap();
	bitmap.nWidth = 2;
	bitmap.nHeight = 1;
	bitmap.fromRGBA(new Uint8ClampedArray([255, 0, 0, 255, 0, 0, 255, 255]));

	// the editors tint every glyph with the run color before drawing it
	bitmap.oGlyphData.checkColor(0, 255, 0, 2, 1);

	const canvas = document.createElement("canvas");
	canvas.width = 2;
	canvas.height = 1;
	const ctx = canvas.getContext("2d");
	bitmap.draw(ctx, 0, 0);

	assert.deepEqual(Array.from(ctx.getImageData(0, 0, 2, 1).data), [255, 0, 0, 255, 0, 0, 255, 255]);
});
