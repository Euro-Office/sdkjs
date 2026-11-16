/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	QUnit.module("AscFonts.CFontFile.SetTextMatrix");

	// A font file that counts the glyph cache clears. UpdateMatrix only passes the matrix to FreeType.
	function createFontFile()
	{
		let file = new AscFonts.CFontFile();
		file.clears = 0;
		file.ClearCache = function () { this.clears++; };
		file.UpdateMatrix = function () {};
		file.SetTextMatrix(1, 0, 0, 1, 0, 0);
		file.clears = 0;
		return file;
	}

	// A shape rotated by whole turns (n * 2pi) gives cos and sin that are 1 and 0 only up to rounding.
	QUnit.test("Identity with float noise keeps the glyph cache", function (assert)
	{
		let file = createFontFile();
		let c = Math.cos(2 * Math.PI), s = Math.sin(2 * Math.PI);
		file.SetTextMatrix(c, s, -s, c, 0, 0);
		c = 0.9999999999999998;
		s = 2.4492935982947064e-16;
		file.SetTextMatrix(c, s, -s, c, 0, 0);
		assert.strictEqual(file.clears, 0, "no cache clear");
		assert.strictEqual(file.m_bIsTransform, false, "no transform");
	});

	QUnit.test("A real rotation clears the glyph cache", function (assert)
	{
		let file = createFontFile();
		let c = Math.cos(0.3), s = Math.sin(0.3);
		file.SetTextMatrix(c, s, -s, c, 0, 0);
		assert.strictEqual(file.clears, 1, "one cache clear");
		assert.strictEqual(file.m_bIsTransform, true, "transform");
	});
});
