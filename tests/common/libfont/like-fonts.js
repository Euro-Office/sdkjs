/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

// The "like fonts" groups in common/libfont/map.js must match the ones in core's
// DesktopEditor/fontengine/ApplicationFonts.h (CFontListNamePicker), otherwise the editor and x2t
// pick different fonts for the same missing font. These are the same checks as core's
// CFontListNamePicker.MetricAliasesFindTheirTwin test.
$(function ()
{
	QUnit.module("Metric twins of fonts");

	QUnit.test("A missing font is matched by its metric twin, also under its aliases", function (assert)
	{
		const dict = AscFonts.g_fontApplication.g_fontDictionary;

		["Times New Roman", "Times", "Thorndale", "Thorndale AMT", "Tinos", "Nimbus Roman"].forEach(function (requested)
		{
			assert.ok(dict.CheckLikeFonts("Liberation Serif", requested), "Liberation Serif for " + requested);
		});
		["Courier New", "Courier", "Cumberland", "Cumberland AMT", "Cousine", "Nimbus Mono PS", "Nimbus Mono"].forEach(function (requested)
		{
			assert.ok(dict.CheckLikeFonts("Liberation Mono", requested), "Liberation Mono for " + requested);
		});
		["Arial", "Helvetica", "Albany", "Albany AMT", "Arimo"].forEach(function (requested)
		{
			assert.ok(dict.CheckLikeFonts("Liberation Sans", requested), "Liberation Sans for " + requested);
		});
	});

	QUnit.test("The other clones are offered as twins where Liberation is not installed", function (assert)
	{
		const dict = AscFonts.g_fontApplication.g_fontDictionary;

		["Tinos", "Nimbus Roman", "Nimbus Roman No9 L"].forEach(function (candidate)
		{
			assert.ok(dict.CheckLikeFonts(candidate, "Times New Roman"), candidate + " for Times New Roman");
		});
		["Cousine", "Nimbus Mono PS", "Nimbus Mono", "Nimbus Mono L"].forEach(function (candidate)
		{
			assert.ok(dict.CheckLikeFonts(candidate, "Courier New"), candidate + " for Courier New");
		});
		["Arimo", "Nimbus Sans"].forEach(function (candidate)
		{
			assert.ok(dict.CheckLikeFonts(candidate, "Arial"), candidate + " for Arial");
		});
	});

	QUnit.test("Only twins match", function (assert)
	{
		const dict = AscFonts.g_fontApplication.g_fontDictionary;

		assert.notOk(dict.CheckLikeFonts("Liberation Sans", "Times"), "no sans for Times");
		assert.notOk(dict.CheckLikeFonts("Liberation Serif", "Courier"), "no serif for Courier");
		assert.ok(dict.CheckLikeFonts("Caladea", "Cambria"), "Caladea for Cambria");
		assert.notOk(dict.CheckLikeFonts("Liberation Serif", "Verdana"), "nothing for Verdana");
		// Names that only contain a twin's name are not in the groups.
		assert.notOk(dict.CheckLikeFonts("Liberation Mono", "Courier 10 Pitch"), "no twin for Courier 10 Pitch");
	});
});
