/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

$(function () {
	QUnit.module("Font like names");

	QUnit.test("StarSymbol falls back to OpenSymbol without remapping glyphs", function (assert) {
		const fontApplication = AscFonts.g_fontApplication;
		const dictionary = fontApplication.g_fontDictionary;

		assert.true(dictionary.CheckLikeFonts("OpenSymbol", "Symbol"), "Symbol is like OpenSymbol");
		assert.true(dictionary.CheckLikeFonts("OpenSymbol", "StarSymbol"), "StarSymbol is like OpenSymbol");
		assert.false(dictionary.CheckLikeFonts("Arial", "StarSymbol"), "StarSymbol is not like Arial");

		// OpenSymbol is StarSymbol renamed: same code points, so no Symbol PUA remap
		assert.notStrictEqual(fontApplication.CheckReplaceGlyphsMap("Symbol", {Name: "OpenSymbol"}), null, "Symbol is remapped");
		assert.strictEqual(fontApplication.CheckReplaceGlyphsMap("StarSymbol", {Name: "OpenSymbol"}), null, "StarSymbol is not remapped");
	});
});
