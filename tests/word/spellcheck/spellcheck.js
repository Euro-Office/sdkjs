/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function () {

	let logicDocument = AscTest.CreateLogicDocument();

	function collectWords(text)
	{
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(0, p);
		AscTest.AddTextToParagraph(p, text);

		p.SpellChecker.Clear();
		p.private_CollectSpellCheckerElements(true);
		return p.SpellChecker.Elements.map(function(element) { return element.GetWord(); });
	}

	QUnit.module("Spell check word collection");

	QUnit.test("Emoji are not words", function (assert)
	{
		assert.deepEqual(collectWords("Hallo 😀🦄🌈 Welt"), ["Hallo", "Welt"], "standalone emoji");
		assert.deepEqual(collectWords("Hallo🦄Welt"), ["Hallo", "Welt"], "emoji between letters splits the word");
		assert.deepEqual(collectWords("👍🏽 🇩🇪 ❤️"), [], "skin tone, flag and heart");
		assert.deepEqual(collectWords("Hallo 🏳️‍🌈 ❤️‍🔥 Welt"), ["Hallo", "Welt"], "ZWJ sequences with variation selectors");
	});

	QUnit.test("Letters are still words", function (assert)
	{
		assert.deepEqual(collectWords("Straße über café"), ["Straße", "über", "café"], "BMP letters");
		// Deseret lowercase, above U+FFFF
		assert.deepEqual(collectWords("𐐨𐐯"), ["𐐨𐐯"], "letters above U+FFFF stay intact");
		assert.deepEqual(collectWords("𐐨 a"), [], "a single letter above U+FFFF is skipped like a single BMP letter");
		assert.deepEqual(collectWords("क्‍ष"), ["क्‍ष"], "ZWJ inside an Indic word");
	});
});
