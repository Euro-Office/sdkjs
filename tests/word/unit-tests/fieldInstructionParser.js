/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

$(function () {

	QUnit.module("Field instruction parser");

	function parseDate(instruction)
	{
		let parser = new CFieldInstructionParser();
		return parser.GetInstructionClass(instruction);
	}

	QUnit.test("Date format switch with and without a space", function (assert)
	{
		let withSpace = parseDate('DATE \\@ "dd.MM.yyyy"');
		assert.strictEqual(withSpace.Type, AscWord.fieldtype_DATE, "DATE with a space before the quoted format is a date field");
		assert.strictEqual(withSpace.GetFormat(), "dd.MM.yyyy", "Format read with a space");

		let noSpace = parseDate('DATE \\@"dd.MM.yyyy"');
		assert.strictEqual(noSpace.Type, AscWord.fieldtype_DATE, "DATE without a space before the quoted format is a date field");
		assert.strictEqual(noSpace.GetFormat(), "dd.MM.yyyy", "Format read without a space, as Word does");

		let noSpaceMore = parseDate('DATE \\@"d MMMM yyyy" \\* MERGEFORMAT');
		assert.strictEqual(noSpaceMore.GetFormat(), "d MMMM yyyy", "Format without a space followed by another switch");
	});
});
