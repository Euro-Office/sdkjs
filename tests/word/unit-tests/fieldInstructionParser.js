/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

$(function () {

	function parse(line)
	{
		return (new CFieldInstructionParser()).GetInstructionClass(line);
	}

	QUnit.module("Unit-tests for CFieldInstructionParser");


	QUnit.test("Date-time picture switch of TIME and DATE fields", function (assert)
	{
		let instruction = parse(" TIME \\@ \"dd.MM.yyyy\" ");
		assert.strictEqual(instruction.GetType(), AscWord.fieldtype_TIME, "TIME with a space after the switch: type");
		assert.strictEqual(instruction.GetFormat(), "dd.MM.yyyy", "TIME with a space after the switch: format");

		// Word writes the switch without a space when the date is inserted with "Update automatically"
		instruction = parse(" TIME \\@\"dd.MM.yyyy\" ");
		assert.strictEqual(instruction.GetType(), AscWord.fieldtype_TIME, "TIME without a space after the switch: type");
		assert.strictEqual(instruction.GetFormat(), "dd.MM.yyyy", "TIME without a space after the switch: format");

		instruction = parse(" DATE \\@\"d. MMMM yyyy\" ");
		assert.strictEqual(instruction.GetType(), AscWord.fieldtype_DATE, "DATE without a space after the switch: type");
		assert.strictEqual(instruction.GetFormat(), "d. MMMM yyyy", "DATE without a space after the switch: format with spaces");

		instruction = parse(" DATE \\@ \"dd.MM.yyyy\" \\* MERGEFORMAT ");
		assert.strictEqual(instruction.GetFormat(), "dd.MM.yyyy", "DATE followed by a general formatting switch: format");

		instruction = parse(" DATE ");
		assert.strictEqual(instruction.GetType(), AscWord.fieldtype_DATE, "DATE without switches: type");
		assert.strictEqual(instruction.GetFormat(), "", "DATE without switches: no format");
	});
});
