/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
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
