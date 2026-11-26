/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	QUnit.module("AscCommon.CCollaborativeEditingBase fast sync pace");

	QUnit.test("No pause before the first sync", function (assert)
	{
		let coedit = new AscCommon.CCollaborativeEditingBase();
		assert.strictEqual(coedit.IsFastSyncPaused(), false);
	});

	QUnit.test("After a sync, the next one waits three times as long as the sync took", function (assert)
	{
		let coedit = new AscCommon.CCollaborativeEditingBase();
		let before = new Date().getTime();
		coedit.OnFastSync(before - 100);
		let after = new Date().getTime();
		assert.strictEqual(coedit.IsFastSyncPaused(), true, "paused");
		assert.ok(coedit.m_nFastSyncNextTime >= before + 300, "at least 300 ms after a sync of 100 ms");
		assert.ok(coedit.m_nFastSyncNextTime <= after + 3 * (after - before + 100), "at most 3 times the sync");
	});

	QUnit.test("The wait is at most 5 s", function (assert)
	{
		let coedit = new AscCommon.CCollaborativeEditingBase();
		let before = new Date().getTime();
		coedit.OnFastSync(before - 60000);
		let after = new Date().getTime();
		assert.ok(coedit.m_nFastSyncNextTime >= before + 5000, "5 s");
		assert.ok(coedit.m_nFastSyncNextTime <= after + 5000, "not more than 5 s");
	});
});
