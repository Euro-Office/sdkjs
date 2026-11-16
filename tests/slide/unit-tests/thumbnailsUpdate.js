/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	QUnit.module("CThumbnailsManager.onCheckUpdate");

	// A thumbnails manager with no slides in view that counts the redraws (OnPaint).
	function createThumbnails()
	{
		let thumbnails = new CThumbnailsManager();
		thumbnails.isThumbnailsShown = function () { return true; };
		thumbnails.DigitWidths = [1];
		thumbnails.m_oWordControl = {
			m_oApi             : {isSaveFonts_Images : false},
			m_oDrawingDocument : {IsEmptyPresentation : false}
		};
		thumbnails.m_lDrawingFirst = 1;
		thumbnails.m_lDrawingEnd = 0;
		thumbnails.paints = 0;
		thumbnails.OnPaint = function () { this.paints++; };
		return thumbnails;
	}

	function withUsers(singleUser, fn)
	{
		let collaborativeEditing = AscCommon.CollaborativeEditing;
		AscCommon.CollaborativeEditing = {Is_SingleUser : function () { return singleUser; }};
		try { fn(); } finally { AscCommon.CollaborativeEditing = collaborativeEditing; }
	}

	QUnit.test("While others edit, thumbnails redraw at most every 2 s", function (assert)
	{
		withUsers(false, function ()
		{
			let thumbnails = createThumbnails();
			thumbnails.m_bIsUpdate = true;
			thumbnails.onCheckUpdate();
			assert.strictEqual(thumbnails.paints, 1, "the first update redraws");
			assert.strictEqual(thumbnails.m_bIsUpdate, false, "nothing pending after a redraw");

			thumbnails.m_bIsUpdate = true;
			thumbnails.onCheckUpdate();
			assert.strictEqual(thumbnails.paints, 1, "an update within 2 s does not redraw");
			assert.strictEqual(thumbnails.m_bIsUpdate, true, "the skipped update stays pending");

			thumbnails.m_nLastUpdateTime -= 2000;
			thumbnails.onCheckUpdate();
			assert.strictEqual(thumbnails.paints, 2, "the pending update redraws after 2 s");
		});
	});

	QUnit.test("A single user's thumbnails redraw on each update", function (assert)
	{
		withUsers(true, function ()
		{
			let thumbnails = createThumbnails();
			for (let i = 0; i < 3; ++i)
			{
				thumbnails.m_bIsUpdate = true;
				thumbnails.onCheckUpdate();
			}
			assert.strictEqual(thumbnails.paints, 3, "three updates, three redraws");
		});
	});
});
