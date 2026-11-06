/*
 * SPDX-FileCopyrightText: Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

$(function ()
{
	function checkUserColor(assert, userId)
	{
		const normal = AscCommon.getUserColorById(userId, null, false, false);
		assert.ok(normal, 'Normal color is stored');
		assert.deepEqual([normal.r, normal.g, normal.b], [8, 145, 178], 'Normal color keeps the given rgb');
		assert.strictEqual(AscCommon.getUserColorById(userId, null, false, true), 0x0891B2, 'Numeric Normal color');
		assert.ok(AscCommon.getUserColorById(userId, null, -1, false), 'Light color is stored');
		assert.ok(AscCommon.getUserColorById(userId, null, true, false), 'Dark color is stored');
	}

	QUnit.module('User color');

	QUnit.test('setUserColorById stores Normal, Light and Dark colors', function (assert)
	{
		AscCommon.setUserColorById('test-user-color', {r : 8, g : 145, b : 178});
		checkUserColor(assert, 'test-user-color');
	});

	QUnit.test('SetAssistantTrackRevisions gives the assistant a full user color', function (assert)
	{
		const editor = AscTest.Editor;
		const savedDocInfo = editor.DocInfo;
		editor.DocInfo = new Asc.asc_CDocInfo();
		editor.DocInfo.put_UserInfo(new Asc.asc_CUserInfo());

		const doc = AscTest.JsApi.GetDocument();
		try
		{
			doc.SetAssistantTrackRevisions(true, 'Assistant');
			checkUserColor(assert, 'uid-Assistant');
			doc.SetAssistantTrackRevisions(false);
		}
		finally
		{
			editor.DocInfo = savedDocInfo;
		}
	});
});
