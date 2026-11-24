/*
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

$(function ()
{
	let logicDocument = AscTest.CreateLogicDocument();

	// undo of a comment change tells the interface about it, the test editor has no interface
	if (!editor.sync_ChangeCommentData)
		editor.sync_ChangeCommentData = function() {};

	function createCommentData(userName, text)
	{
		let data = new AscCommon.CCommentData();
		data.m_sText     = text;
		data.m_sUserName = userName;
		data.m_sUserId   = userName.toLowerCase();
		data.m_sInitials = userName.charAt(0);
		data.m_sTime     = "1700000000000";
		return data;
	}

	function addComments()
	{
		AscTest.ClearDocument();
		logicDocument.PushToContent(AscTest.CreateParagraph());

		let first = createCommentData("Erin", "Please check");
		first.Add_Reply(createCommentData("Frank", "Done"));
		let comment1 = new AscCommon.CComment(logicDocument.Comments, first);
		logicDocument.Comments.Add(comment1);
		let comment2 = new AscCommon.CComment(logicDocument.Comments, createCommentData("Grace", "Typo"));
		logicDocument.Comments.Add(comment2);
		return [comment1.GetId(), comment2.GetId()];
	}

	function getNames(ids)
	{
		let data1 = logicDocument.Comments.Get_ById(ids[0]).GetData();
		let data2 = logicDocument.Comments.Get_ById(ids[1]).GetData();
		return [data1.m_sUserName, data1.m_aReplies[0].m_sUserName, data2.m_sUserName];
	}

	// a stand-in for the editor: the test editor has no interface and no document info
	function createApi(isAuthorOverride)
	{
		let userInfo = new Asc.asc_CUserInfo();
		userInfo.asc_putId("uid-1");
		userInfo.asc_putFullName("Erin Example");
		let docInfo = new Asc.asc_CDocInfo();
		docInfo.asc_putUserInfo(userInfo);

		let api = {
			DocInfo                  : docInfo,
			changed                  : [],
			private_GetLogicDocument : function() { return logicDocument; },
			sync_ChangeCommentData   : function(id) { this.changed.push(id); }
		};
		AscCommon.baseEditorsApi.prototype.asc_setAuthorOverride.call(api, isAuthorOverride);
		return api;
	}

	QUnit.module("Author override");

	QUnit.test("Off by default", function (assert)
	{
		assert.strictEqual(editor.asc_getAuthorOverride(), false, "A new editor keeps the logged-in user as author");

		let api = createApi(undefined);
		assert.strictEqual(api.isAuthorOverride, false, "Only true switches it on");

		let currentName = AscCommon.UserInfoParser.getCurrentName();
		AscCommon.baseEditorsApi.prototype.asc_setLocalUserName.call(api, "Reviewer 1");
		assert.strictEqual(api.DocInfo.asc_getUserName(), "Erin Example", "The user name does not change");
		assert.strictEqual(AscCommon.UserInfoParser.getCurrentName(), currentName, "The current name does not change");

		let ids = addComments();
		Asc.asc_docs_api.prototype.asc_ChangeCommentAuthor.call(api, "Reviewer 1");
		assert.deepEqual(getNames(ids), ["Erin", "Frank", "Grace"], "The comment authors do not change");
		assert.deepEqual(api.changed, [], "The interface gets no change");
	});

	QUnit.test("Your name: new comments and changes get it, the user id stays", function (assert)
	{
		let api = createApi(true);
		let currentName = AscCommon.UserInfoParser.getCurrentName();

		AscCommon.baseEditorsApi.prototype.asc_setLocalUserName.call(api, "Reviewer 1");
		assert.strictEqual(api.DocInfo.asc_getUserName(), "Reviewer 1", "Tracked changes take the name from here");
		assert.strictEqual(api.DocInfo.asc_getUserInfo().asc_getId(), "uid-1", "The user id stays");
		assert.strictEqual(AscCommon.UserInfoParser.getCurrentName(), "Reviewer 1", "New comments take the name from here");

		AscCommon.baseEditorsApi.prototype.asc_setLocalUserName.call(api, "");
		assert.strictEqual(api.DocInfo.asc_getUserName(), "Reviewer 1", "An empty name is ignored");

		AscCommon.UserInfoParser.setCurrentName(currentName);
	});

	QUnit.test("Change comment authors: one comment or all comments and replies, as one undo step", function (assert)
	{
		let api = createApi(true);
		let ids = addComments();

		AscCommon.History.Clear();
		Asc.asc_docs_api.prototype.asc_ChangeCommentAuthor.call(api, "Reviewer 1", [ids[0]]);
		assert.deepEqual(getNames(ids), ["Reviewer 1", "Frank", "Grace"], "Only the chosen comment, not its replies");
		assert.strictEqual(logicDocument.Comments.Get_ById(ids[0]).GetData().m_sUserId, "erin", "The user id stays");
		assert.deepEqual(api.changed, [ids[0]], "The interface gets the changed comment");

		Asc.asc_docs_api.prototype.asc_ChangeCommentAuthor.call(api, "Reviewer 2");
		assert.deepEqual(getNames(ids), ["Reviewer 2", "Reviewer 2", "Reviewer 2"], "All comments and replies");

		Asc.asc_docs_api.prototype.asc_ChangeCommentAuthor.call(api, "");
		assert.deepEqual(getNames(ids), ["Reviewer 2", "Reviewer 2", "Reviewer 2"], "An empty name is ignored");

		logicDocument.Document_Undo();
		assert.deepEqual(getNames(ids), ["Reviewer 1", "Frank", "Grace"], "One undo step for all");
	});
});
