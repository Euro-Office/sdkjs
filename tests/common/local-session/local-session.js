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

	QUnit.module("Local session transport");

	function createFakeTransport()
	{
		return {
			handlers : null,
			sent : [],
			closed : 0,
			connect : function (handlers)
			{
				this.handlers = handlers;
				handlers["onOpen"]();
			},
			send : function (data)
			{
				this.sent.push(data);
			},
			close : function ()
			{
				this.closed++;
			}
		};
	}

	function createCoApi(transport, callbacks)
	{
		window.Asc.editor = window.Asc.editor || {};
		if (!window.Asc.editor.isPdfEditor)
			window.Asc.editor.isPdfEditor = function () { return false; };

		AscCommon.setLocalSessionTransport(transport);

		var coApi = new AscCommon.CDocsCoApi();
		coApi._CoAuthoringApi.set_url("local");
		for (var name in callbacks)
			coApi[name] = callbacks[name];

		var user = new AscCommon.asc_CUser({"id" : "u10", "idOriginal" : "u1", "username" : "Alice", "indexUser" : 0, "view" : true});

		var docInfo = new Asc.asc_CDocInfo();
		docInfo.put_Mode("view");
		docInfo.put_Permissions({"edit" : false});
		docInfo.put_Lang("en");

		var openCmd = {"c" : "open", "id" : "doc1", "format" : "docx"};
		coApi.init(user, "doc1", "", "", AscCommon.c_oEditorId.Word, "docx", docInfo, undefined, undefined, undefined, undefined, openCmd);
		return coApi;
	}

	QUnit.test("Registration", function (assert)
	{
		var transport = createFakeTransport();
		AscCommon.setLocalSessionTransport(transport);
		assert.strictEqual(AscCommon.getLocalSessionTransport(), transport, "transport is registered");
		AscCommon.setLocalSessionTransport(null);
		assert.strictEqual(AscCommon.getLocalSessionTransport(), null, "transport is cleared");
	});

	QUnit.test("Auth goes through the transport and replies reach the api", function (assert)
	{
		var transport = createFakeTransport();
		var events = [];
		var coApi = createCoApi(transport, {
			onFirstConnect : function () { events.push("firstConnect"); },
			onLicense : function (license) { events.push("license:" + license["type"]); },
			onAuthParticipantsChanged : function (participants, userId) { events.push("participants:" + Object.keys(participants).length); },
			onDocumentOpen : function (data) { events.push("documentOpen:" + data["data"]["data"]["Editor.bin"]); }
		});

		assert.ok(transport.handlers, "transport.connect was called");
		assert.deepEqual(events, ["firstConnect"], "onOpen triggers onFirstConnect");
		assert.strictEqual(transport.sent.length, 0, "nothing is sent before auth");

		coApi.auth(true, {"c" : "open", "id" : "doc1"});
		assert.strictEqual(transport.sent.length, 1, "auth message was sent");
		var auth = transport.sent[0];
		assert.strictEqual(auth["type"], "auth", "message type is auth");
		assert.strictEqual(auth["docid"], "doc1", "auth carries the document id");
		assert.strictEqual(auth["user"]["id"], "u10", "auth carries the user id");
		assert.strictEqual(auth["mode"], "view", "auth carries the mode");
		assert.strictEqual(auth["openCmd"]["c"], "open", "auth carries the open command");

		transport.handlers["onMessage"]({
			"type" : "license",
			"license" : {"type" : 3, "mode" : 0, "rights" : 1, "buildVersion" : "9.3.8", "buildNumber" : 1}
		});
		transport.handlers["onMessage"]({"type" : "authChanges", "changes" : []});
		transport.handlers["onMessage"]({
			"type" : "auth",
			"result" : 1,
			"sessionId" : "s1",
			"participants" : [{"id" : "u10", "idOriginal" : "u1", "username" : "Alice", "indexUser" : 0, "view" : true}],
			"locks" : [],
			"changes" : [],
			"changesIndex" : 0,
			"indexUser" : 0,
			"buildVersion" : "9.3.8",
			"buildNumber" : 1,
			"licenseType" : 3
		});
		assert.ok(coApi.get_isAuth(), "api is authorized after the auth reply");
		assert.ok(events.indexOf("license:3") !== -1, "license reply reached onLicense");
		assert.ok(events.indexOf("participants:1") !== -1, "participants reached onAuthParticipantsChanged");

		transport.handlers["onMessage"]({
			"type" : "documentOpen",
			"data" : {"type" : "open", "status" : "ok", "data" : {"Editor.bin" : "blob:editor-bin"}}
		});
		assert.ok(events.indexOf("documentOpen:blob:editor-bin") !== -1, "documentOpen reached onDocumentOpen");

		coApi.disconnect(AscCommon.c_oCloseCode.quiet, "done");
		assert.strictEqual(transport.closed, 1, "disconnect closes the transport");

		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("Messages sent before auth are buffered", function (assert)
	{
		var transport = createFakeTransport();
		var coApi = createCoApi(transport, {});
		coApi.sendMessage("hello");
		assert.strictEqual(transport.sent.length, 1, "state WaitAuth already sends through the transport");
		AscCommon.setLocalSessionTransport(null);
	});
});
