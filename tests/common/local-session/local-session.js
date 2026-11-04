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

		if (transport)
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

	QUnit.test("Closing the adapter closes the transport without reconnecting", function (assert)
	{
		var transport = createFakeTransport();
		var disconnects = [];
		var coApi = createCoApi(transport, {
			onDisconnect : function (reason, code) { disconnects.push(code); }
		});
		var socket = coApi._CoAuthoringApi.socketio;
		assert.ok(socket, "adapter is used as the socket");

		coApi.disconnect(AscCommon.c_oCloseCode.quiet, "done");
		assert.strictEqual(transport.closed, 1, "transport.close was called once");
		assert.strictEqual(coApi._CoAuthoringApi.get_state(), AscCommon.ConnectionState.ClosedAll, "state is ClosedAll, not Reconnect");
		assert.deepEqual(disconnects, [AscCommon.c_oCloseCode.quiet], "onDisconnect reports the explicit code once");

		socket.close();
		socket.disconnect();
		assert.strictEqual(transport.closed, 1, "closing again is a no-op");

		transport.handlers["onClose"](false);
		assert.strictEqual(coApi._CoAuthoringApi.get_state(), AscCommon.ConnectionState.Reconnect, "an unexpected close reports Reconnect state to the api");
		assert.strictEqual(disconnects.length, 2, "an unexpected close reports a disconnect");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("Without a transport the socket.io path is used", function (assert)
	{
		AscCommon.setLocalSessionTransport(null);
		var origGetSocketIO = AscCommon.getSocketIO;
		var ioOptions = null;
		var fakeSocket = {
			events : {},
			on : function (name, cb) { this.events[name] = cb; },
			io : {on : function () {}, opts : {}},
			emit : function () {},
			disconnect : function () {}
		};
		AscCommon.getSocketIO = function ()
		{
			return function (options)
			{
				ioOptions = options;
				return fakeSocket;
			};
		};
		try
		{
			var coApi = createCoApi(null, {});
			assert.strictEqual(coApi._CoAuthoringApi.socketio, fakeSocket, "socket.io client is used");
			assert.ok(ioOptions && /doc1/.test(ioOptions["path"]), "socket.io is opened on the document path");
			assert.ok(fakeSocket.events["message"], "message handler is registered on the socket");
		}
		finally
		{
			AscCommon.getSocketIO = origGetSocketIO;
		}
	});

	QUnit.module("Local session download and upload");

	function createApi()
	{
		var api = new AscCommon.baseEditorsApi({});
		api.documentId = "doc1";
		api.documentTitle = "report.docx";
		return api;
	}

	function createDownloadData(isPrint)
	{
		return {
			oAdditionalData : {
				"c" : "save",
				"id" : "doc1",
				"outputformat" : Asc.c_oAscFileType.PDF,
				"title" : "report.pdf"
			},
			dataContainer : {data : "BIN", part : null, index : 0, count : 0},
			downloadType : isPrint ? AscCommon.DownloadType.Print : AscCommon.DownloadType.Download
		};
	}

	QUnit.test("downloadAs converts through the transport", function (assert)
	{
		var requests = [];
		var transport = {
			convert : function (request, cb)
			{
				requests.push(request);
				cb({"status" : "ok", "data" : "blob:converted", "filetype" : "pdf"});
			}
		};
		AscCommon.setLocalSessionTransport(transport);
		var api = createApi();
		var urls = [];
		var errors = [];
		api.sendEvent = function (name, a, b)
		{
			if ("asc_onPrintUrl" === name)
				urls.push(a + ":" + b);
			else if ("asc_onError" === name)
				errors.push(a);
		};

		var d = createDownloadData(true);
		api._downloadAsUsingServer(Asc.c_oAscAsyncAction.Print, {}, d.oAdditionalData, d.dataContainer, d.downloadType);

		assert.strictEqual(requests.length, 1, "transport.convert was called once");
		var request = requests[0];
		assert.strictEqual(request["data"], "BIN", "request carries the editor data");
		assert.strictEqual(request["fileType"], Asc.c_oAscFileType.PDF, "request carries the numeric file type");
		assert.strictEqual(request["format"], "pdf", "request carries the extension");
		assert.strictEqual(request["title"], "report.pdf", "request carries the title");
		assert.strictEqual(request["isPrint"], true, "request flags printing");
		assert.strictEqual(request["options"], d.oAdditionalData, "request carries the downloadas parameters");
		assert.deepEqual(urls, ["blob:converted:pdf"], "the converted url reaches the same download event as the server path");
		assert.deepEqual(errors, [], "no error is raised");

		var callbackInput = null;
		var d2 = createDownloadData(false);
		api._downloadAsUsingServer(null, {callback : function (input) { callbackInput = input; }}, d2.oAdditionalData, d2.dataContainer, d2.downloadType);
		assert.ok(callbackInput && "ok" === callbackInput["status"] && "blob:converted" === callbackInput["data"], "a caller supplied callback receives the server shaped reply");
		assert.strictEqual(requests[1]["isPrint"], false, "download is not flagged as print");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("downloadAs reports a conversion error from the transport", function (assert)
	{
		AscCommon.setLocalSessionTransport({
			convert : function (request, cb) { cb({"status" : "err", "error" : 1}); }
		});
		var api = createApi();
		var urls = [];
		var errors = [];
		api.sendEvent = function (name, a, b)
		{
			if ("asc_onDownloadUrl" === name)
				urls.push(a);
			else if ("asc_onError" === name)
				errors.push(a + ":" + b);
		};

		var d = createDownloadData(false);
		api._downloadAsUsingServer(Asc.c_oAscAsyncAction.DownloadAs, {}, d.oAdditionalData, d.dataContainer, d.downloadType);

		assert.deepEqual(urls, [], "no url is delivered");
		assert.deepEqual(errors, [Asc.c_oAscError.ID.ConvertationSaveError + ":" + Asc.c_oAscError.Level.NoCritical], "a non critical conversion save error is raised");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("Image upload goes through the transport", function (assert)
	{
		var uploaded = null;
		AscCommon.setLocalSessionTransport({
			uploadImages : function (files, cb)
			{
				uploaded = files;
				cb(Asc.c_oAscError.ID.No, ["blob:img1"]);
			}
		});
		var result = null;
		var files = [{name : "a.png"}];
		AscCommon.UploadImageFiles(files, "doc1", "u1", "", undefined, undefined, undefined, function (error, urls)
		{
			result = {error : error, urls : urls};
		});
		assert.strictEqual(uploaded, files, "transport.uploadImages receives the files");
		assert.deepEqual(result, {error : Asc.c_oAscError.ID.No, urls : ["blob:img1"]}, "error and urls are forwarded");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("Image upload fails cleanly without transport support", function (assert)
	{
		AscCommon.setLocalSessionTransport({});
		var errors = [];
		AscCommon.UploadImageFiles([{name : "a.png"}], "doc1", "u1", "", undefined, undefined, undefined, function (error) { errors.push(error); });
		AscCommon.UploadImageUrls(["http://example.com/a.png"], "doc1", "u1", "", undefined, undefined, undefined, function (error) { errors.push(error); });
		assert.deepEqual(errors, [Asc.c_oAscError.ID.UplImageUrl, Asc.c_oAscError.ID.UplImageUrl], "both upload paths report the upload url error");
		AscCommon.setLocalSessionTransport(null);
	});
});
