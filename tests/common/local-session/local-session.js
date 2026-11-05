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

	// Transport callbacks reach DocsCoApi asynchronously (like socket.io events)
	function tick()
	{
		return new Promise(function (resolve) { setTimeout(resolve, 0); });
	}

	function createFakeTransport()
	{
		return {
			handlers : null,
			connects : 0,
			sent : [],
			closed : 0,
			onSend : null,
			connect : function (handlers)
			{
				this.connects++;
				this.handlers = handlers;
				handlers["onOpen"]();
			},
			send : function (data)
			{
				this.sent.push(data);
				if (this.onSend)
					this.onSend(data);
			},
			close : function ()
			{
				this.closed++;
			},
			reply : function (data)
			{
				this.handlers["onMessage"](data);
			},
			sentOfType : function (type)
			{
				return this.sent.filter(function (m) { return m["type"] === type; });
			}
		};
	}

	function createCoApi(transport, callbacks, editorType)
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

		var user = new AscCommon.asc_CUser({"id" : "u1", "idOriginal" : "u1", "username" : "Alice", "indexUser" : -1, "view" : false});

		var docInfo = new Asc.asc_CDocInfo();
		docInfo.put_Mode(callbacks && callbacks.mode || "view");
		docInfo.put_Permissions({"edit" : !!(callbacks && callbacks.mode === "edit")});
		docInfo.put_Lang("en");

		var openCmd = {"c" : "open", "id" : "doc1", "format" : "docx"};
		coApi.init(user, "doc1", "", "", undefined !== editorType ? editorType : AscCommon.c_oEditorId.Word, "docx", docInfo, undefined, undefined, undefined, undefined, openCmd);
		return coApi;
	}

	function participant(id, indexUser, view)
	{
		return {"id" : id + indexUser, "idOriginal" : id, "username" : id, "indexUser" : indexUser, "view" : !!view};
	}

	// Opens a session and answers auth like the DocService: the own user id is user.id + indexUser ("u11").
	async function createAuthorizedCoApi(callbacks, editorType)
	{
		var transport = createFakeTransport();
		callbacks = callbacks || {};
		callbacks.mode = "edit";
		var coApi = createCoApi(transport, callbacks, editorType);
		await tick();
		coApi.auth(false, {"c" : "open", "id" : "doc1"});
		transport.reply({"type" : "license", "license" : {"type" : 3, "mode" : 0, "rights" : 1, "buildVersion" : "9.3.8", "buildNumber" : 1}});
		transport.reply({
			"type" : "auth",
			"result" : 1,
			"sessionId" : "s1",
			"participants" : [participant("u1", 1)],
			"locks" : [],
			"indexUser" : 1,
			"buildVersion" : "9.3.8",
			"buildNumber" : 1,
			"licenseType" : 3
		});
		await tick();
		transport.sent = [];
		return {coApi : coApi, transport : transport, docsCoApi : coApi._CoAuthoringApi};
	}

	QUnit.test("Registration", function (assert)
	{
		var transport = createFakeTransport();
		AscCommon.setLocalSessionTransport(transport);
		assert.strictEqual(AscCommon.getLocalSessionTransport(), transport, "transport is registered");
		AscCommon.setLocalSessionTransport(null);
		assert.strictEqual(AscCommon.getLocalSessionTransport(), null, "transport is cleared");
	});

	QUnit.test("Auth goes through the transport and replies reach the api", async function (assert)
	{
		var transport = createFakeTransport();
		var events = [];
		var coApi = createCoApi(transport, {
			onFirstConnect : function () { events.push("firstConnect"); },
			onLicense : function (license) { events.push("license:" + license["type"]); },
			onAuthParticipantsChanged : function (participants, userId) { events.push("participants:" + Object.keys(participants).length + ":" + userId); },
			onDocumentOpen : function (data) { events.push("documentOpen:" + data["data"]["data"]["Editor.bin"]); }
		});

		assert.ok(transport.handlers, "transport.connect was called");
		assert.deepEqual(events, [], "onOpen is delivered asynchronously, after init returned");
		await tick();
		assert.deepEqual(events, ["firstConnect"], "onOpen triggers onFirstConnect");
		assert.ok(coApi.get_onlineWork(), "the api works online when onOpen arrives");
		assert.strictEqual(transport.sent.length, 0, "nothing is sent before auth");

		coApi.auth(true, {"c" : "open", "id" : "doc1"});
		assert.strictEqual(transport.sent.length, 1, "auth message was sent");
		var auth = transport.sent[0];
		assert.strictEqual(auth["type"], "auth", "message type is auth");
		assert.strictEqual(auth["docid"], "doc1", "auth carries the document id");
		assert.strictEqual(auth["user"]["id"], "u1", "auth carries the user id");
		assert.strictEqual(auth["mode"], "view", "auth carries the mode");
		assert.strictEqual(auth["openCmd"]["c"], "open", "auth carries the open command");

		transport.reply({
			"type" : "license",
			"license" : {"type" : 3, "mode" : 0, "rights" : 1, "buildVersion" : "9.3.8", "buildNumber" : 1}
		});
		transport.reply({"type" : "authChanges", "changes" : []});
		transport.reply({
			"type" : "auth",
			"result" : 1,
			"sessionId" : "s1",
			"participants" : [participant("u1", 0, true)],
			"locks" : [],
			"changes" : [],
			"changesIndex" : 0,
			"indexUser" : 0,
			"buildVersion" : "9.3.8",
			"buildNumber" : 1,
			"licenseType" : 3
		});
		transport.reply({
			"type" : "documentOpen",
			"data" : {"type" : "open", "status" : "ok", "data" : {"Editor.bin" : "blob:editor-bin"}}
		});
		assert.notOk(coApi.get_isAuth(), "replies are not handled inside the transport call");
		await tick();
		assert.ok(coApi.get_isAuth(), "api is authorized after the auth reply");
		assert.ok(events.indexOf("license:3") !== -1, "license reply reached onLicense");
		assert.ok(events.indexOf("participants:1:u10") !== -1, "participants reached onAuthParticipantsChanged with the user id + indexUser");
		assert.ok(events.indexOf("documentOpen:blob:editor-bin") !== -1, "documentOpen reached onDocumentOpen");
		assert.deepEqual(events.slice(1), ["license:3", "participants:1:u10", "documentOpen:blob:editor-bin"], "replies are delivered in order");
		assert.strictEqual(transport.sentOfType("authChangesAck").length, 1, "authChanges is acknowledged");

		coApi.disconnect(AscCommon.c_oCloseCode.quiet, "done");
		assert.strictEqual(transport.closed, 1, "disconnect closes the transport");

		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("Messages sent before the connection opens are buffered", async function (assert)
	{
		var transport = createFakeTransport();
		var coApi = createCoApi(transport, {});
		coApi.sendMessage("early");
		assert.strictEqual(transport.sent.length, 0, "a message sent before onOpen is buffered");
		await tick();
		coApi.sendMessage("hello");
		assert.deepEqual(transport.sent, [{"type" : "message", "message" : "hello"}], "state WaitAuth sends through the transport");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("Closing the adapter closes the transport and ignores its later callbacks", async function (assert)
	{
		var transport = createFakeTransport();
		var disconnects = [];
		var messages = [];
		var coApi = createCoApi(transport, {
			onDisconnect : function (reason, code) { disconnects.push(code); },
			onMessage : function (m) { messages.push(m); }
		});
		await tick();
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
		transport.handlers["onDisconnect"]("gone", AscCommon.c_oCloseCode.drop);
		transport.handlers["onMessage"]({"type" : "message", "messages" : ["late"]});
		transport.handlers["onOpen"]();
		await tick();
		assert.strictEqual(coApi._CoAuthoringApi.get_state(), AscCommon.ConnectionState.ClosedAll, "callbacks after an explicit close do not start a reconnect");
		assert.strictEqual(disconnects.length, 1, "no further disconnect is reported");
		assert.strictEqual(transport.connects, 1, "the transport is not reopened");
		assert.strictEqual(coApi._CoAuthoringApi.socketio.queue.length, 0, "nothing is queued");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("Messages queued before an explicit close are dropped", async function (assert)
	{
		var transport = createFakeTransport();
		var opened = 0;
		var coApi = createCoApi(transport, {onFirstConnect : function () { opened++; }});
		coApi.disconnect(AscCommon.c_oCloseCode.quiet, "done");
		await tick();
		assert.strictEqual(opened, 0, "a queued onOpen is not delivered after close");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("The host can drop and reopen its connection", async function (assert)
	{
		var t = await createAuthorizedCoApi({});
		var disconnects = [];
		var opens = 0;
		t.coApi.onDisconnect = function (reason, code) { disconnects.push(code); };
		t.coApi.onFirstConnect = function () { opens++; };

		t.transport.handlers["onClose"](false);
		await tick();
		assert.strictEqual(t.docsCoApi.get_state(), AscCommon.ConnectionState.Reconnect, "an unexpected close reports Reconnect state to the api");
		assert.deepEqual(disconnects, [undefined], "an unexpected close reports a disconnect without a code");

		t.transport.handlers["onOpen"]();
		await tick();
		assert.strictEqual(opens, 1, "onOpen after a drop reaches onFirstConnect again, so the api re-authenticates");
		assert.strictEqual(t.docsCoApi.get_state(), AscCommon.ConnectionState.WaitAuth, "state is WaitAuth again");
		assert.strictEqual(t.transport.connects, 1, "sdkjs did not reconnect on its own");

		t.coApi.disconnect(AscCommon.c_oCloseCode.quiet, "done");
		assert.strictEqual(t.transport.closed, 1, "the reopened connection can be closed");
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

	QUnit.module("Local session editing");

	QUnit.test("Word lock: getLock request and reply keyed by block id", async function (assert)
	{
		var acquired = [];
		var t = await createAuthorizedCoApi({onLocksAcquired : function (lock) { acquired.push(lock); }});
		assert.strictEqual(t.docsCoApi.getUserConnectionId(), "u11", "own user id is user.id + indexUser");

		var results = [];
		var returned = false;
		t.transport.onSend = function (msg)
		{
			if ("getLock" === msg["type"])
			{
				var locks = {};
				msg["block"].forEach(function (b) { locks[b] = {"time" : 1, "user" : "u11", "block" : b}; });
				t.transport.reply({"type" : "getLock", "locks" : locks});
			}
		};
		t.coApi.askLock(["p1", "p2"], function (res)
		{
			results.push({returned : returned, res : res});
		});
		returned = true;
		assert.deepEqual(t.transport.sent, [{"type" : "getLock", "block" : ["p1", "p2"]}], "getLock carries the block ids");
		assert.strictEqual(results.length, 0, "a reply sent inside transport.send is not handled before askLock returns");
		await tick();
		assert.strictEqual(results.length, 1, "the lock callback is called once");
		assert.ok(results[0].returned, "the callback runs after askLock returned");
		assert.strictEqual(results[0].res["lock"]["state"], 2, "the lock is ours");
		assert.strictEqual(acquired.length, 2, "onLocksAcquired is raised for every block");
		assert.strictEqual(t.docsCoApi._lockCallbacksErrorTimerId["p1"], undefined, "the lock timeout is cleared");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("Spreadsheet lock: reply keyed by block.guid, foreign lock refused", async function (assert)
	{
		var t = await createAuthorizedCoApi({}, AscCommon.c_oEditorId.Spreadsheet);
		var block = {"sheetId" : "1", "type" : 1, "rangeOrObjectId" : {"c1" : 0, "r1" : 0, "c2" : 0, "r2" : 0}, "guid" : "g1"};
		var results = [];
		t.coApi.askLock([block], function (res) { results.push(res); });
		assert.deepEqual(t.transport.sent[0]["block"], [block], "getLock carries the block objects");
		t.transport.reply({"type" : "getLock", "locks" : {"g1" : {"time" : 1, "user" : "u11", "block" : block}}});
		await tick();
		assert.ok(results[0]["lock"], "own lock is granted");
		assert.strictEqual(t.docsCoApi.getLocks()["g1"]["state"], 2, "lock is stored under block.guid");

		var block2 = {"sheetId" : "1", "type" : 1, "rangeOrObjectId" : {"c1" : 1, "r1" : 1, "c2" : 1, "r2" : 1}, "guid" : "g2"};
		t.coApi.askLock([block2], function (res) { results.push(res); });
		t.transport.reply({"type" : "getLock", "locks" : {"g2" : {"time" : 1, "user" : "u21", "block" : block2}}});
		await tick();
		assert.strictEqual(results[1]["error"], "Already locked by u21", "a lock of another user is refused");
		assert.strictEqual(t.docsCoApi.getLocks()["g2"]["state"], 3, "the foreign lock is recorded");

		t.transport.reply({"type" : "releaseLock", "locks" : [{"block" : block2, "user" : "u21", "time" : 2, "changes" : null}]});
		await tick();
		assert.strictEqual(t.docsCoApi.getLocks()["g2"]["state"], 0, "releaseLock frees the block");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("Lock -> edit -> chunked save -> unSaveLock", async function (assert)
	{
		var unSaveLocks = 0;
		var t = await createAuthorizedCoApi({onUnSaveLock : function () { unSaveLocks++; }});
		var docsCoApi = t.docsCoApi;
		var session = {changes : [], index : 0};
		t.transport.onSend = function (msg)
		{
			switch (msg["type"])
			{
				case "getLock":
					var locks = {};
					msg["block"].forEach(function (b) { locks[b] = {"time" : 1, "user" : "u11", "block" : b}; });
					t.transport.reply({"type" : "getLock", "locks" : locks});
					break;
				case "isSaveLock":
					t.transport.reply({"type" : "saveLock", "saveLock" : false});
					break;
				case "saveChanges":
					var startIndex = session.index;
					var changes = JSON.parse(msg["changes"]);
					session.changes = session.changes.concat(changes);
					session.index += changes.length;
					var changesIndex = msg["startSaveChanges"] ? startIndex : -1;
					if (msg["endSaveChanges"])
						t.transport.reply({"type" : "unSaveLock", "index" : changesIndex, "time" : 1000, "syncChangesIndex" : session.index});
					else
						t.transport.reply({"type" : "savePartChanges", "changesIndex" : changesIndex, "syncChangesIndex" : session.index});
					break;
			}
		};

		var lockResult = null;
		t.coApi.askLock(["p1"], function (res) { lockResult = res; });
		await tick();
		assert.ok(lockResult && lockResult["lock"], "lock acquired");

		var saveLock = null;
		t.coApi.askSaveChanges(function (e) { saveLock = e; });
		assert.strictEqual(docsCoApi.get_state(), AscCommon.ConnectionState.AskSaveChanges, "state is AskSaveChanges while waiting");
		assert.deepEqual(t.transport.sentOfType("isSaveLock"), [{"type" : "isSaveLock", "syncChangesIndex" : 0}], "isSaveLock carries syncChangesIndex");

		var buffered = null;
		t.coApi.askLock(["p2"], function (res) { buffered = res; });
		assert.strictEqual(t.transport.sentOfType("getLock").length, 1, "a lock asked during the save is buffered");

		await tick();
		assert.deepEqual(saveLock, {"type" : "saveLock", "saveLock" : false}, "saveLock reply reaches the callback");
		assert.strictEqual(docsCoApi.get_state(), AscCommon.ConnectionState.AskSaveChanges, "state stays until the changes are sent");

		docsCoApi.websocketMaxPayloadSize = 10;
		var edits = ["change-1-xx", "change-2-xx", "c3"];
		t.coApi.saveChanges(edits, null, null, false, false);
		assert.strictEqual(docsCoApi.get_state(), AscCommon.ConnectionState.SaveChanges, "state is SaveChanges");
		await tick();
		await tick();
		await tick();

		var saves = t.transport.sentOfType("saveChanges");
		assert.strictEqual(saves.length, 3, "the changes are sent in three chunks");
		assert.deepEqual(saves.map(function (m) { return [m["startSaveChanges"], m["endSaveChanges"]]; }), [[true, false], [false, false], [false, true]], "start/end flags mark the chunks");
		assert.deepEqual(session.changes, edits, "all changes arrive in order");
		assert.strictEqual(typeof saves[0]["changes"], "string", "changes are a JSON string without binaryChanges");
		assert.strictEqual(saves[0]["isExcel"], false, "isExcel flag");
		assert.strictEqual(saves[0]["isCoAuthoring"], false, "isCoAuthoring flag");
		assert.strictEqual(saves[0]["unlock"], false, "unlock flag");
		assert.strictEqual(saves[0]["releaseLocks"], false, "releaseLocks flag");
		assert.strictEqual(unSaveLocks, 1, "onUnSaveLock raised once");
		assert.strictEqual(docsCoApi.get_state(), AscCommon.ConnectionState.Authorized, "state back to Authorized");
		assert.strictEqual(docsCoApi.changesIndex, 0, "changesIndex from the first chunk");
		assert.strictEqual(docsCoApi.syncChangesIndex, 3, "syncChangesIndex from unSaveLock");
		assert.strictEqual(docsCoApi.lastOwnSaveTime, 1000, "own save time from unSaveLock");
		assert.strictEqual(docsCoApi.saveCallbackErrorTimeOutId, null, "the save timeout is cleared");
		assert.strictEqual(docsCoApi.getLocks()["p1"], undefined, "own locks are dropped with the last chunk");
		assert.ok(buffered && buffered["lock"], "the buffered lock is asked and granted after the save");

		t.transport.sent = [];
		t.coApi.askSaveChanges(function () {});
		assert.deepEqual(t.transport.sent[0], {"type" : "isSaveLock", "syncChangesIndex" : 3}, "the next save sends the new syncChangesIndex");
		await tick();
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("Emergency unSaveLock is answered and its retry timer cleared", async function (assert)
	{
		var unSaveLocks = 0;
		var t = await createAuthorizedCoApi({onUnSaveLock : function () { unSaveLocks++; }});
		t.transport.onSend = function (msg)
		{
			if ("unSaveLock" === msg["type"])
				t.transport.reply({"type" : "unSaveLock", "index" : -1, "time" : -1, "syncChangesIndex" : -1});
		};
		t.coApi.unSaveLock();
		assert.deepEqual(t.transport.sent, [{"type" : "unSaveLock"}], "unSaveLock request");
		await tick();
		assert.strictEqual(unSaveLocks, 1, "onUnSaveLock raised");
		assert.strictEqual(t.docsCoApi.unSaveLockCallbackErrorTimeOutId, null, "retry timer cleared");
		assert.strictEqual(t.docsCoApi.lastOwnSaveTime, -1, "time -1 leaves the save time");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("Incoming saveChanges of another participant, in chunks", async function (assert)
	{
		var applied = [];
		var released = [];
		var indexes = [];
		var recalc = 0;
		var t = await createAuthorizedCoApi({
			onSaveChanges : function (change, userId) { applied.push([change, userId]); },
			onLocksReleased : function (lock, bChanges) { released.push([lock["block"], bChanges]); },
			onChangesIndex : function (index) { indexes.push(index); },
			onRecalcLocks : function () { recalc++; }
		});
		var entry = function (change) { return {"docid" : "doc1", "change" : JSON.stringify(change), "time" : 2000, "user" : "u21", "useridoriginal" : "u2"}; };
		t.transport.reply({"type" : "saveChanges", "changes" : [entry("r1")], "changesIndex" : 1, "syncChangesIndex" : 1, "endSaveChanges" : false, "locks" : []});
		await tick();
		assert.strictEqual(applied.length, 0, "a partial chunk is buffered");
		t.transport.reply({
			"type" : "saveChanges", "changes" : [entry("r2")], "changesIndex" : 2, "syncChangesIndex" : 2, "endSaveChanges" : true,
			"locks" : [{"block" : "p9", "user" : "u21", "time" : 2000, "changes" : null}], "excelAdditionalInfo" : null
		});
		await tick();
		assert.deepEqual(applied, [["r1", "u2"], ["r2", "u2"]], "both chunks are applied in order with the original user id");
		assert.deepEqual(released, [["p9", true]], "the sender's locks are released with changes");
		assert.deepEqual(indexes, [2], "changes index reported");
		assert.strictEqual(t.docsCoApi.syncChangesIndex, 2, "syncChangesIndex updated");
		assert.strictEqual(t.docsCoApi.lastOtherSaveTime, 2000, "lastOtherSaveTime from the change time");
		assert.strictEqual(recalc, 1, "onRecalcLocks raised");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("connectState from the host starts and ends co-authoring", async function (assert)
	{
		var events = [];
		var t = await createAuthorizedCoApi({
			onStartCoAuthoring : function (isStart, isWaitAuth) { events.push("start:" + isStart + ":" + !!isWaitAuth); },
			onEndCoAuthoring : function (isStart) { events.push("end:" + isStart); },
			onParticipantsChanged : function (users) { events.push("participants:" + Object.keys(users).length); },
			onConnectionStateChanged : function (user) { events.push("user:" + user.asc_getId() + ":" + user.asc_getState()); }
		});
		t.transport.reply({"type" : "connectState", "participantsTimestamp" : 10, "participants" : [participant("u1", 1), participant("u2", 1)], "waitAuth" : false});
		await tick();
		assert.deepEqual(events, ["start:false:false", "participants:2", "user:u21:true"], "a second editor starts co-authoring");

		events = [];
		t.transport.reply({"type" : "connectState", "participantsTimestamp" : 9, "participants" : [participant("u1", 1)], "waitAuth" : false});
		await tick();
		assert.deepEqual(events, [], "an older participantsTimestamp is ignored");

		t.transport.reply({"type" : "connectState", "participantsTimestamp" : 11, "participants" : [participant("u1", 1)], "waitAuth" : false});
		await tick();
		assert.deepEqual(events, ["end:false", "participants:1", "user:u21:false"], "the editor leaving ends co-authoring");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("Forcesave in a local session does not wait for the host", async function (assert)
	{
		var forceSaves = [];
		var t = await createAuthorizedCoApi({onForceSave : function (e) { forceSaves.push(e); }});
		assert.strictEqual(t.coApi.forceSave(), false, "forceSave does not ask the editor to wait");
		assert.deepEqual(t.transport.sent, [{"type" : "forceSaveStart"}], "forceSaveStart is still sent to the host");
		assert.strictEqual(t.coApi.forceSave(), false, "a second forcesave without a new save");
		assert.strictEqual(t.transport.sent.length, 1, "is not sent again");

		var Button = AscCommon.c_oAscForceSaveTypes.Button;
		t.transport.reply({"type" : "forceSaveStart", "messages" : {"code" : AscCommon.c_oAscServerCommandErrors.NoError, "time" : 5}});
		t.transport.reply({"type" : "forceSave", "messages" : {"type" : Button, "time" : 5, "success" : true}});
		await tick();
		assert.deepEqual(forceSaves, [{type : Button, start : true}, {type : Button, success : true}], "host replies are still reported");

		var api = new AscCommon.baseEditorsApi({});
		api.CoAuthoringApi = t.coApi;
		api.getViewMode = function () { return false; };
		t.docsCoApi.lastOwnSaveTime = 6;
		assert.strictEqual(api.forceSave(), false, "api.forceSave returns false so the Save action is ended on unSaveLock");
		assert.strictEqual(t.transport.sentOfType("forceSaveStart").length, 2, "after a new save forceSaveStart is sent again");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("Forcesave over socket.io still waits for the server", function (assert)
	{
		var docsCoApi = new AscCommon.CDocsCoApi()._CoAuthoringApi;
		var sent = [];
		docsCoApi._state = AscCommon.ConnectionState.Authorized;
		docsCoApi.socketio = {emit : function (name, data) { sent.push(data); }};
		assert.strictEqual(docsCoApi.forceSave(), true, "forceSave returns true");
		assert.deepEqual(sent, [{"type" : "forceSaveStart"}], "forceSaveStart sent");
	});

	QUnit.test("rpc requests go through the transport", async function (assert)
	{
		var t = await createAuthorizedCoApi({});
		var results = [];
		t.transport.onSend = function (msg)
		{
			if ("rpc" === msg["type"])
				t.transport.reply({"type" : "rpc", "responseKey" : msg["responseKey"], "data" : {"status" : "ok", "data" : ["blob:img9"]}});
		};
		t.coApi.callPRC({"type" : "pathurls", "data" : ["media/image9.png"]}, 1000, function (isTimeout, res) { results.push([isTimeout, res]); });
		assert.deepEqual(t.transport.sent[0]["data"], {"type" : "pathurls", "data" : ["media/image9.png"]}, "rpc carries the request");
		await tick();
		assert.deepEqual(results, [[false, {"status" : "ok", "data" : ["blob:img9"]}]], "the reply reaches the callback");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("imgurls (paste, insert by url) go through the transport", async function (assert)
	{
		var t = await createAuthorizedCoApi({});
		var api = new AscCommon.baseEditorsApi({});
		api.documentId = "doc1";
		api.CoAuthoringApi = t.coApi;
		api.isOpenedFrameEditor = true;
		api.frameManager = {isFrameEditor : function () { return false; }};
		var result = null;
		AscCommon.sendImgUrls(api, ["data:image/png;base64,AAAA"], function (data) { result = data; });
		var msg = t.transport.sent[0];
		assert.strictEqual(msg["type"], "openDocument", "imgurls is sent as openDocument");
		assert.strictEqual(msg["message"]["c"], "imgurls", "command imgurls");
		assert.deepEqual(msg["message"]["data"], ["data:image/png;base64,AAAA"], "with the image sources");

		api.fCurCallback({"type" : "imgurls", "status" : "ok", "data" : {"urls" : [{"path" : "media/image5.png", "url" : "blob:img5"}], "error" : 0}});
		assert.deepEqual(result, [{"path" : "media/image5.png", "url" : "blob:img5"}], "the documentOpen reply shape reaches the caller");
		assert.strictEqual(AscCommon.g_oDocumentUrls.getImageLocal("blob:img5"), "image5.png", "the url is mapped to its media path");
		AscCommon.setLocalSessionTransport(null);
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

	QUnit.test("Image upload with a media map registers the media path", function (assert)
	{
		AscCommon.setLocalSessionTransport({
			uploadImages : function (files, cb)
			{
				cb(Asc.c_oAscError.ID.No, {"media/image7.png" : "blob:img7", "media/image8.png" : "blob:img8"});
			}
		});
		var result = null;
		AscCommon.UploadImageFiles([{name : "a.png"}, {name : "b.png"}], "doc1", "u1", "", undefined, undefined, undefined, function (error, urls)
		{
			result = {error : error, urls : urls};
		});
		assert.deepEqual(result, {error : Asc.c_oAscError.ID.No, urls : ["blob:img7", "blob:img8"]}, "the urls are forwarded in order");
		assert.strictEqual(AscCommon.g_oDocumentUrls.getImageLocal("blob:img7"), "image7.png", "the url is saved as its media path");
		assert.strictEqual(AscCommon.g_oDocumentUrls.getImageUrl("image8.png"), "blob:img8", "the media path resolves to the url");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("asc_addImage upload path inserts the transport urls", function (assert)
	{
		var uploaded = null;
		AscCommon.setLocalSessionTransport({
			uploadImages : function (files, cb)
			{
				uploaded = files;
				cb(Asc.c_oAscError.ID.No, {"media/image3.png" : "blob:img3"});
			}
		});
		var api = createApi();
		api.CoAuthoringApi = new AscCommon.CDocsCoApi();
		var actions = [];
		var inserted = null;
		api.sync_StartAction = function (type, id) { actions.push("start:" + id); };
		api.sync_EndAction = function (type, id) { actions.push("end:" + id); };
		api.sendEvent = function () {};
		api._addImageUrl = function (urls) { inserted = urls; };
		var files = [{name : "c.png"}];
		api._uploadCallback(Asc.c_oAscError.ID.No, files, {});
		assert.strictEqual(uploaded, files, "the chosen files reach transport.uploadImages");
		assert.deepEqual(inserted, ["blob:img3"], "the image is inserted with the transport url");
		assert.deepEqual(actions, ["start:" + Asc.c_oAscAsyncAction.UploadImage, "end:" + Asc.c_oAscAsyncAction.UploadImage], "the upload action is ended");
		AscCommon.setLocalSessionTransport(null);
	});

	QUnit.test("Image upload by url goes through the transport", function (assert)
	{
		var requested = null;
		AscCommon.setLocalSessionTransport({
			uploadImageUrls : function (urls, cb)
			{
				requested = urls;
				cb(Asc.c_oAscError.ID.No, [{"path" : "media/image11.png", "url" : "blob:img11"}]);
			}
		});
		var result = null;
		AscCommon.UploadImageUrls(["http://example.com/a.png"], "doc1", "u1", "", undefined, undefined, undefined, function (data) { result = data; });
		assert.strictEqual(Asc.c_oAscError.ID.No, 0, "success code is 0");
		assert.deepEqual(requested, ["http://example.com/a.png"], "transport.uploadImageUrls receives the urls");
		assert.deepEqual(result, [{"path" : "media/image11.png", "url" : "blob:img11"}], "the DS shaped reply is forwarded");
		assert.strictEqual(AscCommon.g_oDocumentUrls.getImageLocal("blob:img11"), "image11.png", "the url is mapped to its media path");

		AscCommon.setLocalSessionTransport({
			uploadImageUrls : function (urls, cb) { cb(Asc.c_oAscError.ID.UplImageUrl); }
		});
		AscCommon.UploadImageUrls(["http://example.com/b.png"], "doc1", "u1", "", undefined, undefined, undefined, function (data) { result = data; });
		assert.deepEqual(result, [], "a transport error yields no urls, like a failed DS upload");
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
