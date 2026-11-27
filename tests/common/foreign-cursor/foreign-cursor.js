/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

// The cursors of co-authors that come before the own editor can use them: before the own auth (DocsCoApi), and
// before the document is ready (baseEditorsApi.deferCursorUntilReady, applyDeferredCursors).
$(function ()
{
	// A cursor message as the server sends it: a list, the last entry counts
	function cursor(user, place)
	{
		return [{cursor : place, user : user, useridoriginal : user, time : 0}];
	}

	// An editor API with the real methods of the base, and a co-editing API of its own
	function createApi(participants)
	{
		let api = Object.create(AscCommon.baseEditorsApi.prototype);
		api.isDocumentLoadComplete = false;
		api.preReadyCursors = {};
		api.applied = [];
		api.logged = [];
		api.CoAuthoringApi = {
			isParticipant : function (userId) { return -1 !== participants.indexOf(userId); },
			sendClientLog : function (level, msg) { api.logged.push(level + ": " + msg); },
			onCursor      : function (e)
			{
				let last = e[e.length - 1];
				if ("throw" === last.cursor)
					throw new Error("bad cursor");
				api.applied.push(last.user + ":" + last.cursor);
			}
		};
		return api;
	}

	QUnit.module("Cursors of co-authors before the document is ready");

	QUnit.test("A cursor waits until the document is ready, and only the newest of each user is applied", function (assert)
	{
		let api = createApi(["A", "B"]);

		assert.strictEqual(api.deferCursorUntilReady(cursor("A", "1")), true, "the first cursor of A waits");
		assert.strictEqual(api.deferCursorUntilReady(cursor("B", "1")), true, "the cursor of B waits");
		assert.strictEqual(api.deferCursorUntilReady(cursor("A", "2")), true, "the second cursor of A waits");
		assert.deepEqual(api.applied, [], "nothing is applied yet");

		api.isDocumentLoadComplete = true;
		api.applyDeferredCursors();

		assert.deepEqual(api.applied.slice().sort(), ["A:2", "B:1"], "the newest cursor of each user is applied");
		assert.deepEqual(Object.keys(api.preReadyCursors), [], "nothing waits any more");
	});

	QUnit.test("An empty cursor message does not wait and does not fail", function (assert)
	{
		let api = createApi(["A"]);

		assert.strictEqual(api.deferCursorUntilReady([]), false, "an empty message does not wait");
		assert.strictEqual(api.deferCursorUntilReady(cursor("A", "1")), true, "a cursor waits");
		assert.deepEqual(Object.keys(api.preReadyCursors), ["A"], "only the cursor waits");
	});

	QUnit.test("A cursor that comes when the document is ready does not wait", function (assert)
	{
		let api = createApi(["A"]);
		api.isDocumentLoadComplete = true;

		assert.strictEqual(api.deferCursorUntilReady(cursor("A", "1")), false, "the cursor does not wait");
		assert.deepEqual(Object.keys(api.preReadyCursors), [], "nothing waits");
	});

	QUnit.test("The cursor of a user who left before the document was ready is not applied", function (assert)
	{
		let api = createApi(["B"]);
		api.deferCursorUntilReady(cursor("A", "1"));
		api.deferCursorUntilReady(cursor("B", "1"));

		api.isDocumentLoadComplete = true;
		api.applyDeferredCursors();

		assert.deepEqual(api.applied, ["B:1"], "only the cursor of the user who is there is applied");
	});

	QUnit.test("A cursor that throws does not stop the others", function (assert)
	{
		let api = createApi(["A", "B"]);
		api.deferCursorUntilReady(cursor("A", "throw"));
		api.deferCursorUntilReady(cursor("B", "1"));

		api.isDocumentLoadComplete = true;
		api.applyDeferredCursors();

		assert.deepEqual(api.applied, ["B:1"], "the cursor of B is applied");
		assert.strictEqual(api.logged.length, 1, "the error is logged");
		assert.deepEqual(Object.keys(api.preReadyCursors), [], "nothing waits any more");
	});

	QUnit.module("Cursors of co-authors before the own auth");

	// The connection of CDocsCoApi, before and after its auth
	function createConnection()
	{
		let connection = new AscCommon.CDocsCoApi()._CoAuthoringApi;
		connection.applied = [];
		connection.onCursor = function (e) { connection.applied.push(e[e.length - 1].user); };
		return connection;
	}

	QUnit.test("A cursor before the own auth is applied after it, unless its user left", function (assert)
	{
		let connection = createConnection();
		connection._state = AscCommon.ConnectionState.WaitAuth;

		connection._onCursor({messages : cursor("A", "1")});
		connection._onCursor({messages : cursor("B", "1")});
		assert.deepEqual(connection.applied, [], "nothing is applied before the auth");

		// The auth: B is there, A left in the meantime
		connection._participants = {"B" : {}};
		connection._state = AscCommon.ConnectionState.Authorized;
		connection._applyPrebuffered();

		assert.deepEqual(connection.applied, ["B"], "the cursor of B is applied, the one of A is not");
	});

	QUnit.test("Before the connection is made, no user is a participant", function (assert)
	{
		let coApi = new AscCommon.CDocsCoApi();
		coApi._CoAuthoringApi._participants = {"A" : {}};

		assert.strictEqual(coApi.isParticipant("A"), false, "not online: no participant");
		coApi._onlineWork = true;
		assert.strictEqual(coApi.isParticipant("A"), true, "online: A is a participant");
		assert.strictEqual(coApi.isParticipant("B"), false, "online: B is not");
	});

	QUnit.test("A cursor after the own auth is applied at once", function (assert)
	{
		let connection = createConnection();
		connection._state = AscCommon.ConnectionState.Authorized;
		connection._participants = {"A" : {}};

		connection._onCursor({messages : cursor("A", "1")});

		assert.deepEqual(connection.applied, ["A"], "the cursor is applied");
	});
});
