/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

// The loads of fonts and images: a paste (pre_Paste) and the images of the changes of others (pre_Save) share the
// font loader and the image loader, and asyncFontsDocumentEndLoaded / asyncImagesDocumentEndLoaded end them. The
// page sets window.pasteQueueEditors to the names of the editor APIs to test.
$(function ()
{
	// An editor API object with the real methods. Its font loader loads the fonts whose name starts with "Load",
	// and records the loads; its image loader records the image maps. The end of a load is a call of
	// asyncFontsDocumentEndLoaded, then of asyncImagesDocumentEndLoaded.
	function createApi(Api)
	{
		let api = Object.create(Api.prototype);
		api.isPasteFonts_Images = false;
		api.isSaveFonts_Images = false;
		api.isLoadImagesCustom = false;
		api.pasteCallback = null;
		api.pasteImageMap = null;
		api.saveImageMap = null;
		api.pasteQueue = [];
		api.isSaveWaitingForPaste = false;
		api.IsLongActionCurrent = 0;
		api.groupActionsCounter = 0;
		api.LongActionCallbacks = [];
		api.LongActionCallbacksParams = [];
		api.EndActionLoadImages = 0;
		api.bInit_word_control = true;
		api.editable = true;
		api.canEdit = function () { return this.editable; };
		api.loads = [];
		api.imageLoads = [];
		api.saves = 0;
		api.FontLoader = {
			CheckFontsNeedLoading : function (fonts)
			{
				return fonts.some(function (font) { return 0 === font.name.indexOf("Load"); });
			},
			LoadDocumentFonts2    : function (fonts) { api.loads.push(fonts.map(function (font) { return font.name; })); }
		};
		api.ImageLoader = {
			bIsLoadDocumentFirst       : false,
			bIsAsyncLoadDocumentImages : true,
			LoadDocumentImages         : function (images) { api.imageLoads.push(images); }
		};
		api.WordControl = {
			m_oDrawingDocument : {m_oDocumentRenderer : null, CheckFontNeeds : function () {}},
			m_oLogicDocument   : {Fonts : fonts("LoadSave")}
		};
		api.getPDFDoc = function () { return {drawings : []}; };
		api.pre_SaveCallback = function () { this.saves++; };
		api.sync_StartAction = function () {};
		api.updateSelection = function () {};
		api.sync_EndAction = function () {};
		return api;
	}

	function fonts()
	{
		return Array.prototype.map.call(arguments, function (name) { return {name : name}; });
	}

	// Runs fn with a global lock of the co-editing of its own, held count times, and gives the lock to fn
	function withGlobalLock(count, fn)
	{
		let lock = {
			count                   : count,
			Get_GlobalLock          : function () { return 0 !== this.count; },
			Set_GlobalLock          : function (isLock) { this.count += isLock ? 1 : -1; },
			Set_GlobalLockSelection : function () {}
		};
		let old = AscCommon.CollaborativeEditing;
		AscCommon.CollaborativeEditing = lock;
		try
		{
			fn(lock);
		}
		finally
		{
			AscCommon.CollaborativeEditing = old;
		}
	}

	// The end of a load: the fonts, then the images
	function endLoad(api)
	{
		api.asyncFontsDocumentEndLoaded();
		api.asyncImagesDocumentEndLoaded();
	}

	(window.pasteQueueEditors || []).forEach(function (name)
	{
		let Api = (window.Asc && window.Asc[name]) || (window.AscCommon && window.AscCommon[name]);

		QUnit.module("pre_Paste of " + name);

		QUnit.test("A paste with nothing to load is done at once", function (assert)
		{
			let api = createApi(Api);
			let done = [];

			api.pre_Paste(fonts(), {}, function () { done.push("A"); });

			assert.deepEqual(done, ["A"], "the paste is done");
			assert.deepEqual(api.loads, [], "nothing is loaded");
			assert.strictEqual(api.isPasteFonts_Images, false, "no load runs");
		});

		QUnit.test("A paste without a load, during the load of another paste, does not break it", function (assert)
		{
			let api = createApi(Api);
			let done = [];

			api.pre_Paste(fonts("LoadA"), {}, function () { done.push("A"); });
			api.pre_Paste(fonts(), {}, function () { done.push("B"); });
			assert.deepEqual(done, [], "B waits for A");

			endLoad(api);

			assert.deepEqual(done, ["A", "B"], "both pastes are done, in their order");
			assert.strictEqual(api.isPasteFonts_Images, false, "no load runs");
			assert.strictEqual(api.IsLongActionCurrent, 0, "the long action ended");
		});

		QUnit.test("The images of the paste that loads are loaded, then those of the paste that waits", function (assert)
		{
			let api = createApi(Api);
			let done = [];
			let imagesA = {"a.png" : "a.png"};
			let imagesB = {"b.png" : "b.png"};

			api.pre_Paste(fonts("LoadA"), imagesA, function () { done.push("A"); });
			api.pre_Paste(fonts("LoadB"), imagesB, function () { done.push("B"); });
			assert.deepEqual(api.loads, [["LoadA"]], "only the fonts of A load");

			api.asyncFontsDocumentEndLoaded();
			assert.deepEqual(api.imageLoads, [imagesA], "the images of A load");
			api.asyncImagesDocumentEndLoaded();
			assert.deepEqual(done, ["A"], "A is done");
			assert.deepEqual(api.loads, [["LoadA"], ["LoadB"]], "then the fonts of B load");

			endLoad(api);
			assert.deepEqual(api.imageLoads, [imagesA, imagesB], "then the images of B");
			assert.deepEqual(done, ["A", "B"], "B is done");
			assert.strictEqual(api.IsLongActionCurrent, 0, "the long action ended");
		});

		QUnit.test("Two pastes that wait are done in their order, the second after the load of the first", function (assert)
		{
			let api = createApi(Api);
			let done = [];

			api.pre_Paste(fonts("LoadA"), {}, function () { done.push("A"); });
			api.pre_Paste(fonts("LoadB"), {}, function () { done.push("B"); });
			api.pre_Paste(fonts(), {}, function () { done.push("C"); });

			endLoad(api);
			assert.deepEqual(done, ["A"], "C waits for the load of B");
			assert.strictEqual(api.pasteQueue.length, 1, "C waits");

			endLoad(api);
			assert.deepEqual(done, ["A", "B", "C"], "B and C are done");
			assert.strictEqual(api.pasteQueue.length, 0, "nothing waits");
			assert.strictEqual(api.IsLongActionCurrent, 0, "the long action ended");
		});

		QUnit.test("A paste that the callback of a paste without a load starts is kept", function (assert)
		{
			let api = createApi(Api);
			let done = [];

			api.pre_Paste(fonts(), {}, function ()
			{
				done.push("A");
				api.pre_Paste(fonts("LoadB"), {}, function () { done.push("B"); });
			});
			assert.deepEqual(api.loads, [["LoadB"]], "B loads");

			endLoad(api);
			assert.deepEqual(done, ["A", "B"], "B is done");
			assert.strictEqual(api.isPasteFonts_Images, false, "no load runs");
			assert.strictEqual(api.IsLongActionCurrent, 0, "the long action ended");
		});

		QUnit.test("A paste that the callback of a paste starts is kept", function (assert)
		{
			let api = createApi(Api);
			let done = [];

			api.pre_Paste(fonts("LoadA"), {}, function ()
			{
				done.push("A");
				api.pre_Paste(fonts("LoadB"), {}, function () { done.push("B"); });
			});
			endLoad(api);
			assert.deepEqual(api.loads, [["LoadA"], ["LoadB"]], "B loads");

			endLoad(api);
			assert.deepEqual(done, ["A", "B"], "B is done");
			assert.strictEqual(api.isPasteFonts_Images, false, "no load runs");
			assert.strictEqual(api.IsLongActionCurrent, 0, "the long action ended");
		});

		QUnit.test("A callback that throws ends the long action and drops the pastes that wait", function (assert)
		{
			let api = createApi(Api);
			let done = [];

			api.pre_Paste(fonts("LoadA"), {}, function () { throw new Error("A"); });
			api.pre_Paste(fonts(), {}, function () { done.push("B"); });

			api.asyncFontsDocumentEndLoaded();
			assert.throws(function () { api.asyncImagesDocumentEndLoaded(); }, /A/, "the error of A comes through");

			assert.strictEqual(api.IsLongActionCurrent, 0, "the long action ended");
			assert.strictEqual(api.pasteQueue.length, 0, "B is dropped");
			assert.deepEqual(done, [], "B is not done");
		});

		QUnit.test("A paste waits for the load of the images of the changes of others", function (assert)
		{
			let api = createApi(Api);
			let done = [];

			// CollaborativeEditing.SendImagesUrlsFromChanges: the images of the changes of others are to load
			api.isSaveFonts_Images = true;
			api.pre_Paste(fonts("LoadA"), {}, function () { done.push("A"); });
			assert.deepEqual(api.loads, [], "A does not load");

			api.pre_Save({});
			endLoad(api);
			assert.strictEqual(api.saves, 1, "the changes of others are done");
			assert.strictEqual(api.isPasteFonts_Images, true, "then A loads");

			endLoad(api);
			assert.deepEqual(done, ["A"], "A is done");
			assert.strictEqual(api.IsLongActionCurrent, 0, "the long action ended");
		});

		QUnit.test("The load of the images of the changes of others waits for the load of a paste", function (assert)
		{
			let api = createApi(Api);
			let done = [];

			api.pre_Paste(fonts("LoadA"), {}, function () { done.push("A"); });
			api.pre_Save({"c.png" : "c.png"});
			assert.strictEqual(api.loads.length, 1, "the save does not load yet");

			endLoad(api);
			assert.deepEqual(done, ["A"], "A is done");
			assert.strictEqual(api.loads.length, 2, "then the save loads");
			assert.strictEqual(api.saves, 0, "the save is not done yet");

			endLoad(api);
			assert.strictEqual(api.saves, 1, "the save is done");
			assert.strictEqual(api.isSaveFonts_Images, false, "no load runs");
			assert.strictEqual(api.IsLongActionCurrent, 0, "the long action ended");
		});

		QUnit.test("A paste of the clipboard data waits for the load of a paste", function (assert)
		{
			let api = createApi(Api);
			let calls = [];

			api.pre_Paste(fonts("LoadA"), {}, function () {});
			api.asc_PasteData(1, "text");
			api.asc_PasteData(3, "more");
			assert.strictEqual(api.pasteQueue.length, 2, "both calls wait, and change nothing yet");

			// The calls that wait are the calls of the API method, with their arguments
			api.asc_PasteData = function () { calls.push(Array.prototype.slice.call(arguments)); };
			endLoad(api);

			assert.deepEqual(calls, [[1, "text"], [3, "more"]], "they run after the load, in their order");
		});

		QUnit.test("A special paste is refused during a load", function (assert)
		{
			// A special paste undoes the last paste and pastes its data again: after a paste that waited, that
			// would be the wrong paste
			let api = createApi(Api);
			let done = [];

			api.pre_Paste(fonts("LoadA"), {}, function () { done.push("A"); });
			api.asc_SpecialPasteData({id : 2}, false);
			assert.strictEqual(api.pasteQueue.length, 0, "the special paste does not wait");

			let special = 0;
			api.asc_SpecialPasteData = function () { special++; };
			endLoad(api);

			assert.deepEqual(done, ["A"], "the paste is done");
			assert.strictEqual(special, 0, "the special paste does not run later");
		});

		QUnit.test("A paste of the shared element of the clipboard does not wait as a whole", function (assert)
		{
			let api = createApi(Api);
			let clipboard = AscCommon.g_clipboardBase;
			let oldDiv = clipboard.CommonDiv;
			let div = {};
			clipboard.CommonDiv = div;
			api.editable = false; // the call then ends early in each editor, with no document to paste to
			try
			{
				withGlobalLock(1, function ()
				{
					api.pre_Paste(fonts("LoadA"), {}, function () {});
					api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, div);
					assert.strictEqual(api.pasteQueue.length, 0, "the paste of the clipboard element does not wait");
					api.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, {});
					assert.strictEqual(api.pasteQueue.length, 1, "the paste of another element waits");
				});
			}
			finally
			{
				clipboard.CommonDiv = oldDiv;
			}
		});

		QUnit.test("The pastes that wait run also when the document can not be edited any more", function (assert)
		{
			// The caller of a paste can wait for its callback: a plugin method, or an action that pre_Paste ends
			let api = createApi(Api);
			let done = [];

			api.pre_Paste(fonts("LoadA"), {}, function () { done.push("A"); });
			api.pre_Paste(fonts(), {}, function () { done.push("B"); });
			api.editable = false;

			endLoad(api);

			assert.deepEqual(done, ["A", "B"], "B is done");
			assert.strictEqual(api.pasteQueue.length, 0, "nothing waits");
		});

		QUnit.test("A paste that waits and throws drops the pastes after it", function (assert)
		{
			let api = createApi(Api);
			let done = [];

			api.pre_Paste(fonts("LoadA"), {}, function () { done.push("A"); });
			api.pre_Paste(fonts(), {}, function () { throw new Error("B"); });
			api.pre_Paste(fonts(), {}, function () { done.push("C"); });

			api.asyncFontsDocumentEndLoaded();
			assert.throws(function () { api.asyncImagesDocumentEndLoaded(); }, /B/, "the error of B comes through");

			assert.deepEqual(done, ["A"], "C is not done");
			assert.strictEqual(api.pasteQueue.length, 0, "C is dropped");
			assert.strictEqual(api.IsLongActionCurrent, 0, "the long action ended");
		});

		QUnit.test("A paste that a plugin makes in its group actions runs with the global lock released", function (assert)
		{
			let api = createApi(Api);
			let seen = [];

			withGlobalLock(1, function (lock)
			{
				api.groupActionsCounter = 1; // a plugin runs group actions: they hold the global lock
				api.pre_Paste(fonts("LoadA"), {}, function () {});
				api.executeGroupActions(function ()
				{
					api.pre_Paste(fonts(), {}, function () { seen.push(lock.Get_GlobalLock()); });
				});
				endLoad(api);
				assert.strictEqual(lock.count, 1, "the group actions hold the lock again");
			});

			assert.deepEqual(seen, [false], "B runs with the lock released");
		});

		QUnit.test("A paste that waits, then runs in group actions, runs with the global lock released", function (assert)
		{
			let api = createApi(Api);
			let seen = [];

			withGlobalLock(0, function (lock)
			{
				api.pre_Paste(fonts("LoadA"), {}, function () {});
				api.pre_Paste(fonts(), {}, function () { seen.push(lock.Get_GlobalLock()); });

				api.groupActionsCounter = 1; // then a plugin starts group actions
				lock.Set_GlobalLock(true);
				endLoad(api);
			});

			assert.deepEqual(seen, [false], "B runs with the lock released, as when it was made");
		});

		QUnit.test("A paste made during the global lock does not get the lock released", function (assert)
		{
			let api = createApi(Api);
			let seen = [];

			withGlobalLock(1, function (lock)
			{
				api.groupActionsCounter = 1; // group actions hold the lock, the paste does not come from them
				api.pre_Paste(fonts("LoadA"), {}, function () {});
				api.pre_Paste(fonts(), {}, function () { seen.push(lock.Get_GlobalLock()); });
				endLoad(api);
			});

			assert.deepEqual(seen, [true], "B runs with the lock held");
		});

		QUnit.test("The pastes that wait start at the end of SyncLoadImages", function (assert)
		{
			if (-1 === (window.pasteQueueCustomLoadEditors || []).indexOf(name))
			{
				assert.ok(true, "this editor has no SyncLoadImages end");
				return;
			}
			let api = createApi(Api);
			let done = [];

			api.isLoadImagesCustom = true;
			api.pre_Paste(fonts(), {}, function () { done.push("B"); });
			assert.deepEqual(done, [], "B waits");
			api.asyncImagesDocumentEndLoaded();

			assert.deepEqual(done, ["B"], "B is done");
		});

		QUnit.test("The shared element of the clipboard is known", function (assert)
		{
			let api = createApi(Api);
			let clipboard = AscCommon.g_clipboardBase;
			let oldDiv = clipboard.CommonDiv, oldIframe = clipboard.CommonIframe;
			let div = {}, iframe = {};
			clipboard.CommonDiv = div;
			clipboard.CommonIframe = iframe;
			try
			{
				assert.strictEqual(api.isClipboardPasteElement(div, undefined), true, "the div of the clipboard");
				assert.strictEqual(api.isClipboardPasteElement({}, iframe), true, "the iframe of the clipboard");
				assert.strictEqual(api.isClipboardPasteElement({}, undefined), false, "another element, a plugin");
				assert.strictEqual(api.isClipboardPasteElement("text", null), false, "text");
			}
			finally
			{
				clipboard.CommonDiv = oldDiv;
				clipboard.CommonIframe = oldIframe;
			}
		});
	});
});
