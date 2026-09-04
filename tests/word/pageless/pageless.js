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

	const proto = AscCommon.DocumentEditorApi.prototype;

	// Builds a minimal api-shaped object that exercises the real SetPagelessMode/
	// GetPagelessMode implementations against an injected WordControl, recording the
	// refresh side effects (cache clear + layout re-fire) the setter is expected to trigger.
	function createApi() {
		const calls = { clearCache : 0, zoomFire : 0 };
		return {
			_calls : calls,
			SetPagelessMode : proto.SetPagelessMode,
			GetPagelessMode : proto.GetPagelessMode,
			WordControl : {
				isPagelessMode    : false,
				m_nZoomValue      : 100,
				m_oEditor         : { HtmlElement : { fullRepaint : false } },
				m_oDrawingDocument: { ClearCachePages : function () { calls.clearCache++; } },
				zoom_Fire         : function () { calls.zoomFire++; }
			}
		};
	}

	QUnit.module("Test pageless view mode (api)");

	QUnit.test("GetPagelessMode reflects the WordControl flag", function (assert) {
		let api = createApi();
		assert.strictEqual(api.GetPagelessMode(), false, "Off by default");
		api.WordControl.isPagelessMode = true;
		assert.strictEqual(api.GetPagelessMode(), true, "Reflects an enabled flag");
	});

	QUnit.test("GetPagelessMode is safe without a WordControl", function (assert) {
		assert.strictEqual(proto.GetPagelessMode.call({ WordControl : null }), false, "Returns false, does not throw");
	});

	QUnit.test("SetPagelessMode enables the flag and refreshes the layout", function (assert) {
		let api = createApi();
		api.SetPagelessMode(true);
		assert.strictEqual(api.WordControl.isPagelessMode, true, "Flag enabled");
		assert.strictEqual(api.GetPagelessMode(), true, "Getter agrees");
		assert.strictEqual(api.WordControl.m_oEditor.HtmlElement.fullRepaint, true, "Full repaint requested");
		assert.strictEqual(api._calls.clearCache, 1, "Page cache cleared (outline is cached)");
		assert.strictEqual(api._calls.zoomFire, 1, "Layout re-fired (collapses page gaps)");
	});

	QUnit.test("SetPagelessMode disables the flag again", function (assert) {
		let api = createApi();
		api.SetPagelessMode(true);
		api.SetPagelessMode(false);
		assert.strictEqual(api.GetPagelessMode(), false, "Flag disabled");
		assert.strictEqual(api._calls.zoomFire, 2, "Layout re-fired on each real change");
	});

	QUnit.test("SetPagelessMode is a no-op when the value is unchanged", function (assert) {
		let api = createApi();
		api.SetPagelessMode(true);
		api.SetPagelessMode(true);
		assert.strictEqual(api._calls.clearCache, 1, "No extra cache clear");
		assert.strictEqual(api._calls.zoomFire, 1, "No extra layout re-fire");
	});

	QUnit.test("SetPagelessMode coerces its argument to a boolean", function (assert) {
		let api = createApi();
		api.SetPagelessMode("yes");
		assert.strictEqual(api.WordControl.isPagelessMode, true, "Truthy value stored as true");
		api.SetPagelessMode(0);
		assert.strictEqual(api.WordControl.isPagelessMode, false, "Falsy value stored as false");
	});

	QUnit.test("SetPagelessMode is safe without a WordControl", function (assert) {
		proto.SetPagelessMode.call({ WordControl : null }, true);
		assert.ok(true, "Does not throw when WordControl is missing");
	});
});
