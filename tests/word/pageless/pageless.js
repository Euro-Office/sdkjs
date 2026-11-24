/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
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

	QUnit.module("Test pageless view mode (layout)");

	function createWordControl(pageCount) {
		let wordControl = new AscCommonWord.CEditorPage({ isMobileVersion : false });
		let pages = [];
		for (let i = 0; i < pageCount; i++)
			pages.push({ height_mm : 297 });
		wordControl.m_oDrawingDocument = { m_arrPages : pages };
		wordControl.m_nZoomValue = 100;
		return wordControl;
	}

	QUnit.test("Page gaps collapse to 0 when pageless", function (assert) {
		let wordControl = createWordControl(0);
		assert.strictEqual(wordControl.getPageGapVer(), wordControl.viewBetweenPagesVer, "Vertical gap kept in page view");
		assert.strictEqual(wordControl.getPageGapHor(), wordControl.viewBetweenPagesHor, "Horizontal gap kept in page view");
		wordControl.isPagelessMode = true;
		assert.strictEqual(wordControl.getPageGapVer(), 0, "No vertical gap when pageless");
		assert.strictEqual(wordControl.getPageGapHor(), 0, "No horizontal gap when pageless");
	});

	QUnit.test("Scrolling to a later page skips the collapsed gaps", function (assert) {
		let wordControl = createWordControl(3);
		let withGaps = wordControl.GetVerticalScrollTo(0, 3);
		wordControl.isPagelessMode = true;
		let pageless = wordControl.GetVerticalScrollTo(0, 3);
		assert.strictEqual(withGaps - pageless, 3 * wordControl.viewBetweenPagesVer, "Scroll target no longer counts one gap per preceding page");
	});

	QUnit.test("Visible height does not subtract collapsed gaps", function (assert) {
		let wordControl = createWordControl(3);
		wordControl.m_oEditor = { HtmlElement : { height : 1000 * AscCommon.AscBrowser.retinaPixelRatio } };
		let drawingDocument = new AscCommon.CDrawingDocument();
		drawingDocument.m_oWordControl = wordControl;
		drawingDocument.m_lDrawingFirst = 0;
		drawingDocument.m_lDrawingEnd = 2;
		let withGaps = drawingDocument.GetVisibleMMHeight();
		wordControl.isPagelessMode = true;
		let pageless = drawingDocument.GetVisibleMMHeight();
		assert.strictEqual(pageless, 1000 * g_dKoef_pix_to_mm, "Whole viewport counts as document height at 100% zoom");
		assert.ok(pageless > withGaps, "Page view still subtracts the gaps between visible pages");
	});
});
