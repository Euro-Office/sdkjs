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

	// =====================================================================
	// AscCommonExcel.isColorAutomatic
	// =====================================================================
	QUnit.module('isColorAutomatic');

	QUnit.test('null/undefined color is automatic', function (assert) {
		assert.strictEqual(AscCommonExcel.isColorAutomatic(null), true, 'null');
		assert.strictEqual(AscCommonExcel.isColorAutomatic(undefined), true, 'undefined');
	});

	QUnit.test('RgbColor with isAutoColor=true is automatic, regardless of rgb value', function (assert) {
		var autoBlack = AscCommonExcel.createRgbColor(0, 0, 0);
		autoBlack.isAutoColor = true;
		assert.strictEqual(AscCommonExcel.isColorAutomatic(autoBlack), true, 'auto-flagged black');

		var autoWhite = AscCommonExcel.createRgbColor(255, 255, 255);
		autoWhite.isAutoColor = true;
		assert.strictEqual(AscCommonExcel.isColorAutomatic(autoWhite), true, 'auto-flagged white');
	});

	QUnit.test('explicit RgbColor(0,0,0) is NOT automatic - literal black is not the auto default', function (assert) {
		var explicitBlack = AscCommonExcel.createRgbColor(0, 0, 0);
		assert.strictEqual(explicitBlack.isAutoColor, false, 'isAutoColor defaults to false on a fresh RgbColor');
		assert.strictEqual(AscCommonExcel.isColorAutomatic(explicitBlack), false,
			'explicit literal black must stay explicit, not be treated as automatic');
	});

	QUnit.test('RgbColor.clone() preserves isAutoColor', function (assert) {
		var auto = AscCommonExcel.createRgbColor(12, 34, 56);
		auto.isAutoColor = true;
		var clonedAuto = auto.clone();
		assert.strictEqual(clonedAuto.isAutoColor, true, 'cloned auto color stays automatic');
		assert.strictEqual(AscCommonExcel.isColorAutomatic(clonedAuto), true, 'isColorAutomatic agrees on the clone');

		var explicit = AscCommonExcel.createRgbColor(12, 34, 56);
		var clonedExplicit = explicit.clone();
		assert.strictEqual(clonedExplicit.isAutoColor, false, 'cloned explicit color stays explicit');
		assert.strictEqual(AscCommonExcel.isColorAutomatic(clonedExplicit), false, 'isColorAutomatic agrees on the clone');
	});

	QUnit.test('ThemeColor identity against g_oDefaultFormat.Font.c survives .clone()', function (assert) {
		var defaultFormat = AscCommonExcel.g_oDefaultFormat;
		var savedFont = defaultFormat.Font;

		try {
			var defaultFontColor = new AscCommonExcel.ThemeColor();
			defaultFormat.Font = {c: defaultFontColor};

			assert.strictEqual(AscCommonExcel.isColorAutomatic(defaultFontColor), true,
				'the workbook default font color is automatic by identity, not by a flag');

			// ThemeColor.clone() intentionally returns `this` (see WorkbookElems.js) - this is
			// the mechanism isColorAutomatic's identity check relies on to survive cloning
			// through charProperties. If a future change makes clone() return a real copy,
			// this assertion is what should start failing.
			var cloned = defaultFontColor.clone();
			assert.strictEqual(cloned, defaultFontColor, 'ThemeColor.clone() returns the same instance');
			assert.strictEqual(AscCommonExcel.isColorAutomatic(cloned), true,
				'identity survives the clone, so the cloned color is still seen as automatic');

			var otherThemeColor = new AscCommonExcel.ThemeColor();
			assert.strictEqual(AscCommonExcel.isColorAutomatic(otherThemeColor), false,
				'a different ThemeColor instance (not the default font color) is not automatic');
		} finally {
			defaultFormat.Font = savedFont;
		}
	});

	// =====================================================================
	// AscCommonExcel / Asc.DrawingContext.prototype.getDarkModeCorrectedColor
	// =====================================================================
	QUnit.module('getDarkModeCorrectedColor');

	// getDarkModeCorrectedColor only reads _darkModeRgbCache/_darkModeColorShuttle off `this`,
	// so a real DrawingContext (which needs a canvas, fmgrGraphics and font to construct) isn't
	// required - a manufactured `this` matching that shape is enough to exercise the real method.
	function makeFakeDrawingContext() {
		return {
			_darkModeRgbCache: {},
			_darkModeColorShuttle: new AscCommon.CColor(0, 0, 0, 1)
		};
	}

	QUnit.test('returns the shared shuttle instance, not a new object', function (assert) {
		var ctx = makeFakeDrawingContext();
		var result = Asc.DrawingContext.prototype.getDarkModeCorrectedColor.call(ctx, 10, 20, 30, 1);
		assert.strictEqual(result, ctx._darkModeColorShuttle, 'shuttle instance is returned as-is');
	});

	QUnit.test('repeated calls with the same rgb return the same corrected value (cache hit)', function (assert) {
		var ctx = makeFakeDrawingContext();
		Asc.DrawingContext.prototype.getDarkModeCorrectedColor.call(ctx, 10, 20, 30, 1);
		var first = {r: ctx._darkModeColorShuttle.getR(), g: ctx._darkModeColorShuttle.getG(), b: ctx._darkModeColorShuttle.getB()};

		Asc.DrawingContext.prototype.getDarkModeCorrectedColor.call(ctx, 10, 20, 30, 1);
		var second = {r: ctx._darkModeColorShuttle.getR(), g: ctx._darkModeColorShuttle.getG(), b: ctx._darkModeColorShuttle.getB()};

		assert.deepEqual(second, first, 'same input rgb produces the same corrected rgb on a cache hit');
		assert.strictEqual(Object.keys(ctx._darkModeRgbCache).length, 1, 'only one cache entry was created for one distinct color');
	});

	QUnit.test('shuttle is overwritten (not stale) across consecutive different-color calls', function (assert) {
		var ctx = makeFakeDrawingContext();

		Asc.DrawingContext.prototype.getDarkModeCorrectedColor.call(ctx, 0, 0, 0, 1);
		var forBlack = {r: ctx._darkModeColorShuttle.getR(), g: ctx._darkModeColorShuttle.getG(), b: ctx._darkModeColorShuttle.getB()};

		Asc.DrawingContext.prototype.getDarkModeCorrectedColor.call(ctx, 255, 255, 255, 1);
		var forWhite = {r: ctx._darkModeColorShuttle.getR(), g: ctx._darkModeColorShuttle.getG(), b: ctx._darkModeColorShuttle.getB()};

		assert.notDeepEqual(forWhite, forBlack,
			'the shuttle reflects the second call\'s color, not a leftover from the first call');
		assert.strictEqual(Object.keys(ctx._darkModeRgbCache).length, 2, 'two distinct colors produced two cache entries');
	});

	QUnit.test('alpha is passed through unmodified, only rgb goes through correction', function (assert) {
		var ctx = makeFakeDrawingContext();
		Asc.DrawingContext.prototype.getDarkModeCorrectedColor.call(ctx, 100, 100, 100, 0.5);
		assert.strictEqual(ctx._darkModeColorShuttle.getA(), 0.5, 'alpha is carried through as given, not corrected');
	});

	// =====================================================================
	// WorksheetView.prototype._getKeepsAutomaticTextColorAsIs
	// =====================================================================
	QUnit.module('_getKeepsAutomaticTextColorAsIs');

	// Only this.handlers/this.workbook/this.settings.findFillColor and the passed-in cell's
	// Fill are touched, so a real WorksheetView (which needs a full editor boot to construct)
	// isn't required - a manufactured `this` matching that shape is enough.
	function makeFakeWorksheetView(opts) {
		opts = opts || {};
		return {
			handlers: {
				trigger: function (name) {
					return name === 'selectSearchingResults' ? !!opts.searchHighlightOn : false;
				}
			},
			workbook: {
				inFindResults: function () {
					return opts.isFindResult ? true : undefined;
				}
			},
			settings: {
				findFillColor: opts.findFillColor
			},
			// real method deliberately, not a fake: its only two deps are already modeled above; a copied reimplementation could silently drift from the real logic
			_isFindResultHighlighted: AscCommonExcel.WorksheetView.prototype._isFindResultHighlighted
		};
	}

	function callKeepsAsIs(wsOpts, fill, resolvedFallbackBg) {
		var ws = makeFakeWorksheetView(wsOpts);
		var cell = {getFill: function () { return fill; }};
		return AscCommonExcel.WorksheetView.prototype._getKeepsAutomaticTextColorAsIs.call(ws, cell, 0, 0, resolvedFallbackBg);
	}

	function makeSolidFill(r, g, b) {
		var fill = new AscCommonExcel.Fill();
		fill.fromColor(AscCommonExcel.createRgbColor(r, g, b));
		return fill;
	}

	function makeGradientFill() {
		var fill = new AscCommonExcel.Fill();
		fill.gradientFill = new AscCommonExcel.GradientFill();
		return fill;
	}

	QUnit.test('no fill at all: does not keep as-is (needs correction, same as bare canvas)', function (assert) {
		var noFill = new AscCommonExcel.Fill();
		assert.notOk(noFill.hasFill(), 'sanity check: this Fill really has no fill (hasFill() can be null, not just false)');
		assert.strictEqual(callKeepsAsIs({}, noFill), false);
	});

	QUnit.test('solid light fill keeps automatic text as-is', function (assert) {
		assert.strictEqual(callKeepsAsIs({}, makeSolidFill(255, 255, 255)), true, 'white fill');
		assert.strictEqual(callKeepsAsIs({}, makeSolidFill(255, 255, 0)), true, 'yellow fill');
	});

	QUnit.test('solid dark fill does NOT keep automatic text as-is', function (assert) {
		assert.strictEqual(callKeepsAsIs({}, makeSolidFill(0, 0, 0)), false, 'black fill');
		assert.strictEqual(callKeepsAsIs({}, makeSolidFill(10, 10, 10)), false, 'near-black fill');
	});

	QUnit.test('pattern/gradient fill without resolvedFallbackBg: conservative true', function (assert) {
		assert.strictEqual(callKeepsAsIs({}, makeGradientFill()), true,
			'unknown per-pixel contrast is exempted from correction by default');
	});

	QUnit.test('pattern/gradient fill WITH resolvedFallbackBg: resolves against it instead of the conservative default', function (assert) {
		var gradient = makeGradientFill();
		assert.strictEqual(callKeepsAsIs({}, gradient, new AscCommon.CColor(255, 255, 255)), true,
			'light resolvedFallbackBg overrides the conservative true with a true - but for the real reason (contrast), not by default');
		assert.strictEqual(callKeepsAsIs({}, gradient, new AscCommon.CColor(0, 0, 0)), false,
			'dark resolvedFallbackBg flips the result to false, proving the fallback is actually consulted, not ignored');
	});

	QUnit.test('search-highlighted cell overrides the fill-based result entirely', function (assert) {
		// dark fill would normally return false; a light findFillColor while highlighted
		// must still return true, proving the highlight branch short-circuits before the
		// fill is ever consulted (not that it coincidentally agrees with it)
		var darkFill = makeSolidFill(0, 0, 0);
		assert.strictEqual(
			callKeepsAsIs({searchHighlightOn: true, isFindResult: true, findFillColor: new AscCommon.CColor(255, 255, 0)}, darkFill),
			true,
			'light search-highlight color wins over the cell\'s own dark fill'
		);

		// light fill would normally return true; a dark findFillColor while highlighted must
		// still return false, for the same reason in the opposite direction
		var lightFill = makeSolidFill(255, 255, 255);
		assert.strictEqual(
			callKeepsAsIs({searchHighlightOn: true, isFindResult: true, findFillColor: new AscCommon.CColor(0, 0, 0)}, lightFill),
			false,
			'dark search-highlight color wins over the cell\'s own light fill'
		);
	});

	QUnit.test('selectSearchingResults off, or cell not actually a find result: falls through to the fill', function (assert) {
		var lightFill = makeSolidFill(255, 255, 255);
		assert.strictEqual(
			callKeepsAsIs({searchHighlightOn: false, isFindResult: true, findFillColor: new AscCommon.CColor(0, 0, 0)}, lightFill),
			true,
			'search highlighting disabled entirely: falls through to the fill-based result'
		);
		assert.strictEqual(
			callKeepsAsIs({searchHighlightOn: true, isFindResult: false, findFillColor: new AscCommon.CColor(0, 0, 0)}, lightFill),
			true,
			'highlighting is on but this specific cell is not a find result: falls through to the fill-based result'
		);
	});

	// =====================================================================
	// AscCommonExcel.drawFillCell - regression pin for the opt-in guard (point 1 of the
	// PR #70 review) and for the exact bug class already fixed once in this branch's
	// history (commit 42705a02f4, conditional-formatting Data Bar colors)
	// =====================================================================
	QUnit.module('drawFillCell dark-mode guard');

	// setFillStyle/fillRect are the only ctx methods drawFillCell calls for a solid fill, so a
	// manufactured ctx recording what it was asked to draw is enough - no real canvas needed.
	function makeFakeFillCtx(isDarkMode) {
		var lastFillColor = null;
		return {
			isDarkMode: !!isDarkMode,
			_darkModeRgbCache: {},
			_darkModeColorShuttle: new AscCommon.CColor(0, 0, 0, 1),
			getDarkModeCorrectedColor: Asc.DrawingContext.prototype.getDarkModeCorrectedColor,
			setFillStyle: function (color) {
				lastFillColor = color;
				return this;
			},
			fillRect: function () {
			},
			getLastFillColor: function () {
				return lastFillColor;
			}
		};
	}

	function makeRect() {
		return new AscCommon.asc_CRect(0, 0, 10, 10);
	}

	QUnit.test('caller that omits the flag draws the literal color unmodified in dark mode (the fixed bug)', function (assert) {
		var ctx = makeFakeFillCtx(true);
		var fill = new AscCommonExcel.Fill();
		fill.fromColor(AscCommonExcel.createRgbColor(10, 20, 30));

		AscCommonExcel.drawFillCell(ctx, null, fill, makeRect());

		var drawn = ctx.getLastFillColor();
		assert.strictEqual(drawn.getR(), 10, 'red channel untouched');
		assert.strictEqual(drawn.getG(), 20, 'green channel untouched');
		assert.strictEqual(drawn.getB(), 30, 'blue channel untouched');
	});

	QUnit.test('caller that explicitly opts in gets the dark-mode-corrected color', function (assert) {
		var ctx = makeFakeFillCtx(true);
		var fill = new AscCommonExcel.Fill();
		fill.fromColor(AscCommonExcel.createRgbColor(0, 0, 0));

		AscCommonExcel.drawFillCell(ctx, null, fill, makeRect(), true);

		var drawn = ctx.getLastFillColor();
		assert.notStrictEqual(drawn.getR(), 0, 'literal black must have actually been corrected, not left as-is');
	});

	QUnit.test('outside dark mode, the flag has no effect either way', function (assert) {
		var ctx = makeFakeFillCtx(false);
		var fill = new AscCommonExcel.Fill();
		fill.fromColor(AscCommonExcel.createRgbColor(0, 0, 0));

		AscCommonExcel.drawFillCell(ctx, null, fill, makeRect(), true);

		var drawn = ctx.getLastFillColor();
		assert.strictEqual(drawn.getR(), 0, 'light mode never corrects, even when bIsFillRecolorable is true');
	});

	// =====================================================================
	// AscCommon.updateGlobalSkin - regression pin for the theme-corruption bug fixed in
	// commit 0f55894fb7 (interface theme switches were silently corrupting content dark
	// mode's cell background/grid color through a shared GlobalSkin/EditorSkins alias)
	// =====================================================================
	QUnit.module('interface theme switch must not corrupt cell background/grid colors');

	QUnit.test('canvas-cell-background/canvas-cell-grid no longer alias into GlobalSkin.CellBackground/CellGrid', function (assert) {
		var before = AscCommon.GlobalSkin;
		var savedCellBackground = before.CellBackground;
		var savedCellGrid = before.CellGrid;
		var savedBackground = before.Background;

		try {
			// simulates a web-apps interface-theme switch supplying CSS-custom-property-backed
			// values, including the two keys that used to alias straight into CellBackground/CellGrid
			AscCommon.updateGlobalSkin({
				type: 'light',
				'canvas-cell-background': '#010203',
				'canvas-cell-grid': '#040506',
				'canvas-cell-title-background': '#0a0b0c'
			});

			var after = AscCommon.GlobalSkin;
			assert.strictEqual(after.CellBackground, savedCellBackground,
				'interface theme switch must not overwrite content dark mode\'s cell background color');
			assert.strictEqual(after.CellGrid, savedCellGrid,
				'interface theme switch must not overwrite content dark mode\'s cell grid color');

			// sanity check the update mechanism itself still works for a color that IS still
			// coupled (Background <- canvas-cell-title-background), so the two assertions above
			// are proven by an actually-removed mapping, not by updateGlobalSkin silently no-oping
			assert.notStrictEqual(after.Background, savedBackground,
				'a genuinely-mapped color changes, proving updateGlobalSkin is not a global no-op');
		} finally {
			var skin = AscCommon.GlobalSkin;
			skin.CellBackground = savedCellBackground;
			skin.CellGrid = savedCellGrid;
			skin.Background = savedBackground;
			delete skin['canvas-cell-background'];
			delete skin['canvas-cell-grid'];
			delete skin['canvas-cell-title-background'];
			delete skin.type;
		}
	});

	// =====================================================================
	// Shared harness for the two "drive a real render pass" tests below
	// (Model-integrity round-trip, and Print-path isolation). Built the same
	// way as tests/cell/js-api/common.js: a real Workbook/WorkbookView/
	// WorksheetView, given a real DrawingContext so _drawCellsAndBorders et al.
	// have sane cols/rows/zoom instead of the empty state the stubbed
	// WorksheetView.prototype._init below would otherwise leave.
	//
	// Memoized rather than rebuilt per test: this boots a full spreadsheet_api/
	// Workbook/WorkbookView once, and AscCommon.History/g_oTableId/CollaborativeEditing
	// are page-wide singletons that assume exactly one such boot per page (a second
	// OpenDocumentFromBin trips over the first build's now-stale History.workbook and
	// leftover Index/Points state, e.g. CChangesTableIdAdd.WriteToBinary throwing on a
	// bootstrap-time TableId.Add that a fresh page would have no-op'd). Every test in
	// this file that calls this gets the same fixture instance either way.
	// =====================================================================
	var sharedDarkModeRenderFixture = null;
	function buildDarkModeRenderFixture(assert) {
		if (sharedDarkModeRenderFixture) {
			return sharedDarkModeRenderFixture;
		}
		// ---- REQUIRED ENVIRONMENT SETUP (same stubs as tests/cell/js-api/common.js,
		// minus WorksheetView.prototype.draw/getZoom/_getPPIX/_getPPIY - these tests need a
		// real render pass, so they give the WorksheetView a real DrawingContext below and
		// let those read real values off it instead of stubbing them away) ----
		var drawingDocument = {
			CanvasHit: null, CanvasHitContext: null,
			OnStartRecalculate: function () {}, OnRecalculatePage: function () {}, OnEndRecalculate: function () {},
			UpdateTargetTransform: function () {}, SelectEnabled: function () {}, SelectShow: function () {},
			TargetStart: function () {}, TargetShow: function () {}, TargetEnd: function () {}, showTarget: function () {},
			Set_RulerState_Start: function () {}, Set_RulerState_Paragraph: function () {}, Set_RulerState_End: function () {},
			Update_MathTrack: function () {}, startCollectContentControlTracks: function () {},
			endCollectContentControlTracks: function () {}, addContentControlTrack: function () {},
			removeContentControlTrackHover: function () {}, Update_FieldTrack: function () {},
			SetTargetColor: function () {}, SetTargetSize: function () {}, UpdateTarget: function () {},
			ClearCachePages: function () {}, OnRepaintPage: function () {}, FirePaint: function () {},
			GetMMPerDot: function (value) { return value / this.GetDotsPerMM(1); }, GetDotsPerMM: function () { return 72; },
			EndTrackTable: function () {}, SetCurrentPage: function () {}, SelectClear: function () {},
			Start_CollaborationEditing: function () {}, End_CollaborationEditing: function () {},
			ConvertCoordsToCursorWR: function () { return {X: 0, Y: 0}; }, Set_RulerState_Table: function () {},
			scrollToTarget: function () {}, GetVisibleRegion: function () { return [{Page: 0, Y: 0}, {Page: 0, Y: 0}]; }
		};
		drawingDocument.CanvasHit = document.createElement('canvas');
		drawingDocument.CanvasHitContext = drawingDocument.CanvasHit.getContext('2d');

		Asc.spreadsheet_api.prototype._init = function () {};
		Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) { callback(); };
		AscCommonExcel.Workbook.prototype._getSnapshot = function () { return null; };
		AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () {};
		AscCommonExcel.WorkbookView.prototype._init = function () {};
		AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {};
		AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {};
		AscCommonExcel.WorksheetView.prototype._init = function () {};
		AscCommonExcel.WorksheetView.prototype._onUpdateFormatTable = function () {};
		AscCommonExcel.WorksheetView.prototype.setSelection = function () {};
		AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {};
		AscCommonExcel.WorksheetView.prototype._reinitializeScroll = function () {};
		AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {};
		Asc.ReadDefTableStyles = function () {};

		var api = new Asc.spreadsheet_api({"id-view": "editor_sdk"});
		api.FontLoader = {LoadDocumentFonts: function () {}};
		window["Asc"]["editor"] = api;
		AscCommon.g_oTableId.init();
		// Real font-application init (name -> font-info dictionary lookup used by
		// DrawingContext.setFont, see cell/graphics/DrawingContext.js and
		// common/libfont/map.js CApplicationFonts.LoadFontWithoutEmbed). These tests
		// render for real, unlike tests/cell/js-api/common.js's harness, so this is
		// needed - without it, measuring/drawing any cell text throws. Same call as
		// tests/cell/spreadsheet-calculation/PrintTests.js's _onEndLoadSdk override.
		AscFonts.g_fontApplication.Init();
		// Actual glyph shaping/rasterization needs a real FreeType/HarfBuzz WASM module
		// (common/libfont/engine/fonts.js), loaded asynchronously and not present in this
		// QUnit page (develop/sdkjs/cell/scripts.js only lists the plain-JS loader/wrapper,
		// see common/libfont/engine.js). That module is orthogonal to what these tests check
		// (color/fill/CF model integrity and print-path isolation, not glyph metrics), so
		// DrawingContext's two FreeType-dependent entry points are given harmless real-shaped-object stand-ins
		// here instead: setFont skips the actual font-file load (CApplicationFonts /
		// CFontManagerEngine, which calls AscFonts.FT_CreateLibrary()), and measureText
		// returns placeholder-but-well-formed Asc.TextMetrics instead of going through
		// CFontManager.MeasureChar (which needs a loaded font file too). Every color
		// decision under test (isColorAutomatic, _getKeepsAutomaticTextColorAsIs,
		// getDarkModeCorrectedColor, drawFillCell, fill/stroke style calls) happens in real,
		// unstubbed code both before and independently of these two calls.
		Asc.DrawingContext.prototype.setFont = function (font, angle) {
			this.font.assign(font);
			this.setTextRotated(!!angle);
			return this;
		};
		Asc.DrawingContext.prototype.measureText = function (text, units, aCodes) {
			var len = aCodes ? aCodes.length : (text ? text.length : 0);
			var charWidth = 8;
			return new Asc.TextMetrics(len * charWidth, 12, 14, 10, 2, this.font.getSize(), len * charWidth);
		};
		Asc.DrawingContext.prototype.getFontMetrics = function (units) {
			var res = new Asc.FontMetrics();
			res.ascender = 10;
			res.descender = 2;
			res.lineGap = 2;
			res.nat_scale = 1;
			res.nat_y1 = 0;
			res.nat_y2 = 0;
			return res;
		};
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api.OpenDocumentFromBin(null, AscCommon.getEmpty());
		api.initCollaborativeEditing({});
		api.wbModel.DrawingDocument = drawingDocument;
		api.wbModel.mathTrackHandler = new AscWord.CMathTrackHandler(drawingDocument, api);
		api.wb = new AscCommonExcel.WorkbookView(
			api.wbModel,
			api.controller,
			api.handlers,
			api.HtmlElement,
			api.topLineEditorElement,
			api,
			api.collaborativeEditing,
			api.fontRenderingMode
		);

		var wsView = api.wb.getWorksheet(0);
		wsView.handlers = api.handlers;
		wsView.objectRender = new AscFormat.DrawingObjects();
		wsView.objectRender.OnUpdateOverlay = function () {};
		wsView.objectRender.drawingDocument = drawingDocument;
		wsView.objectRender.controller = new AscFormat.DrawingObjectsController(wsView.objectRender);

		// ---- Real geometry + a real DrawingContext, so the actual render path
		// (_drawCellsAndBorders et al.) has sane cols/rows/zoom instead of the empty
		// state left by the stubbed WorksheetView.prototype._init above. Built the same
		// way asc_initPrintPreview (cell/api.js) builds its own DrawingContext. ----
		var canvas = document.createElement('canvas');
		canvas.width = 400;
		canvas.height = 400;
		var drawingCtx = new Asc.DrawingContext({
			canvas: canvas, units: 0 /*px*/, fmgrGraphics: api.wb.fmgrGraphics, font: api.wb.m_oFont
		});
		wsView.buffers.main = drawingCtx;
		wsView.drawingCtx = drawingCtx;
		wsView.stringRender = new AscCommonExcel.StringRender(drawingCtx);

		var nCols = 2, nRows = 8;
		wsView.cellsLeft = 0;
		wsView.cellsTop = 0;
		wsView.cols = [];
		wsView.rows = [];
		for (var col = 0; col < nCols; ++col) {
			wsView.cols.push({left: col * 80, width: 80});
		}
		for (var row = 0; row < nRows; ++row) {
			wsView.rows.push({top: row * 20, height: 20, descender: 4});
		}
		wsView.updateColumnsStart = Number.MAX_VALUE;
		wsView.nColsCount = nCols;
		wsView.nRowsCount = nRows;
		var range = new Asc.Range(0, 0, nCols - 1, nRows - 1);
		wsView.visibleRange = range;
		// this.rows[i] above is a plain {top,height,descender} object, not a real RowInfo
		// instance (no .setHeight() method) - skip the row-auto-fit-to-content step that
		// would call it. Irrelevant to what these tests check (color/fill/CF model integrity
		// and print-path isolation).
		wsView.skipUpdateRowHeight = true;

		var jsApi = {};
		jsApi.GetActiveSheet = AscBuilder.Cell.Api.GetActiveSheet.bind(jsApi);
		jsApi.CreateColorFromRGB = AscBuilder.Cell.Api.CreateColorFromRGB.bind(jsApi);

		function color(r, g, b) {
			return jsApi.CreateColorFromRGB(r, g, b);
		}

		// ---- Fixture: a handful of cells covering the color-handling cases the review
		// flagged as untested against a real model ----
		var wsApi = jsApi.GetActiveSheet();
		var wsModel = wsView.model;

		// Case 1: explicit fill + explicit font color
		wsApi.GetRange("A1").SetValue(111);
		wsApi.GetRange("A1").SetFillColor(color(30, 60, 90));
		wsApi.GetRange("A1").SetFontColor(color(200, 120, 10));

		// Case 2: no fill, automatic font color (left untouched -> stays the workbook default)
		wsApi.GetRange("A2").SetValue(222);

		// Case 3: explicit (dark) fill + automatic font color - exercises
		// _getKeepsAutomaticTextColorAsIs (a dark background must not "keep as-is")
		wsApi.GetRange("A3").SetValue(333);
		wsApi.GetRange("A3").SetFillColor(color(10, 10, 10));

		// Case 4: a conditional-formatting data bar rule over a small range
		wsApi.GetRange("A4").SetValue(10);
		wsApi.GetRange("A5").SetValue(50);
		wsApi.GetRange("A6").SetValue(90);
		var dataBar = wsApi.GetRange("A4:A6").GetFormatConditions().AddDatabar();

		// Case 5: an automatic-colored border - the one case that actually exercises
		// drawBorder's isColorAutomatic gate (WorksheetView.js). An explicit-colored
		// border never reaches getDarkModeCorrectedColor at all, on screen or in print,
		// so it wouldn't prove anything about isolation either way; this has to be
		// automatic to be a meaningful check. Built via the low-level model API (not
		// the ApiRange.SetBorders builder) because CreateColorFromRGB always produces
		// an explicit color - there's no builder-level way to mark a border color
		// automatic, only isAutoColor set directly, same as the isColorAutomatic tests
		// above do for font colors.
		wsApi.GetRange("A7").SetValue(444);
		var autoBorderColor = AscCommonExcel.createRgbColor(0, 0, 0);
		autoBorderColor.isAutoColor = true;
		var autoBorderProp = new AscCommonExcel.BorderProp();
		autoBorderProp.setStyle(Asc.c_oAscBorderStyles.Thin);
		autoBorderProp.c = autoBorderColor;
		var autoBorder = new AscCommonExcel.Border();
		autoBorder.b = autoBorderProp;
		wsModel.getCell3(6, 0).setBorder(autoBorder);

		var cellRefs = {
			explicitFillAndFont: wsModel.getCell3(0, 0),
			noFillAutoFont: wsModel.getCell3(1, 0),
			darkFillAutoFont: wsModel.getCell3(2, 0),
			dataBarLow: wsModel.getCell3(3, 0),
			dataBarMid: wsModel.getCell3(4, 0),
			dataBarHigh: wsModel.getCell3(5, 0),
			autoBorderedCell: wsModel.getCell3(6, 0)
		};

		// The CDataBar rule element itself, resolved the same way the renderer resolves it
		// (WorksheetView.prototype._getCellCF / _drawCellCFDataBar)
		var cfRule = wsModel.getConditionalFormattingRangeIterator().get(4, 0)[0];
		var dataBarElement = cfRule && cfRule.asc_getColorScaleOrDataBarOrIconSetRule();
		// A solid (non-gradient) bar is a real, supported data bar style - it just renders
		// through drawFillCell's plain solid-color branch instead of the gradient branch,
		// which goes through the full DrawingML shape-drawing stack (AscFormat.CreateGeometry /
		// CShapeDrawer) via a "getMainGraphics" AscCommon.CGraphics that this minimal harness
		// has no reason to stand up: that stack is shape-rendering machinery, wholly unrelated
		// to what these tests check.
		if (dataBarElement) {
			dataBarElement.Gradient = false;
		}

		// ---- sanity: the fixture really is what the caller thinks it is, so a passing
		// check below isn't vacuously true because the setup above silently no-op'd ----
		assert.ok(!!dataBar, 'fixture sanity: AddDatabar() created a rule');
		assert.ok(cellRefs.explicitFillAndFont.getFill().hasFill(), 'fixture sanity: case 1 has an explicit fill');
		assert.notOk(AscCommonExcel.isColorAutomatic(cellRefs.explicitFillAndFont.getFont().getColor()),
			'fixture sanity: case 1 font color is explicit, not automatic');
		assert.notOk(cellRefs.noFillAutoFont.getFill().hasFill(), 'fixture sanity: case 2 has no fill');
		assert.ok(AscCommonExcel.isColorAutomatic(cellRefs.noFillAutoFont.getFont().getColor()),
			'fixture sanity: case 2 font color is automatic');
		assert.ok(cellRefs.darkFillAutoFont.getFill().hasFill(), 'fixture sanity: case 3 has an explicit fill');
		assert.ok(AscCommonExcel.isColorAutomatic(cellRefs.darkFillAutoFont.getFont().getColor()),
			'fixture sanity: case 3 font color is automatic (on a dark fill)');
		assert.ok(!!dataBarElement, 'fixture sanity: case 4 resolves to a data bar rule element via the same path the renderer uses');
		var case5Border = cellRefs.autoBorderedCell.getBorder();
		assert.ok(case5Border && case5Border.b, 'fixture sanity: case 5 has a bottom border');
		assert.ok(AscCommonExcel.isColorAutomatic(case5Border.b.c), 'fixture sanity: case 5 border color is automatic');

		sharedDarkModeRenderFixture = {
			api: api,
			wsView: wsView,
			drawingCtx: drawingCtx,
			range: range,
			wsModel: wsModel,
			cellRefs: cellRefs,
			dataBarElement: dataBarElement
		};
		return sharedDarkModeRenderFixture;
	}

	// =====================================================================
	// Model-integrity round-trip (PR #70 review, Critical risk): a dark-mode
	// render pass must not mutate the document model. Every other test in this
	// file drives dark-mode logic through hand-built fakes (makeFake*); this one
	// drives the real render entry point (WorksheetView.prototype._drawCellsAndBorders)
	// against a real Workbook/WorksheetView (buildDarkModeRenderFixture above), then
	// re-reads the cell/CF model afterwards to prove nothing changed.
	// =====================================================================
	QUnit.module('Model-integrity round-trip (dark-mode render must not mutate the model)');

	QUnit.test('rendering a fixture workbook in light and dark mode leaves every color-bearing model object unchanged', function (assert) {
		var fx = buildDarkModeRenderFixture(assert);
		var api = fx.api, wsView = fx.wsView, drawingCtx = fx.drawingCtx, range = fx.range,
			cellRefs = fx.cellRefs, dataBarElement = fx.dataBarElement;

		// ---- snapshot helpers ----
		function snapshotColor(c) {
			if (!c) {
				return {kind: 'none'};
			}
			if (c instanceof AscCommonExcel.ThemeColor) {
				return {kind: 'theme', isAutoColor: !!c.isAutoColor, theme: c.theme, tint: c.tint, rgb: c.rgb};
			}
			return {kind: 'rgb', isAutoColor: !!c.isAutoColor, r: c.getR(), g: c.getG(), b: c.getB()};
		}

		function snapshotCell(cell) {
			var fill = cell.getFill();
			var font = cell.getFont();
			var fontColor = font.getColor();
			return {
				fillRef: fill,
				fontRef: font,
				fontColorRef: fontColor,
				fillHasFill: fill.hasFill(),
				fillColorValue: snapshotColor(fill.getSolidFill()),
				fontColorValue: snapshotColor(fontColor)
			};
		}

		function snapshotDataBarElement(el) {
			if (!el) {
				return null;
			}
			return {
				ref: el,
				Color: snapshotColor(el.Color),
				NegativeColor: snapshotColor(el.NegativeColor),
				BorderColor: snapshotColor(el.BorderColor),
				NegativeBorderColor: snapshotColor(el.NegativeBorderColor),
				AxisColor: snapshotColor(el.AxisColor)
			};
		}

		function snapshotAll() {
			var snap = {dataBar: snapshotDataBarElement(dataBarElement)};
			for (var key in cellRefs) {
				snap[key] = snapshotCell(cellRefs[key]);
			}
			return snap;
		}

		function assertCellUnchanged(label, before, after) {
			assert.strictEqual(after.fillRef, before.fillRef, label + ': Fill object reference unchanged');
			assert.strictEqual(after.fontRef, before.fontRef, label + ': Font object reference unchanged');
			assert.strictEqual(after.fontColorRef, before.fontColorRef, label + ': font color object reference unchanged');
			assert.strictEqual(after.fillHasFill, before.fillHasFill, label + ': fill.hasFill() unchanged');
			assert.deepEqual(after.fillColorValue, before.fillColorValue, label + ': fill color value unchanged');
			assert.deepEqual(after.fontColorValue, before.fontColorValue, label + ': font color value unchanged');
		}

		function assertUnchanged(label, before, after) {
			for (var key in cellRefs) {
				assertCellUnchanged(label + ' - ' + key, before[key], after[key]);
			}
			if (before.dataBar) {
				assert.strictEqual(after.dataBar.ref, before.dataBar.ref, label + ' - data bar: rule element reference unchanged');
				assert.deepEqual(after.dataBar.Color, before.dataBar.Color, label + ' - data bar: Color unchanged');
				assert.deepEqual(after.dataBar.NegativeColor, before.dataBar.NegativeColor, label + ' - data bar: NegativeColor unchanged');
				assert.deepEqual(after.dataBar.BorderColor, before.dataBar.BorderColor, label + ' - data bar: BorderColor unchanged');
				assert.deepEqual(after.dataBar.NegativeBorderColor, before.dataBar.NegativeBorderColor, label + ' - data bar: NegativeBorderColor unchanged');
				assert.deepEqual(after.dataBar.AxisColor, before.dataBar.AxisColor, label + ' - data bar: AxisColor unchanged');
			}
		}

		var before = snapshotAll();

		// ---- 1. real render pass, light mode ----
		drawingCtx.isDarkMode = false;
		wsView._drawCellsAndBorders(drawingCtx, range, 0, 0);
		assertUnchanged('after light-mode render', before, snapshotAll());

		// ---- 2. real render pass, dark mode - the Critical risk from the review: a
		// dark-mode render pass must not mutate the document model ----
		api.wb.updateDarkMode(true);
		wsView._drawCellsAndBorders(drawingCtx, range, 0, 0);
		assertUnchanged('after dark-mode render', before, snapshotAll());
	});

	// =====================================================================
	// Print-path isolation (PR #70 review, High risk): automatic-color
	// correction is gated purely on ctx.isDarkMode, and WorkbookView.
	// updateDarkMode() only flips isDarkMode on buffers.main/overlay and
	// cellEditor's contexts - the print/PDF path (CPdfPrinter) is never
	// touched by it. This drives the same fixture cells through the real
	// print entry point (WorksheetView.prototype.drawForPrint) against a
	// real CPdfPrinter while the screen is toggled to dark mode, and checks
	// that print output is unaffected: identical to a light-mode-screen
	// print, and carrying the literal (not dark-corrected) fixture colors.
	// =====================================================================
	QUnit.module('Print-path isolation (dark mode must not leak into printed/PDF output)');

	QUnit.test('printing the fixture workbook through the real print path ignores screen dark mode entirely', function (assert) {
		var fx = buildDarkModeRenderFixture(assert);

		// CPdfPrinter's constructor reads window.Asc.editor.asc_getZoom() -> wb.getZoom() ->
		// wb.drawingCtx.getZoom(). WorkbookView.prototype.drawingCtx is only ever assigned
		// inside the real _init (this.drawingCtx = this.buffers.main), which
		// buildDarkModeRenderFixture stubs to a no-op to skip a full canvas/DOM boot, so
		// api.wb.drawingCtx would otherwise stay undefined and the construction below would
		// throw on undefined.getZoom(). Backfill it with the same real DrawingContext already
		// built for screen rendering - this is test-only, backfilling one side effect of the
		// stubbed _init, not a production bug (in real usage _init always runs for real).
		fx.api.wb.drawingCtx = fx.drawingCtx;

		// A page covering the whole fixture (2 cols * 80px, 8 rows * 20px). No headings/
		// gridlines, so the print pass stays scoped to _drawRowBG's fill portion and
		// _drawCellsBorders, the same surface the review flagged, without also pulling in
		// column/row header text.
		function makeFixturePage() {
			var page = new AscCommonExcel.CPagePrint();
			page.indexWorksheet = 0;
			page.pageRange = fx.range.clone();
			page.scale = 1;
			page.leftFieldInPx = 0;
			page.topFieldInPx = 0;
			page.pageGridLines = false;
			page.pageHeadings = false;
			page.pageClipRectLeft = 0;
			page.pageClipRectTop = 0;
			page.pageClipRectWidth = 160; // 2 cols * 80px
			page.pageClipRectHeight = 160; // 8 rows * 20px
			page.pageWidth = page.pageClipRectWidth * 25.4 / 96; // px -> mm, same koef as pdfprinter.js
			page.pageHeight = page.pageClipRectHeight * 25.4 / 96;
			return page;
		}

		// Drives one real print pass through a fresh CPdfPrinter (built the same way
		// WorkbookView.prototype.printSheets builds its own: fmgrGraphics[3] +
		// m_oFont.clone()), recording every color actually handed to setFillStyle/
		// setStrokeStyle by wrapping the real methods (so the real CDocumentRenderer
		// recording underneath still runs unmodified).
		function printAndRecord(isDarkModeOnScreen) {
			fx.api.wb.updateDarkMode(isDarkModeOnScreen);

			var pdfPrinter = new AscCommonExcel.CPdfPrinter(fx.api.wb.fmgrGraphics[3], fx.api.wb.m_oFont.clone());
			var fillCalls = [], strokeCalls = [];

			// CPdfPrinter.setFont -> CDocumentRenderer.SetFont -> CMetafileFontPicker ->
			// CFontFilesCache actually loads a font file through CFontManagerEngine
			// (common/libfont/manager.js), which needs the same real FreeType/HarfBuzz WASM
			// module noted in buildDarkModeRenderFixture above - not present on this QUnit
			// page. Unlike the screen DrawingContext's setFont (stubbed once, globally, up
			// there), this is the print ctx's own separate implementation, and drawForPrint's
			// _setDefaultFont calls it unconditionally before any cell is drawn. Orthogonal to
			// what this test checks (fill/border color isolation, not font/glyph loading), so
			// it's a no-op here, scoped to this one instance only.
			pdfPrinter.setFont = function () { return this; };

			var origSetFillStyle = pdfPrinter.setFillStyle;
			pdfPrinter.setFillStyle = function (val) {
				fillCalls.push({r: val.getR(), g: val.getG(), b: val.getB(), a: val.getA()});
				return origSetFillStyle.call(this, val);
			};
			var origSetStrokeStyle = pdfPrinter.setStrokeStyle;
			pdfPrinter.setStrokeStyle = function (val) {
				strokeCalls.push({r: val.getR(), g: val.getG(), b: val.getB(), a: val.getA()});
				return origSetStrokeStyle.call(this, val);
			};

			var page = makeFixturePage();
			fx.wsView.drawForPrint(pdfPrinter, page, 0, [page]);

			return {pdfPrinter: pdfPrinter, fillCalls: fillCalls, strokeCalls: strokeCalls};
		}

		var light = printAndRecord(false);
		assert.notOk(light.pdfPrinter.isDarkMode, 'CPdfPrinter never gains an isDarkMode flag (screen in light mode)');
		assert.ok(light.fillCalls.length > 0, 'fixture sanity: the print pass actually issued fill calls to check');
		// CPdfPrinter has no getDarkModeCorrectedColor of its own (unlike Asc.DrawingContext) -
		// so even a future bug that set isDarkMode on it would fail loudly (drawFillCell/
		// drawBorder would throw calling a method that doesn't exist) rather than silently
		// producing a plausible-looking wrong color.
		assert.strictEqual(typeof light.pdfPrinter.getDarkModeCorrectedColor, 'undefined',
			'the print ctx has no dark-mode color-correction method to call in the first place');

		var dark = printAndRecord(true);
		assert.ok(fx.drawingCtx.isDarkMode, 'fixture sanity: the screen context really is in dark mode for this run');
		assert.notOk(dark.pdfPrinter.isDarkMode,
			'CPdfPrinter never gains an isDarkMode flag, even with the screen context in dark mode - the print path has no dark-mode concept to leak through');

		// the actual isolation proof: printing is identical whether or not the screen
		// happens to be in dark mode at the moment print is invoked
		assert.deepEqual(dark.fillCalls, light.fillCalls,
			'fill colors sent to the print renderer are identical regardless of screen dark mode');
		// Case 5's automatic-colored border (see buildDarkModeRenderFixture) makes this a real
		// check: _drawCellsBorders draws an actual stroke for it, so the comparison below is
		// meaningful, not two empty arrays trivially matching each other.
		assert.deepEqual(dark.strokeCalls, light.strokeCalls,
			'stroke colors sent to the print renderer are identical regardless of screen dark mode');

		// the absolute-value proof: those colors are the real, literal fixture ones, not
		// dark-mode-corrected - i.e. print isn't merely *consistently* wrong either way
		var explicitFill = light.fillCalls.filter(function (c) { return 30 === c.r && 60 === c.g && 90 === c.b; });
		assert.ok(explicitFill.length > 0, 'case 1 explicit fill (30,60,90) reaches the print renderer unmodified');

		var darkFill = light.fillCalls.filter(function (c) { return 10 === c.r && 10 === c.g && 10 === c.b; });
		assert.ok(darkFill.length > 0, 'case 3 explicit dark fill (10,10,10) reaches the print renderer unmodified, not dark-mode-corrected');

		// case 5's border color is automatic (isAutoColor=true) - the one case that would
		// actually go through getDarkModeCorrectedColor on screen (drawBorder's
		// isColorAutomatic gate). If print incorrectly inherited screen dark-mode
		// correction, this raw (0,0,0) would come out shifted instead of unmodified.
		var autoBorderStroke = light.strokeCalls.filter(function (c) { return 0 === c.r && 0 === c.g && 0 === c.b; });
		assert.ok(autoBorderStroke.length > 0, 'case 5 automatic border color (0,0,0) reaches the print renderer unmodified, not dark-mode-corrected');
	});
});
