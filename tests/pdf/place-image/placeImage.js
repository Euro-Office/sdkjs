/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

$(function () {

	// a PDF viewer without UI and fonts, with two empty A4 pages (595 x 842 pt)
	AscCommon.CViewer.prototype.createComponents = function(){};
	AscCommon.CViewer.prototype.resize = function(){};
	AscCommon.CViewer.prototype._paintForms = function(){};
	AscCommon.CViewer.prototype._paintFormsHighlight = function(){};
	AscPDF.CPDFDoc.prototype.checkFonts = function(){ return true; };

	let viewer = new AscCommon.CViewer(null, AscTest.Editor);
	viewer.file = window["AscViewer"].createEmptyFile();
	viewer.file.pages = [
		{W: 595, H: 842, Dpi: 72, fonts: [], Rotate: 0},
		{W: 595, H: 842, Dpi: 72, fonts: [], Rotate: 0}
	];
	viewer.pagesInfo.setCount(2);

	let pdfDoc = viewer.getPDFDoc();
	AscTest.Editor.getDocumentRenderer = function() { return viewer; };
	AscTest.Editor.getPDFDoc = function() { return pdfDoc; };
	AscTest.Editor.sync_SearchEndCallback = function(){};

	// the PDF editor API working on that document; pictures "load" at once
	let api = Object.create(AscCommon.PDFEditorApi.prototype);
	api.getPDFDoc = function() { return pdfDoc; };
	api.getDocumentRenderer = function() { return viewer; };
	api.ImageLoader = {
		LoadImagesWithCallback : function(urls, callback) { callback(); },
		LoadImage : function(url) { return {src: url, Image: {width: 400, height: 100}}; }
	};

	function getDrawings(page)
	{
		return pdfDoc.GetPageInfo(page).drawings;
	}

	function clearPages()
	{
		[0, 1].forEach(function(page) {
			getDrawings(page).slice().forEach(function(drawing) {
				pdfDoc.GetPageInfo(page).RemoveDrawing(drawing.GetId());
			});
		});
	}

	QUnit.module("PDF: place a picture with the mouse");

	QUnit.test("placeImageAt puts the picture centred on the point, with the given width", function (assert)
	{
		clearPages();
		let picture = api.placeImageAt({src: "signature.png", Image: {width: 200, height: 100}}, {page: 0, x: 100, y: 50, w: 60});

		assert.ok(picture instanceof AscPDF.CPdfImage, "a picture is created");
		assert.strictEqual(getDrawings(0).length, 1, "it is on page 1");
		assert.strictEqual(getDrawings(0)[0], picture, "it is the placed picture");
		assert.strictEqual(getDrawings(1).length, 0, "page 2 stays empty");

		let xfrm = picture.getXfrm();
		assert.strictEqual(xfrm.extX, 60, "width 60 mm");
		assert.strictEqual(xfrm.extY, 30, "height keeps the side ratio (2:1)");
		assert.strictEqual(xfrm.offX, 70, "left edge = 100 - 60 / 2");
		assert.strictEqual(xfrm.offY, 35, "top edge = 50 - 30 / 2");
		assert.strictEqual(picture.getImageUrl(), "signature.png", "the picture shows the given image");
	});

	QUnit.test("AddImageUrl with placeAt places the picture on the clicked page", function (assert)
	{
		clearPages();
		api._addImageUrl(["sig.png"], {placeAt: {page: 1, x: 105, y: 200, w: 40}});

		assert.strictEqual(getDrawings(0).length, 0, "nothing on page 1");
		assert.strictEqual(getDrawings(1).length, 1, "the picture is on page 2");

		let xfrm = getDrawings(1)[0].getXfrm();
		assert.strictEqual(xfrm.extX, 40, "width 40 mm");
		assert.strictEqual(xfrm.extY, 10, "height from the loaded image (400 x 100)");
		assert.strictEqual(xfrm.offX, 85, "left edge = 105 - 40 / 2");
		assert.strictEqual(xfrm.offY, 195, "top edge = 200 - 10 / 2");
	});

	QUnit.test("placing is one undo step", function (assert)
	{
		clearPages();
		let history = pdfDoc.History;
		let points = history.Points.length;
		api.placeImageAt({src: "signature.png", Image: {width: 200, height: 100}}, {page: 0, x: 100, y: 50, w: 60});
		assert.strictEqual(history.Points.length, points + 1, "one new history point");
		assert.strictEqual(history.Points[history.Index].Description, AscDFH.historydescription_Presentation_AddFlowImage, "described as adding a picture");
	});

	QUnit.test("asc_StopPlaceImage without a running placement does nothing", function (assert)
	{
		api.placeImageState = null;
		api.asc_StopPlaceImage();
		assert.strictEqual(api.placeImageState, null, "no state, no error");
	});

	// pictures that load only when the test says so, and a count of the capturing listeners on the document
	function withSlowPictures(test)
	{
		let realImage = window.Image;
		let realAdd = document.addEventListener;
		let realRemove = document.removeEventListener;
		let pictures = [];
		let listeners = [];
		window.Image = function() {
			this.width = 400;
			this.height = 100;
			pictures.push(this);
		};
		document.addEventListener = function(type, handler, options) {
			if (options && options.capture)
				listeners.push(handler);
			return realAdd.apply(this, arguments);
		};
		document.removeEventListener = function(type, handler, options) {
			let index = listeners.indexOf(handler);
			if (options && options.capture && -1 !== index)
				listeners.splice(index, 1);
			return realRemove.apply(this, arguments);
		};
		try {
			test(pictures, listeners);
		}
		finally {
			api.asc_StopPlaceImage();
			window.Image = realImage;
			document.addEventListener = realAdd;
			document.removeEventListener = realRemove;
		}
	}

	function getGhosts(url)
	{
		return document.querySelectorAll('img[src="' + url + '"]').length;
	}

	function isWheelSwallowed()
	{
		let wheel = new WheelEvent("wheel", {deltaY: -1, bubbles: true, cancelable: true});
		document.body.dispatchEvent(wheel);
		return wheel.defaultPrevented;
	}

	QUnit.test("asc_StopPlaceImage before the picture has loaded cancels the placement", function (assert)
	{
		withSlowPictures(function(pictures, listeners) {
			api.asc_StartPlaceImage("slow-stop.png", 60);
			assert.ok(api.placeImageState, "the placement has started");
			api.asc_StopPlaceImage();
			assert.strictEqual(api.placeImageState, null, "stopped");

			pictures[0].onload();
			assert.strictEqual(api.placeImageState, null, "the late load does not start it again");
			assert.strictEqual(listeners.length, 0, "no listeners are left on the document");
			assert.strictEqual(getGhosts("slow-stop.png"), 0, "no ghost is left");
			assert.notOk(isWheelSwallowed(), "the mouse wheel scrolls again");
		});
	});

	QUnit.test("A second asc_StartPlaceImage before the first picture has loaded replaces it", function (assert)
	{
		withSlowPictures(function(pictures, listeners) {
			api.asc_StartPlaceImage("slow-first.png", 60);
			api.asc_StartPlaceImage("slow-second.png", 60);
			pictures[1].onload();
			pictures[0].onload();

			assert.strictEqual(getGhosts("slow-first.png"), 0, "no ghost of the first picture");
			assert.strictEqual(getGhosts("slow-second.png"), 1, "the ghost of the second picture");
			let count = listeners.length;
			assert.ok(count > 0, "the second placement listens to the mouse");
			assert.strictEqual(count, api.placeImageState.listeners.length, "only the listeners of the second placement");

			api.asc_StopPlaceImage();
			assert.strictEqual(listeners.length, 0, "stopping it removes all listeners");
			assert.strictEqual(getGhosts("slow-second.png"), 0, "and its ghost");
			assert.notOk(isWheelSwallowed(), "the mouse wheel scrolls again");
		});
	});

	QUnit.test("A picture that fails to load ends the placement", function (assert)
	{
		withSlowPictures(function(pictures, listeners) {
			api.asc_StartPlaceImage("broken.png", 60);
			pictures[0].onerror();
			assert.strictEqual(api.placeImageState, null, "nothing is waiting any more");
			assert.strictEqual(listeners.length, 0, "no listeners");
		});
	});
});
