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

"use strict";

(	/**
* @param {Window} window
* @param {undefined} undefined
*/
function (window, undefined) {

	// Browsers cap the size of an element (Firefox at about 17.9M px). Longer ranges, such as a sheet
	// with 1,048,576 rows, are mapped onto this many native pixels.
	var MAX_NATIVE_SCROLL = 10000000;

	/**
	 * @constructor
	 */
	function ScrollSettings() {
		this.screenW = -1;
		this.screenH = -1;
		this.contentH = 0;
		this.contentW = 0;

		// thumb and track of the native scrollbar
		this.scrollerHoverColor = "#cfcfcf";
		this.scrollBackgroundColor = "#f4f4f4";

		this.alwaysVisible = false;
		this.isVerticalScroll = true;
		this.isHorizontalScroll = false;
	}

	/**
	 * A native, browser-drawn scrollbar in the element elemID. The element scrolls a spacer that overflows
	 * it by the scroll range, positions stay in the caller's units.
	 * @constructor
	 */
	function ScrollObject( elemID, settings ) {
		this.settings = settings;
		this.eventListeners = [];

		this.scrollVCurrentY = 0;
		this.scrollHCurrentX = 0;
		this.maxScrollY = 0;
		this.maxScrollX = 0;
		// callers may raise these above maxScroll to let scrollBy run past the end (the cell editor adds rows there)
		this.maxScrollY2 = 0;
		this.maxScrollX2 = 0;

		this.mouseDown = false;
		this.scrollerMouseDown = false;
		this.disableCurrentScroll = false;
		this.reinit = false;
		this.skipEndEvent = false;
		this.isRtl = false;

		this.scale = 1; // caller units per native pixel
		this.nativeMax = 0;
		this.nativePos = 0; // last native position written or handled
		this.frameId = 0;
		this.mouseUpFrame = 0;

		this.holder = document.getElementById( elemID );
		// a new ScrollObject on the same element takes over its spacer and handlers
		if ( !this.holder.scrollSpacer )
			this.holder.scrollSpacer = this.holder.appendChild( document.createElement( "div" ) );
		this.spacer = this.holder.scrollSpacer;

		var that = this;
		this.onFrame = function () {
			that._onFrame();
		};
		// a mouseup outside the window or iframe never arrives: end the drag on blur or on a move without buttons
		this.onWindowMouseUp = function ( e ) {
			if ( e.type !== "mousemove" || e.buttons === 0 )
				that.evt_mouseup( e );
		};
		this.holder.onscroll = function () {
			that._requestFrame();
		};
		this.holder.onmousedown = function ( e ) {
			that.evt_mousedown( e );
		};

		this._setRange( settings );
		// the element may still be scrolled by an earlier ScrollObject
		this._writeNative( true );
	}

	ScrollObject.prototype._setRange = function ( settings ) {
		var s = this.settings;
		var maxY = settings.contentH - settings.screenH;
		var maxX = settings.contentW - settings.screenW;
		this.maxScrollY = this.maxScrollY2 = maxY > 0 ? maxY : 0;
		this.maxScrollX = this.maxScrollX2 = maxX > 0 ? maxX : 0;

		var max = s.isVerticalScroll ? this.maxScrollY : this.maxScrollX;
		var screen = s.isVerticalScroll ? settings.screenH : settings.screenW;
		var client = s.isVerticalScroll ? this.holder.clientHeight : this.holder.clientWidth;
		// in pixels of the element, so the thumb shows the visible part whatever the caller's units
		var px = screen > 0 && client > 0 ? max * client / screen : max;
		this.nativeMax = Math.min( px, MAX_NATIVE_SCROLL );
		this.scale = this.nativeMax > 0 ? max / this.nativeMax : 1;

		var overflow = s.alwaysVisible ? "scroll" : "auto";
		var extent = "calc(100% + " + this.nativeMax + "px)";
		var style = this.holder.style;
		style.overflowY = s.isVerticalScroll ? overflow : "hidden";
		style.overflowX = s.isHorizontalScroll ? overflow : "hidden";
		style.overscrollBehavior = "contain";
		style.scrollbarWidth = "thin";
		style.scrollbarColor = s.scrollerHoverColor + " " + s.scrollBackgroundColor;
		style.backgroundColor = s.scrollBackgroundColor;
		this.spacer.style.height = s.isVerticalScroll ? extent : "1px";
		this.spacer.style.width = s.isHorizontalScroll ? extent : "1px";
	};
	ScrollObject.prototype.Repos = function ( settings, bIsHorAttack, bIsVerAttack, pos, isAttack ) {
		var s = this.settings;
		var isChangeTheme = s.scrollBackgroundColor !== settings.scrollBackgroundColor || s.scrollerHoverColor !== settings.scrollerHoverColor;

		if ( isChangeTheme || isAttack ) {
			for ( var i in settings ) {
				s[i] = settings[i];
			}
		}

		if ( !isChangeTheme ) {
			var maxY = settings.contentH - settings.screenH;
			var maxX = settings.contentW - settings.screenW;
			if ( bIsVerAttack && undefined !== settings.contentH && (maxY > 0 ? maxY : 0) === this.maxScrollY )
				return;
			if ( bIsHorAttack && undefined !== settings.contentW && (maxX > 0 ? maxX : 0) === this.maxScrollX )
				return;
		}

		this._setRange( settings );

		this.reinit = true;
		if ( s.isVerticalScroll ) {
			if ( this.scrollVCurrentY > this.maxScrollY )
				this.scrollVCurrentY = this.maxScrollY;
			if ( pos != null )
				this.scrollByY( pos - this.scrollVCurrentY );
		}
		if ( s.isHorizontalScroll ) {
			if ( this.scrollHCurrentX > this.maxScrollX )
				this.scrollHCurrentX = this.maxScrollX;
			if ( pos != null )
				this.scrollByX( pos - this.scrollHCurrentX );
		}
		this.reinit = false;

		// written now and not on the next frame: the spacer changed and the element may have been hidden
		if ( !this.mouseDown )
			this._writeNative( true );
	};
	ScrollObject.prototype.setRightToLeft = function ( isRtl ) {
		isRtl = !!isRtl;
		if ( this.isRtl === isRtl )
			return;
		this.isRtl = isRtl;
		this.holder.style.direction = isRtl ? "rtl" : "";
		if ( !this.mouseDown )
			this._writeNative( true );
	};

	ScrollObject.prototype._getNative = function () {
		if ( this.settings.isVerticalScroll )
			return this.holder.scrollTop;
		// scrollLeft runs from 0 to negative values in right-to-left elements
		return this.isRtl ? -this.holder.scrollLeft : this.holder.scrollLeft;
	};
	ScrollObject.prototype._writeNative = function ( isForce ) {
		var pos = this.settings.isVerticalScroll ? this.scrollVCurrentY : this.scrollHCurrentX;
		var target = Math.max( 0, Math.min( pos / this.scale, this.nativeMax ) );
		if ( !isForce && Math.abs( target - this.nativePos ) < 1 )
			return;

		if ( this.settings.isVerticalScroll )
			this.holder.scrollTop = target;
		else
			this.holder.scrollLeft = this.isRtl ? -target : target;
		// read back the rounded, clamped value, so the scroll event of this write is not taken as user input
		this.nativePos = this._getNative();
	};
	ScrollObject.prototype._requestFrame = function () {
		if ( !this.frameId )
			this.frameId = requestAnimationFrame( this.onFrame );
	};
	// Native scroll events (drag, track click, wheel over the bar) and programmatic positions meet here once
	// per frame. A moved native position is user input, otherwise the native position follows the caller's.
	ScrollObject.prototype._onFrame = function () {
		this.frameId = 0;

		var native = this._getNative();
		if ( Math.abs( native - this.nativePos ) >= 1 ) {
			this.nativePos = native;
			this._onNativeScroll( native );
		}

		// while the thumb is dragged, the native position belongs to the user
		if ( !this.mouseDown )
			this._writeNative( false );
	};
	ScrollObject.prototype._onNativeScroll = function ( native ) {
		if ( this.settings.isVerticalScroll ) {
			var destY = this.nativeMax - native < 1 ? this.maxScrollY : native * this.scale;
			if ( this.mouseDown ) {
				var result = this._correctScrollV( destY );
				if ( result != null && result.isChange === true )
					destY = result.Pos;
			}
			this._scrollV( {}, destY );
		}
		else if ( this.settings.isHorizontalScroll ) {
			this._scrollH( {}, this.nativeMax - native < 1 ? this.maxScrollX : native * this.scale );
		}
	};

	ScrollObject.prototype._scrollV = function ( evt, pos, isAttack ) {
		if ( this.scrollVCurrentY !== pos || isAttack === true ) {
			var oldPos = this.scrollVCurrentY;
			this.scrollVCurrentY = pos;
			evt.scrollD = evt.scrollPositionY = pos;
			evt.maxScrollY = this.maxScrollY;
			this.handleEvents( "onscrollvertical", evt );
			if ( this.checkDisableCurrentScroll() )
				this.scrollVCurrentY = oldPos;
		}
		else if ( pos > 0 && !this.reinit && !this.skipEndEvent ) {
			evt.pos = pos;
			this.handleEvents( "onscrollVEnd", evt );
		}
	};
	ScrollObject.prototype._scrollH = function ( evt, pos ) {
		if ( this.scrollHCurrentX !== pos ) {
			this.scrollHCurrentX = pos;
			evt.scrollD = evt.scrollPositionX = pos;
			evt.maxScrollX = this.maxScrollX;
			this.handleEvents( "onscrollhorizontal", evt );
		}
		else if ( pos > 0 && !this.reinit && !this.skipEndEvent ) {
			evt.pos = pos;
			this.handleEvents( "onscrollHEnd", evt );
		}
	};
	ScrollObject.prototype._correctScrollV = function ( yPos ) {
		var events = this.eventListeners["oncorrectVerticalScroll"];
		if ( events && events.length === 1 )
			return events[0].handler.apply( this, [yPos] );
		return null;
	};
	ScrollObject.prototype._correctScrollByYDelta = function ( delta ) {
		var events = this.eventListeners["oncorrectVerticalScrollDelta"];
		if ( events && events.length === 1 )
			return events[0].handler.apply( this, [delta] );
		return null;
	};
	ScrollObject.prototype.checkDisableCurrentScroll = function () {
		var ret = this.disableCurrentScroll;
		this.disableCurrentScroll = false;
		return ret;
	};

	ScrollObject.prototype.scrollByY = function ( delta, isAttack ) {
		if ( !this.settings.isVerticalScroll )
			return;

		var result = this._correctScrollByYDelta( delta );
		if ( result != null && result.isChange === true )
			delta = result.Pos;

		var destY = this.scrollVCurrentY + delta;
		var isEnd = false;
		if ( destY < 0 ) {
			destY = 0;
		}
		else if ( destY > this.maxScrollY2 ) {
			this.handleEvents( "onscrollVEnd", destY - this.maxScrollY );
			isEnd = true;
			destY = this.maxScrollY2;
		}

		this.skipEndEvent = isEnd;
		this._scrollV( {}, destY, isAttack );
		this.skipEndEvent = false;
		this._requestFrame();
	};
	ScrollObject.prototype.scrollToY = function ( destY ) {
		if ( !this.settings.isVerticalScroll )
			return;
		this._scrollV( {}, destY );
		this._requestFrame();
	};
	ScrollObject.prototype.scrollByX = function ( delta ) {
		if ( !this.settings.isHorizontalScroll )
			return;

		var destX = this.scrollHCurrentX + delta;
		var isEnd = false;
		if ( destX < 0 ) {
			destX = 0;
		}
		else if ( destX > this.maxScrollX2 ) {
			this.handleEvents( "onscrollHEnd", destX - this.maxScrollX );
			isEnd = true;
			destX = this.maxScrollX2;
		}

		this.skipEndEvent = isEnd;
		this._scrollH( {}, destX );
		this.skipEndEvent = false;
		this._requestFrame();
	};
	ScrollObject.prototype.scrollToX = function ( destX ) {
		if ( !this.settings.isHorizontalScroll )
			return;
		this._scrollH( {}, destX );
		this._requestFrame();
	};
	ScrollObject.prototype.scrollTo = function ( destX, destY ) {
		this.scrollToX( destX );
		this.scrollToY( destY );
	};
	ScrollObject.prototype.scrollBy = function ( deltaX, deltaY ) {
		this.scrollByX( deltaX );
		this.scrollByY( deltaY );
	};

	ScrollObject.prototype.getCurScrolledX = function () {
		return this.scrollHCurrentX;
	};
	ScrollObject.prototype.getCurScrolledY = function () {
		return this.scrollVCurrentY;
	};
	ScrollObject.prototype.getMaxScrolledY = function () {
		return this.maxScrollY;
	};
	ScrollObject.prototype.getMaxScrolledX = function () {
		return this.maxScrollX;
	};
	ScrollObject.prototype.getIsLockedMouse = function () {
		return this.mouseDown;
	};

	/************************************************************************/
	/*events*/
	// A press on the element is a press on the scrollbar (thumb or track): the element shows nothing else
	ScrollObject.prototype.evt_mousedown = function ( e ) {
		if ( e.button !== 0 )
			return;
		cancelAnimationFrame( this.mouseUpFrame );
		this.mouseUpFrame = 0;
		this.mouseDown = this.scrollerMouseDown = true;
		this._listenWindow( true );
	};
	ScrollObject.prototype._listenWindow = function ( isOn ) {
		var types = ["mouseup", "mousemove", "blur"];
		for ( var i = 0; i < types.length; i++ ) {
			if ( isOn )
				window.addEventListener( types[i], this.onWindowMouseUp, true );
			else
				window.removeEventListener( types[i], this.onWindowMouseUp, true );
		}
	};
	ScrollObject.prototype.evt_mouseup = function ( e ) {
		if ( !this.mouseDown || this.mouseUpFrame )
			return;
		this._listenWindow( false );

		// Chrome applies the last drag step after mouseup: wait for it, so mouseup handlers (slide snapping)
		// see the final position
		var that = this;
		this.mouseUpFrame = requestAnimationFrame( function () {
			that.mouseUpFrame = requestAnimationFrame( function () {
				that.mouseUpFrame = 0;
				cancelAnimationFrame( that.frameId );
				that._onFrame();
				that.handleEvents( "onmouseup", e );
				that.mouseDown = that.scrollerMouseDown = false;
				that._requestFrame();
			} );
		} );
	};

	ScrollObject.prototype.bind = function ( typesStr, handler ) {
		var types = typesStr.split( " " );
		/*
		 * loop through types and attach event listeners to
		 * each one.  eg. "click mouseover.namespace mouseout"
		 * will create three event bindings
		 */
		for ( var n = 0; n < types.length; n++ ) {
			var type = types[n];
			var event = (type.indexOf( 'touch' ) == -1) ? 'on' + type : type;
			var parts = event.split( "." );
			var baseEvent = parts[0];
			var name = parts.length > 1 ? parts[1] : "";

			if ( !this.eventListeners[baseEvent] ) {
				this.eventListeners[baseEvent] = [];
			}

			this.eventListeners[baseEvent].push( {
				name:name,
				handler:handler
			} );
		}
	};
	ScrollObject.prototype.unbind = function ( typesStr ) {
		var types = typesStr.split( " " );

		for ( var n = 0; n < types.length; n++ ) {
			var type = types[n];
			var event = (type.indexOf( 'touch' ) == -1) ? 'on' + type : type;
			var parts = event.split( "." );
			var baseEvent = parts[0];

			if ( this.eventListeners[baseEvent] && parts.length > 1 ) {
				var name = parts[1];

				for ( var i = 0; i < this.eventListeners[baseEvent].length; i++ ) {
					if ( this.eventListeners[baseEvent][i].name == name ) {
						this.eventListeners[baseEvent].splice( i, 1 );
						if ( this.eventListeners[baseEvent].length === 0 ) {
							this.eventListeners[baseEvent] = undefined;
						}
						break;
					}
				}
			}
			else {
				this.eventListeners[baseEvent] = undefined;
			}
		}
	};
	ScrollObject.prototype.handleEvents = function ( eventType, evt ) {
		var events = this.eventListeners[eventType];
		if ( events ) {
			for ( var i = 0; i < events.length; i++ ) {
				events[i].handler.apply( this, [evt] );
			}
		}
	};

	//---------------------------------------------------------export---------------------------------------------------
	window["AscCommon"].ScrollSettings = ScrollSettings;
	window["AscCommon"].ScrollObject = ScrollObject;
})(window);
