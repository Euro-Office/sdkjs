(function(){
	String.prototype.sentenceCase = function ()
	{
		var trimStr = this.trim();
		if (trimStr.length === 0)
		{
			return "";
		} else if (trimStr.length === 1)
		{
			return trimStr[0].toUpperCase();
		}
		return trimStr[0].toUpperCase() + trimStr.slice(1);
	}
	String.prototype['sentenceCase'] = String.prototype.sentenceCase;

	String.prototype.strongMatch = function (regExp)
	{
		if (regExp && regExp instanceof RegExp)
		{
			var arr = this.toString().match(regExp);
			return !!(arr && arr.length > 0 && arr[0].length == this.length);
		}

		return false;
	};

	function CUnicodeIteratorOld(str)
	{
		this._position = 0;
		this._index = 0;
		this._str = str;
	}

	CUnicodeIteratorOld.prototype.isOutside = function()
	{
		return (this._index >= this._str.length);
	};
	CUnicodeIteratorOld.prototype.isInside = function()
	{
		return (this._index < this._str.length);
	};
	CUnicodeIteratorOld.prototype.value = function()
	{
		if (this._index >= this._str.length)
			return 0;

		var nCharCode = this._str.charCodeAt(this._index);
		if (!AscCommon.isLeadingSurrogateChar(nCharCode))
			return nCharCode;

		if ((this._str.length - 1) === this._index)
			return nCharCode; // error

		var nTrailingChar = this._str.charCodeAt(this._index + 1);
		return AscCommon.decodeSurrogateChar(nCharCode, nTrailingChar);
	};
	CUnicodeIteratorOld.prototype.next = function()
	{
		if (this._index >= this._str.length)
			return;

		this._position++;
		if (!AscCommon.isLeadingSurrogateChar(this._str.charCodeAt(this._index)))
		{
			++this._index;
			return;
		}

		if (this._index === (this._str.length - 1))
		{
			++this._index;
			return;
		}

		this._index += 2;
	};
	CUnicodeIteratorOld.prototype.position = function()
	{
		return this._position;
	};
	CUnicodeIteratorOld.prototype.check = CUnicodeIteratorOld.prototype.isInside;

	var CUnicodeIteratorNew = null;

	if (String.prototype.codePointAt && (typeof Symbol === 'function') && typeof Symbol.iterator === "symbol")
	{
		CUnicodeIteratorNew = function(str)
		{
			this._position = 0;
			this._str = str;
			this._iterator = str[Symbol.iterator]();
			this._current = this._iterator.next();
		}

		CUnicodeIteratorNew.prototype.isOutside = function()
		{
			return this._current.done;
		};
		CUnicodeIteratorNew.prototype.isInside = function()
		{
			return !this._current.done;
		};
		CUnicodeIteratorNew.prototype.value = function()
		{
			if (this._current.done)
				return 0;

			return this._current.value.codePointAt(0);
		};
		CUnicodeIteratorNew.prototype.next = function()
		{
			if (this._current.done)
				return;

			this._position++;
			this._current = this._iterator.next();
		};
		CUnicodeIteratorNew.prototype.position = function()
		{
			return this._position;
		};
		CUnicodeIteratorNew.prototype.check = CUnicodeIteratorNew.prototype.isInside;
	}

	var CUnicodeIterator = CUnicodeIteratorNew ? CUnicodeIteratorNew : CUnicodeIteratorOld;

	/**
	 * @returns {CUnicodeIterator}
	 */
	String.prototype.getUnicodeIterator = function()
	{
		return new CUnicodeIterator(this);
	};
	/**
	 * @returns {number[]}
	 */
	String.prototype.codePointsArray = function(codePoints)
	{
		let _codePoints = codePoints ? codePoints : [];

		for (let iter = this.getUnicodeIterator(); iter.check(); iter.next())
			_codePoints.push(iter.value());

		return _codePoints;
	};

})();
