## Third-party

- XRegExp 3.0.0 in `vendor/xregexp-all-min.js` ([MIT](https://raw.githubusercontent.com/slevithan/xregexp/master/LICENSE))

Vendored copies in `vendor/`, built into the editors. Not tracked by Renovate, update by hand:

- delta.js: Myers LCS from [node-delta](https://github.com/znerol/node-delta), version unknown ([MIT](https://github.com/znerol/node-delta/blob/master/LICENSE))
- easysax.js: [EasySAX](https://github.com/vflash/easysax), version unknown, heavily extended (StaxParser, XML reader/writer contexts) ([MIT](https://github.com/vflash/easysax/blob/master/LICENSE))
- iscroll.js: [iScroll](https://github.com/cubiq/iscroll) 5.2.0 ([MIT](https://github.com/cubiq/iscroll/blob/master/LICENSE))
- gifuct-js/: [gifuct-js](https://github.com/matt-way/gifuct-js), modified (IIFE, `GIFUCT` namespace) ([MIT](https://github.com/matt-way/gifuct-js/blob/master/LICENSE))

Derived code outside `vendor/`, shipped with the editors (`sha256.js` is bundled, `path-boolean-min.js` is copied to the deploy and loaded at runtime). The files carry only the Ascensio header, so the MIT notice lives here:

- `common/digest/sha256.js`: SHA-256 from [js-sha256](https://github.com/emn178/js-sha256), origin verified by code match, 0.9.0 by code markers, trimmed (no module wrapper, `getResult()` added). Copyright (c) 2014-2017 Chen, Yi-Cyuan ([MIT](https://github.com/emn178/js-sha256/blob/master/LICENSE.txt))
- `common/Drawings/Format/path-boolean.js` and its Closure output `path-boolean-min.js`: path boolean operations ported from [paper.js](https://github.com/paperjs/paper.js), origin verified by code match, 0.12.x (0.12.4 or later, exact version unknown), restructured into prototypes. Copyright (c) 2011-2020 Jürg Lehni & Jonathan Puckey ([MIT](https://github.com/paperjs/paper.js/blob/develop/LICENSE.txt))

### WebAssembly engines

Compiled from [Euro-Office/core](https://github.com/Euro-Office/core) with Emscripten and committed as `.wasm` plus JS glue. The libraries below are the third-party sources compiled in according to the build config named for each engine. They were not verified against the committed binaries, which may come from an older core build. Licences: [core 3DPARTY.md](https://github.com/Euro-Office/core/blob/main/3DPARTY.md); entries marked (\*) are not listed there yet (core lists zlib-ng and minizip-ng, not zlib 1.2.11).

- `pdf/src/engine/` (drawingfile, PDF/XPS/DjVu viewer), `DesktopEditor/graphics/pro/js/drawingfile.json`: FreeType 2.10.4, ICU, Crypto++, OpenJPEG, CxImage with its own libjpeg copy (core lists libjpeg-turbo), libpng, libtiff, JasPer (\*), libmng (\*), libpsd (\*) and libdcr (\*), jbig2enc with Leptonica, xpdf-derived PDF reader (\*), DjVuLibre, Anti-Grain Geometry 2.4 (\*), zlib 1.2.11 with minizip (\*), libxml2
- `common/libfont/engine/` (fonts), `DesktopEditor/fontengine/js/libfont.json`: FreeType 2.10.4, HarfBuzz, hyphen, Brotli (decoder), CxImage with libjpeg, libpng and libtiff, Anti-Grain Geometry 2.4 (\*), zlib 1.2.11 with minizip (\*)
- `common/zlib/engine/` (zlib), `OfficeUtils/js/zlib_image.json` (the committed binary contains libpng, libjpeg and libtiff strings, so not the plain `zlib.json`): zlib 1.2.11 with minizip (\*), CxImage with libjpeg, libpng and libtiff
- `common/spell/spell/` (spell), `Common/3dParty/hunspell/hunspell.json`: Hunspell
- `common/hash/hash/` (engine), `DesktopEditor/xmlsec/src/wasm/hash/hash.json`: OpenSSL (SHA, MD2, MD4, MD5, RIPEMD, Whirlpool)
