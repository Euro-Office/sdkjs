/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

/**
 * sdk-concat-loader
 *
 * Reads the ordered SDK JSON configs and returns ALL source files for a given
 * module/chunk as a single concatenated module. This is the only correct way to
 * bundle sdkjs under webpack: 69+ files across the word SDK alone use bare
 * top-level `var` declarations (no IIFE) that communicate across file boundaries
 * via concatenated scope. Putting all files into ONE webpack module preserves
 * that scope — every bare `var` is visible to every other file in the same chunk.
 *
 * Mirrors the CConfig + getFilesMin/getFilesAll logic from the original Gruntfile.js.
 *
 * Options (webpack loader options object):
 *   module    {string}   'word' | 'cell' | 'slide' | 'visio'           required
 *   chunk     {string}   'min' | 'all'                                  required
 *   platform  {string}   '' | 'desktop' | 'mobile'                      default ''
 *   srcRoot   {string}   absolute path to sdkjs root (one level above build/)
 *   addonDirs {string[]} absolute paths to addon directories
 *   buildMeta {object}   { companyName, version, buildNumber, beta } — patched into
 *                        window.AscCommon.g_cXxx after commonDefines.js ('min' chunk only)
 */

'use strict';

const path   = require('path');
const fs     = require('fs');
const { loadAllConfigs, getFilesMin, getFilesAll, expandGlobs } = require('../lib/sdk-configs.cjs');

// Lazy-load source-map-js (optional dep — absent means no source maps, build still works).
let SourceMapGenerator = null;
try {
    SourceMapGenerator = require('source-map-js').SourceMapGenerator;
} catch (_) {}

// Module-scope memo for loadAllConfigs(), keyed on its own inputs. Each webpack
// process (word/cell/slide/visio) runs 2 chunk configs ('min' + 'all'), and
// build-pipeline.cjs spawns 4 such processes in parallel — so without this,
// the same 4 module configs get re-read and re-JSON-parsed (plus a full
// recursive fixPath walk) 8 times across a build instead of once per process.
//
// Cached alongside each entry are the mtimes of every config file that fed it,
// so a long-lived process (npm run watch:*) re-parses once a config actually
// changes on disk instead of serving the pre-edit result forever. Without this,
// webpack correctly detects the addDependency() change and re-invokes the
// loader, but the loader itself would keep returning the stale parsed config.
const configsCache = new Map();

function configFiles(srcRoot, addonDirs) {
    const files = ['word', 'cell', 'slide', 'visio'].map(name =>
        path.join(srcRoot, 'configs', name + '.json'));
    for (const addonDir of addonDirs) {
        for (const name of ['word', 'cell', 'slide', 'visio']) {
            files.push(path.join(addonDir, 'configs', name + '.json'));
        }
    }
    return files;
}

function statMtimes(files) {
    const mtimes = {};
    for (const f of files) {
        try {
            mtimes[f] = fs.statSync(f).mtimeMs;
        } catch (_) {
            mtimes[f] = null; // missing file — still tracked, so it re-triggers if later created
        }
    }
    return mtimes;
}

function mtimesEqual(a, b) {
    const keys = Object.keys(a);
    if (keys.length !== Object.keys(b).length) return false;
    return keys.every(k => a[k] === b[k]);
}

function loadAllConfigsMemoized(srcRoot, addonDirs) {
    const key         = srcRoot + '|' + addonDirs.join(',');
    const currMtimes  = statMtimes(configFiles(srcRoot, addonDirs));
    const cached       = configsCache.get(key);

    if (cached && mtimesEqual(cached.mtimes, currMtimes)) {
        return cached.configs;
    }

    const configs = loadAllConfigs(srcRoot, addonDirs);
    configsCache.set(key, { configs, mtimes: currMtimes });
    return configs;
}

// ---------------------------------------------------------------------------
// Loader
// ---------------------------------------------------------------------------

module.exports = function sdkConcatLoader() {
    // this.resourcePath is dummy.js — its content is irrelevant; we ignore it.
    const callback  = this.async();
    const opts      = this.getOptions(module.exports.schema);
    const srcRoot   = path.resolve(opts.srcRoot   || path.join(this.context, '..'));
    const platform  = opts.platform  || '';
    const addonDirs = (opts.addonDirs || []).map(d => path.resolve(d));

    const configs = loadAllConfigsMemoized(srcRoot, addonDirs);
    const sdkCfg  = configs[opts.module] && configs[opts.module]['sdk'];

    if (!sdkCfg) {
        callback(new Error(`sdk-concat-loader: no config found for module "${opts.module}" at ${srcRoot}`));
        return;
    }

    const rawFiles = opts.chunk === 'min'
        ? getFilesMin(sdkCfg, platform)
        : getFilesAll(sdkCfg, platform);

    const addCtx = this.addContextDependency.bind(this);
    const files  = expandGlobs(rawFiles, addCtx);

    // Register every source file as a webpack dependency so watch mode works.
    for (const f of files) {
        this.addDependency(path.resolve(f));
    }
    // Watch the config file(s) so a config change triggers a rebuild — including
    // addon configs, which loadAllConfigs() merges in but which webpack would
    // otherwise never see, leaving --watch/persistent-cache builds stale.
    this.addDependency(path.join(srcRoot, 'configs', opts.module + '.json'));
    for (const addonDir of addonDirs) {
        this.addDependency(path.join(addonDir, 'configs', opts.module + '.json'));
    }

    // this.sourceMap reflects webpack's own devtool setting — building the
    // per-file map (and embedding every source's content) is wasted CPU/memory
    // when devtool:false, and this loader runs across 4 parallel webpack-cli
    // processes × 2 chunks each in the full pipeline.
    const needSourceMap = !!(this.sourceMap && SourceMapGenerator);

    // Runtime API polyfills (Promise, Map, WeakMap, Array.prototype.includes,
    // etc.) for whatever ES6+ built-ins the source actually uses — orthogonal to
    // syntax (let/const/arrow/classes), which ships as-authored now that no
    // downlevel target is applied. Consumers that load only sdk-all-min.js
    // (e.g. embed pages, which do not load sdkjs/vendor/polyfill.js nearby)
    // still need these built-ins on older browsers, so prepend the same
    // polyfill file the old dev-mode writeScripts() path already loads ahead
    // of sdk-all-min.js, keeping the 'min' chunk self-contained.
    let polyfillContent = '';
    if (opts.chunk === 'min') {
        const polyfillPath = path.join(srcRoot, 'vendor', 'polyfill.js');
        this.addDependency(polyfillPath);
        try {
            polyfillContent = fs.readFileSync(polyfillPath, 'utf8');
        } catch (err) {
            callback(new Error(`sdk-concat-loader: cannot read required polyfill file ${polyfillPath}: ${err.message}`));
            return;
        }
        if (!polyfillContent.endsWith('\n')) polyfillContent += '\n';
    }

    // Read all source files in parallel. A missing/unreadable file must fail
    // the build outright: silently substituting '' would concatenate a chunk
    // with a piece of its bare-var scope missing, producing a build that looks
    // green but throws a confusing ReferenceError deep inside the SDK at runtime.
    Promise.all(files.map(f =>
        fs.promises.readFile(f, 'utf8').catch(err => {
            throw new Error(`sdk-concat-loader: cannot read ${f}: ${err.message}`);
        })
    )).then(contents => {
        const isAll     = opts.chunk === 'all';
        const prefix    = isAll ? '(function(window, undefined) {\n' : '';
        const suffix    = isAll ? '\n})(window);' : '';

        // --- Build output + optional per-file source map ---
        const bundleName = opts.chunk === 'min' ? 'sdk-all-min.js' : 'sdk-all.js';
        const gen = needSourceMap ? new SourceMapGenerator({ file: bundleName }) : null;

        let result  = prefix + polyfillContent;
        // Generated line cursor: prefix occupies line 1 (the wrapper open), content starts at line 2.
        // polyfillContent (min chunk only) is opaque vendor code — not remapped — so it
        // just shifts where the first real source file's mappings begin.
        let genLine = (isAll ? 2 : 1) + (polyfillContent.match(/\n/g) || []).length;

        for (let i = 0; i < files.length; i++) {
            const content  = contents[i];
            const nlCount  = (content.match(/\n/g) || []).length;
            const hasTrail = content.endsWith('\n');

            if (gen) {
                // Files are concatenated verbatim (no per-file transform), so a
                // straight 1:1 line mapping is exact, not best-effort.
                const srcLines = content.split('\n');
                const mapped   = hasTrail ? srcLines.length - 1 : srcLines.length;
                for (let j = 0; j < mapped; j++) {
                    gen.addMapping({
                        generated: { line: genLine + j, column: 0 },
                        source:    files[i],
                        original:  { line: j + 1, column: 0 },
                    });
                }
                gen.setSourceContent(files[i], content);
            }

            // Leading ';' terminates the previous file's last statement (several
            // files end in `})(window)` without a semicolon) and keeps a file's
            // own top-level `"use strict";` from becoming the directive prologue
            // of the sdk-all wrapper, which would make the whole bundle strict.
            // Same line count, so per-file line mappings stay exact.
            result += ';' + content;

            if (hasTrail) {
                // Trailing \n already provides the inter-file separator.
                genLine += nlCount;
            } else {
                result  += '\n';
                genLine += nlCount + 1;
            }
        }

        // commonDefines.js (part of the 'min' chunk) hardcodes
        // window.AscCommon.g_cXxx to placeholder values — DefinePlugin above
        // can't touch that assignment's LHS (see the comment at its call site
        // in webpack.sdk.factory.mjs). Patch the real build metadata in right
        // after, so the runtime-visible globals match the folded call-sites.
        if (opts.chunk === 'min' && opts.buildMeta) {
            const { companyName, version, buildNumber, beta } = opts.buildMeta;
            result +=
                '\nwindow.AscCommon.g_cCompanyName = '    + JSON.stringify(companyName) + ';' +
                '\nwindow.AscCommon.g_cProductVersion = ' + JSON.stringify(version)     + ';' +
                '\nwindow.AscCommon.g_cBuildNumber = '    + JSON.stringify(buildNumber) + ';' +
                '\nwindow.AscCommon.g_cIsBeta = '         + JSON.stringify(beta)        + ';\n';
        }

        result += suffix;

        callback(null, result, gen ? gen.toJSON() : undefined);
    }).catch(callback);
};

module.exports.schema = {
    type: 'object',
    properties: {
        module:    { type: 'string', enum: ['word', 'cell', 'slide', 'visio'] },
        chunk:     { type: 'string', enum: ['min', 'all'] },
        platform:  { type: 'string', enum: ['', 'desktop', 'mobile'] },
        srcRoot:   { type: 'string' },
        addonDirs: { type: 'array', items: { type: 'string' } },
        buildMeta: {
            type: 'object',
            properties: {
                companyName: { type: 'string' },
                version:     { type: 'string' },
                buildNumber: { type: 'string' },
                beta:        { type: 'string' },
            },
        },
    },
    required: ['module', 'chunk'],
    additionalProperties: false,
};
