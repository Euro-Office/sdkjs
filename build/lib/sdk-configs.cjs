/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

/**
 * Shared SDK config loading — exact port of CConfig + getFilesMin/getFilesAll
 * from the original Gruntfile.js. Used by both sdk-concat-loader (webpack) and
 * build-develop.cjs so the two never diverge on how configs/file lists are built.
 */

'use strict';

const path = require('path');
const fs   = require('fs');

// Glob magic characters, not just *?{ — a pattern like foo[0-9].js is a glob too.
const MAGIC_CHARS = /[*?{}[\]!()]/;

// Files (no directories) matching pattern, as paths joined onto options.cwd if given.
function globFiles(pattern, options) {
    return fs.globSync(pattern, { ...options, withFileTypes: true })
        .filter((entry) => entry.isFile())
        .map((entry) => path.join(entry.parentPath, entry.name));
}

function loadJsonConfig(configsDir, name) {
    const file = path.join(configsDir, name + '.json');
    if (!fs.existsSync(file)) return null;
    try {
        return JSON.parse(fs.readFileSync(file, 'utf8'));
    } catch (e) {
        throw new Error(`sdk-configs: failed to parse ${file}: ${e.message}`);
    }
}

function fixPath(obj, basePath) {
    if (Array.isArray(obj)) {
        for (let i = 0; i < obj.length; i++) {
            obj[i] = path.join(basePath, obj[i]);
        }
        return;
    }
    for (const k of Object.keys(obj)) {
        fixPath(obj[k], basePath);
    }
}

function mergeConfigs(base, addon) {
    for (const k of Object.keys(addon)) {
        if (Array.isArray(addon[k])) {
            base[k] = Array.isArray(base[k]) ? base[k].concat(addon[k]) : addon[k];
        } else {
            if (!base[k]) base[k] = {};
            mergeConfigs(base[k], addon[k]);
        }
    }
}

function loadAllConfigs(srcRoot, addonDirs) {
    const configs = {};
    const configsDir = path.join(srcRoot, 'configs');

    for (const name of ['word', 'cell', 'slide', 'visio']) {
        const cfg = loadJsonConfig(configsDir, name);
        if (cfg) {
            fixPath(cfg, srcRoot);
            configs[name] = cfg;
        }
    }

    for (const addonDir of (addonDirs || [])) {
        for (const name of ['word', 'cell', 'slide', 'visio']) {
            if (!configs[name]) continue;
            const addon = loadJsonConfig(path.join(addonDir, 'configs'), name);
            if (!addon) continue;
            fixPath(addon, addonDir);
            mergeConfigs(configs[name], addon);
        }
    }

    return configs;
}

function getFilesMin(sdkCfg, platform) {
    let files = (sdkCfg['min'] || []).slice();
    if (platform === 'mobile' && sdkCfg['mobile_banners']) {
        files = (sdkCfg['mobile_banners']['min'] || []).concat(files);
    }
    if (platform === 'desktop' && sdkCfg['desktop']) {
        files = files.concat(sdkCfg['desktop']['min'] || []);
    }
    return files;
}

function getFilesAll(sdkCfg, platform) {
    let files = (sdkCfg['common'] || []).slice();
    if (platform === 'mobile') {
        if (sdkCfg['mobile_banners']) {
            files = (sdkCfg['mobile_banners']['common'] || []).concat(files);
        }
        const exclude = sdkCfg['exclude_mobile'] || [];
        files = files.filter(f => !exclude.includes(f));
        files = files.concat(sdkCfg['mobile'] || []);
    }
    if (platform === 'desktop' && sdkCfg['desktop']) {
        files = files.concat(sdkCfg['desktop']['common'] || []);
    }
    return files;
}

// Expand any glob patterns in a flat file list (already fixPath'd to absolute).
// Non-glob entries pass through unchanged; order of non-glob entries is preserved.
// Files matched by a glob are sorted (deterministic builds).
// addContextDep, if given, is called with each glob's parent directory so that
// consumers watching the filesystem (e.g. webpack) rebuild when a new file
// matching the pattern appears.
function expandGlobs(files, addContextDep) {
    const result = [];
    for (const f of files) {
        if (MAGIC_CHARS.test(f)) {
            const idx  = f.search(MAGIC_CHARS);
            const base = idx === -1 ? f : f.slice(0, idx);
            // When the magic character follows a slash (e.g. "common/Images/*"),
            // base already ends in a separator and IS the directory to watch —
            // path.dirname(base) would strip one path segment too many (watching
            // "common" instead of "common/Images"), so only fall back to
            // dirname() when base doesn't already end at a directory boundary.
            const dir  = /[\\/]$/.test(base)
                ? base.replace(/[\\/]+$/, '') || '.'
                : (path.dirname(base) || '.');
            if (addContextDep) addContextDep(dir);
            result.push(...globFiles(f).sort());
        } else {
            result.push(f);
        }
    }
    return result;
}

module.exports = {
    loadJsonConfig,
    fixPath,
    mergeConfigs,
    loadAllConfigs,
    getFilesMin,
    getFilesAll,
    expandGlobs,
    globFiles,
};
