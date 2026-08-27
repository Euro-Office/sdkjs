#!/usr/bin/env node
/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

'use strict';

// Replaces the grunt copy-other and copy-standalone tasks.
//
// For each entry in the otherFiles list (mirrors Gruntfile.js):
//   - Non-JS files: plain fs.cp (recursive for directories, flat for globs)
//   - JS files: run through terser (WHITESPACE_ONLY equivalent) + license header
//
// JS files ignored (same ignoreFiles list as Gruntfile.js):
//   jquery_native, fonts_ie, spell_ie, engine_ie, zlib_ie, drawingfile_ie, themes

const path  = require('path');
const fs    = require('fs');
const { sync: globSync } = require('glob');
const { minify }   = require('terser');
const { resolveBuildRoot, buildLicenseHeader } = require('../lib/env.cjs');

const BUILD_DIR = path.resolve(__dirname, '..');
const SRC_ROOT  = path.resolve(BUILD_DIR, '..');

const BUILD_ROOT = resolveBuildRoot(BUILD_DIR);

const licenseText = buildLicenseHeader(BUILD_DIR, { stripSentinel: true });

// JS files skipped from individual minification (same as ignoreFiles in Gruntfile.js)
const IGNORE_NAMES = new Set([
    'jquery_native', 'fonts_ie', 'spell_ie', 'engine_ie',
    'zlib_ie', 'drawingfile_ie', 'themes',
]);

// Mirrors the otherFiles array in Gruntfile.js
const OTHER_FILES = [
    {
        cwd:  path.join(SRC_ROOT, 'vendor'),
        src:  ['polyfill.js'],
        dest: path.join(BUILD_ROOT, 'vendor'),
    },
    {
        cwd:  path.join(SRC_ROOT, 'common'),
        src:  [
            'device_scale.js',
            'Drawings/Format/path-boolean-min.js',
            'Charts/ChartStyles.js',
            'SmartArts/SmartArtData/*',
            'SmartArts/SmartArtDrawing/*',
            'Images/*',
            'Images/placeholders/*',
            'Images/content_controls/*',
            'Images/cursors/*',
            'Images/reporter/*',
            'Images/icons/*',
            'Native/*.js',
            'libfont/engine/*',
            'spell/spell/*',
            'hash/hash/*',
            'zlib/engine/*',
            'serviceworker/*',
        ],
        dest: path.join(BUILD_ROOT, 'common'),
    },
    {
        cwd:  path.join(SRC_ROOT, 'cell', 'css'),
        src:  ['*.css'],
        dest: path.join(BUILD_ROOT, 'cell', 'css'),
    },
    {
        cwd:  path.join(SRC_ROOT, 'slide', 'themes'),
        src:  ['**/**'],
        dest: path.join(BUILD_ROOT, 'slide', 'themes'),
    },
    {
        cwd:  path.join(SRC_ROOT, 'pdf'),
        src:  [
            'src/engine/*',
            'src/annotations/stamps/*.json',
        ],
        dest: path.join(BUILD_ROOT, 'pdf'),
    },
];

async function deployJsFile(srcPath, destPath) {
    const source  = fs.readFileSync(srcPath, 'utf8');
    const result  = await minify(source, {
        compress: false,
        mangle:   false,
        format:   {
            comments: false,
            // Same invariant as the sdk-all bundles in webpack.sdk.factory.mjs:
            // several of the files deployed here (Native/*.js, and the
            // libfont/engine/fonts_*.js that doctrenderer concatenates into the
            // script it compiles) are executed by a V8 built with
            // v8_enable_i18n_support=false, which rejects supplementary-plane
            // characters used as identifiers. These inputs happen to be
            // ASCII-clean today, so this is prophylactic rather than a fix --
            // but nothing else stops the next non-ASCII string landing in one of
            // them and reproducing issue #80 outside the bundle scan's reach.
            ascii_only: true,
            quote_keys: true,
        },
    });
    const content = licenseText + '\n' + (result.code != null ? result.code : source);
    fs.mkdirSync(path.dirname(destPath), { recursive: true });
    fs.writeFileSync(destPath, content, 'utf8');
}

function deployFile(srcPath, destPath) {
    fs.mkdirSync(path.dirname(destPath), { recursive: true });
    fs.copyFileSync(srcPath, destPath);
}

async function main() {
    const tasks = [];
    // Counts every deployed file, not just the async deployJsFile() tasks below —
    // tasks.length alone undercounts, since non-JS assets are copied synchronously
    // via deployFile() outside the tasks array. A wrong count here would mask a
    // real regression (e.g. a broken glob silently matching nothing).
    let fileCount = 0;

    for (const entry of OTHER_FILES) {
        const matches = [];
        for (const pattern of entry.src) {
            const found = globSync(pattern, { cwd: entry.cwd, nodir: true });
            for (const f of found) matches.push(f);
        }

        for (const relFile of matches) {
            const ext      = path.extname(relFile);
            const baseName = path.parse(relFile).name;
            const srcPath  = path.join(entry.cwd, relFile);
            const destPath = path.join(entry.dest, relFile);

            fileCount++;
            if (ext === '.js' && !IGNORE_NAMES.has(baseName)) {
                tasks.push(deployJsFile(srcPath, destPath));
            } else {
                deployFile(srcPath, destPath);
            }
        }
    }

    await Promise.all(tasks);
    process.stdout.write(`deploy-assets: ${fileCount} files deployed to ${BUILD_ROOT}\n`);
}

main().catch(err => {
    process.stderr.write(`deploy-assets FAILED: ${err.message}\n`);
    process.exit(1);
});
